<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/Global.asp"    -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ===========================================
'	  O R C A M E N T O E D I T A . A S P
'     ===========================================
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

'	EXIBI��O DE BOT�ES DE PESQUISA DE CEP
	dim blnPesquisaCEPAntiga, blnPesquisaCEPNova
	
	blnPesquisaCEPAntiga = False
	blnPesquisaCEPNova = True


	dim s, usuario, orcamento_selecionado
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

	orcamento_selecionado = ucase(Trim(request("orcamento_selecionado")))
	if (orcamento_selecionado = "") then Response.Redirect("aviso.asp?id=" & ERR_ORCAMENTO_NAO_ESPECIFICADO)
	s = normaliza_num_orcamento(orcamento_selecionado)
	if s <> "" then orcamento_selecionado = s
	
	dim i, n, s_fabricante, s_produto, s_descricao, s_descricao_html, s_obs, s_qtde, s_preco_lista, s_desc_dado, s_vl_unitario
	dim s_preco_NF, m_total_NF, m_total_RA
	dim s_vl_TotalItem, m_TotalItem, m_TotalItemComRA, m_TotalDestePedido, m_TotalDestePedidoComRA, s_readonly, s_readonly_valor
	dim intIdx
	
'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	dim s_aux, s2, s3, s4, r_loja, r_cliente
	dim r_orcamento, v_item, alerta, msg_erro
	dim strDisabled
	alerta=""
	if Not le_orcamento(orcamento_selecionado, r_orcamento, msg_erro) then 
		alerta = msg_erro
	else
		if Not le_orcamento_item(orcamento_selecionado, v_item, msg_erro) then alerta = msg_erro
		end if

	if alerta = "" then
		if Not orcamento_calcula_total_NF_e_RA(orcamento_selecionado, m_total_NF, m_total_RA, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
		end if

	dim r_pedido
	if alerta = "" then
		if r_orcamento.st_orc_virou_pedido = 1 then
			if Not le_pedido(r_orcamento.pedido, r_pedido, msg_erro) then alerta = msg_erro
			end if
		end if

	dim strTextoIndicador
	dim r_orcamentista_e_indicador
	if alerta = "" then
		call le_orcamentista_e_indicador(r_orcamento.orcamentista, r_orcamentista_e_indicador, msg_erro)
		end if

	dim strFlagEndEntregaEditavel
	if r_orcamento.st_orcamento=ST_ORCAMENTO_CANCELADO then 
		strFlagEndEntregaEditavel = "N" 
	else 
		strFlagEndEntregaEditavel = "S"
		end if
		
	dim blnFormaPagtoBloqueado
	blnFormaPagtoBloqueado = False
	
	dim blnInstaladorInstalaBloqueado
	blnInstaladorInstalaBloqueado = False
	if r_orcamento.st_orc_virou_pedido = 1 then blnInstaladorInstalaBloqueado = True
	
	dim blnGarantiaIndicadorBloqueado
	blnGarantiaIndicadorBloqueado = False
	if r_orcamento.st_orc_virou_pedido = 1 then blnGarantiaIndicadorBloqueado = True


	dim blnUsarMemorizacaoCompletaEnderecos
	blnUsarMemorizacaoCompletaEnderecos = isActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos
    'Definido em 20/03/2020: para os pedidos criado antes da memoriza��o completa, vamos usar a tela anterior.
    'N�o queremos exigir que quem editar o pedido seja obrigado a preenhcer o CNPJ do endere�o sde entrega. Ent�o, para
    'um pedido criado sem a memoriza��o, ele continua sempre sem a memoriza��o.
    if r_orcamento.st_memorizacao_completa_enderecos = 0 then
        blnUsarMemorizacaoCompletaEnderecos  = false
        end if

    'procesamos acessar o cliente s� para saber se � PF ou PJ
	set r_cliente = New cl_CLIENTE
    dim xcliente_bd_resultado
	xcliente_bd_resultado = x_cliente_bd(r_orcamento.id_cliente, r_cliente)

	dim eh_cpf
	if len(r_cliente.cnpj_cpf)=11 then eh_cpf=True else eh_cpf=False

    'le as vari�veis da origem certa: ou do pedido ou do cliente, todas comecam com cliente__
    dim cliente__tipo, cliente__cnpj_cpf, cliente__rg, cliente__ie, cliente__nome
    dim cliente__endereco, cliente__endereco_numero, cliente__endereco_complemento, cliente__bairro, cliente__cidade, cliente__uf, cliente__cep
    dim cliente__tel_res, cliente__ddd_res, cliente__tel_com, cliente__ddd_com, cliente__ramal_com, cliente__tel_cel, cliente__ddd_cel
    dim cliente__tel_com_2, cliente__ddd_com_2, cliente__ramal_com_2, cliente__email, cliente__email_xml, cliente__icms, cliente__produtor_rural_status

    cliente__tipo = r_cliente.tipo
    cliente__cnpj_cpf = r_cliente.cnpj_cpf
	cliente__rg = r_cliente.rg
    cliente__ie = r_cliente.ie
    cliente__nome = r_cliente.nome
    cliente__endereco = r_cliente.endereco
    cliente__endereco_numero = r_cliente.endereco_numero
    cliente__endereco_complemento = r_cliente.endereco_complemento
    cliente__bairro = r_cliente.bairro
    cliente__cidade = r_cliente.cidade
    cliente__uf = r_cliente.uf
    cliente__cep = r_cliente.cep
    cliente__tel_res = r_cliente.tel_res
    cliente__ddd_res = r_cliente.ddd_res
    cliente__tel_com = r_cliente.tel_com
    cliente__ddd_com = r_cliente.ddd_com
    cliente__ramal_com = r_cliente.ramal_com
    cliente__tel_cel = r_cliente.tel_cel
    cliente__ddd_cel = r_cliente.ddd_cel
    cliente__tel_com_2 = r_cliente.tel_com_2
    cliente__ddd_com_2 = r_cliente.ddd_com_2
    cliente__ramal_com_2 = r_cliente.ramal_com_2
    cliente__email = r_cliente.email
	cliente__email_xml = r_cliente.email_xml
	cliente__icms = r_cliente.contribuinte_icms_status
	cliente__produtor_rural_status = r_cliente.produtor_rural_status
   
    if blnUsarMemorizacaoCompletaEnderecos and r_orcamento.st_memorizacao_completa_enderecos <> 0 then 
        cliente__tipo = r_orcamento.endereco_tipo_pessoa
        cliente__cnpj_cpf = r_orcamento.endereco_cnpj_cpf
	    cliente__rg = r_orcamento.endereco_rg
        cliente__ie = r_orcamento.endereco_ie
        cliente__nome = r_orcamento.endereco_nome
        cliente__endereco = r_orcamento.endereco_logradouro
        cliente__endereco_numero = r_orcamento.endereco_numero
        cliente__endereco_complemento = r_orcamento.endereco_complemento
        cliente__bairro = r_orcamento.endereco_bairro
        cliente__cidade = r_orcamento.endereco_cidade
        cliente__uf = r_orcamento.endereco_uf
        cliente__cep = r_orcamento.endereco_cep
        cliente__tel_res = r_orcamento.endereco_tel_res
        cliente__ddd_res = r_orcamento.endereco_ddd_res
        cliente__tel_com = r_orcamento.endereco_tel_com
        cliente__ddd_com = r_orcamento.endereco_ddd_com
        cliente__ramal_com = r_orcamento.endereco_ramal_com
        cliente__tel_cel = r_orcamento.endereco_tel_cel
        cliente__ddd_cel = r_orcamento.endereco_ddd_cel
        cliente__tel_com_2 = r_orcamento.endereco_tel_com_2
        cliente__ddd_com_2 = r_orcamento.endereco_ddd_com_2
        cliente__ramal_com_2 = r_orcamento.endereco_ramal_com_2
        cliente__email = r_orcamento.endereco_email
		cliente__email_xml = r_orcamento.endereco_email_xml
	    cliente__icms = r_orcamento.endereco_contribuinte_icms_status
		cliente__produtor_rural_status = r_orcamento.endereco_produtor_rural_status
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


<script src="<%=URL_FILE__JQUERY%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_I18N%>" Language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>
<script src="<%=URL_FILE__AJAX_JS%>" Language="JavaScript" Type="text/javascript"></script>
<script src="<%=URL_FILE__JANELACEP_JS%>" language="JavaScript" type="text/javascript"></script>

<script type="text/javascript">
    $(function() {
        var f;
        f = fORC;
	    $("#divAjaxRunning").css('filter', 'alpha(opacity=60)'); // TRANSPAR�NCIA NO IE8

	    $("#EndEtg_obs option[value='<%=r_orcamento.EndEtg_cod_justificativa%>']").attr("selected", true);
        
	    // VERIFICAR MUDAN�A NOS CAMPOS
	    f.Verifica_End_Entrega.value = f.EndEtg_endereco.value;
	    f.Verifica_num.value = f.EndEtg_endereco_numero.value;
	    f.Verifica_Cidade.value = f.EndEtg_cidade.value;
	    f.Verifica_UF.value = f.EndEtg_uf.value;
	    f.Verifica_CEP.value = f.EndEtg_cep.value;
	    f.Verifica_Justificativa.value = f.EndEtg_obs.value;

        $("#c_data_previsao_entrega").hUtilUI('datepicker_padrao');

        $("input[name = 'rb_etg_imediata']").change(function () {
            configuraCampoDataPrevisaoEntrega();
        });

        configuraCampoDataPrevisaoEntrega();

        trataProdutorRuralEndEtg_PF(null);
        trocarEndEtgTipoPessoa(null);
	});

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
		$("#divAjaxRunning").css("top", newTop);
	}

    function configuraCampoDataPrevisaoEntrega() {
        if ($("input[name='rb_etg_imediata']:checked").val() == '<%=COD_ETG_IMEDIATA_NAO%>') {
            $("#c_data_previsao_entrega").prop("readonly", false);
            $("#c_data_previsao_entrega").prop("disabled", false);
            $("#c_data_previsao_entrega").datepicker("enable");
        }
        else {
            $("#c_data_previsao_entrega").val("");
            $("#c_data_previsao_entrega").prop("readonly", true);
            $("#c_data_previsao_entrega").prop("disabled", true);
            $("#c_data_previsao_entrega").datepicker("disable");
        }
    }
</script>

<script language="JavaScript" type="text/javascript">
var objAjaxCustoFinancFornecConsultaPreco;
var fCepPopup;

function fPEDConcluir(s_pedido){
	window.status = "Aguarde ...";
	fPED.pedido_selecionado.value=s_pedido;
	fPED.submit(); 
}

function ProcessaSelecaoCEP(){};

function AbrePesquisaCepEndEtg(){
var f, strUrl;
	try
		{
	//  SE J� HOUVER UMA JANELA DE PESQUISA DE CEP ABERTA, GARANTE QUE ELA SER� FECHADA 
	// E UMA NOVA SER� CRIADA (EVITA PROBLEMAS C/ O 'WINDOW.OPENER')	
		fCepPopup=window.open("", "AjaxCepPesqPopup","status=1,toolbar=0,location=0,menubar=0,directories=0,resizable=1,scrollbars=1,width=5,height=5,left=0,top=0");
		fCepPopup.close();
		}
	catch (e) {
	 // NOP
		}
	f=fORC;
	ProcessaSelecaoCEP=TrataCepEnderecoEntrega;
	strUrl="../Global/AjaxCepPesqPopup.asp";
	if (trim(f.EndEtg_cep.value)!="") strUrl=strUrl+"?CepDefault="+trim(f.EndEtg_cep.value);
	fCepPopup=window.open(strUrl, "AjaxCepPesqPopup", "status=1,toolbar=0,location=0,menubar=0,directories=0,resizable=1,scrollbars=1,width=980,height=650,left=0,top=0");
	fCepPopup.focus();
}

function TrataCepEnderecoEntrega(strCep, strUF, strLocalidade, strBairro, strLogradouro, strEnderecoNumero, strEnderecoComplemento) {
var f;
	f=fORC;
	f.EndEtg_cep.value=cep_formata(strCep);
	f.EndEtg_uf.value=strUF;
	f.EndEtg_cidade.value=strLocalidade;
	f.EndEtg_bairro.value=strBairro;
	f.EndEtg_endereco.value=strLogradouro;
	f.EndEtg_endereco_numero.value=strEnderecoNumero;
	f.EndEtg_endereco_complemento.value=strEnderecoComplemento;
	f.EndEtg_endereco.focus();
	window.status="Conclu�do";
}

function processaFormaPagtoDefault() {
var f, i;
	f=fORC;
	
//  O pedido foi cadastrado j� com a nova pol�tica de custo financeiro por fornecedor?
	if (f.c_custoFinancFornecTipoParcelamento.value=="") return;
	if (f.FormaPagtoBloqueado.value=="<%=Cstr(True)%>") return;
	
	for (i=0; i<fORC.rb_forma_pagto.length; i++) {
		if (fORC.rb_forma_pagto[i].checked) {
			fORC.rb_forma_pagto[i].click();
			break;
			}
		}

	f.c_custoFinancFornecParcelamentoDescricao.value=descricaoCustoFinancFornecTipoParcelamento(f.c_custoFinancFornecTipoParcelamento.value);
	if (f.c_custoFinancFornecTipoParcelamento.value==COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__COM_ENTRADA) {
		f.c_custoFinancFornecParcelamentoDescricao.value += " (1+" + f.c_custoFinancFornecQtdeParcelas.value + ")";
		}
	else if (f.c_custoFinancFornecTipoParcelamento.value==COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__SEM_ENTRADA) {
		f.c_custoFinancFornecParcelamentoDescricao.value += " (0+" + f.c_custoFinancFornecQtdeParcelas.value + ")";
		}
}

function trataRespostaAjaxCustoFinancFornecSincronizaPrecos() {
var f, strResp, i, j, xmlDoc, oNodes;
var strFabricante,strProduto, strStatus, strPrecoLista, strMsgErro, strCodigoErro;
var percDesc,vlLista,vlVenda,strMsgErroAlert;
	f=fORC;

//  O pedido foi cadastrado j� com a nova pol�tica de custo financeiro por fornecedor?
	if (f.c_custoFinancFornecTipoParcelamento.value=="") return;
	if (f.FormaPagtoBloqueado.value=="<%=Cstr(True)%>") return;

	strMsgErroAlert="";
	if (objAjaxCustoFinancFornecConsultaPreco.readyState==AJAX_REQUEST_IS_COMPLETE) {
		strResp=objAjaxCustoFinancFornecConsultaPreco.responseText;
		if (strResp=="") {
			alert("Falha ao consultar o pre�o!!");
			window.status="Conclu�do";
			$("#divAjaxRunning").hide();
			return;
			}

		if (strResp!="") {
			try
				{
				xmlDoc=objAjaxCustoFinancFornecConsultaPreco.responseXML.documentElement;
				for (i=0; i < xmlDoc.getElementsByTagName("ItemConsulta").length; i++) {
				//  Fabricante
					oNodes=xmlDoc.getElementsByTagName("fabricante")[i];
					if (oNodes.childNodes.length > 0) strFabricante=oNodes.childNodes[0].nodeValue; else strFabricante="";
					if (strFabricante==null) strFabricante="";
				//  Produto
					oNodes=xmlDoc.getElementsByTagName("produto")[i];
					if (oNodes.childNodes.length > 0) strProduto=oNodes.childNodes[0].nodeValue; else strProduto="";
					if (strProduto==null) strProduto="";
				//  Status
					oNodes=xmlDoc.getElementsByTagName("status")[i];
					if (oNodes.childNodes.length > 0) strStatus=oNodes.childNodes[0].nodeValue; else strStatus="";
					if (strStatus==null) strStatus="";
					if (strStatus=="OK") {
					//  Pre�o
						oNodes=xmlDoc.getElementsByTagName("precoLista")[i];
						if (oNodes.childNodes.length > 0) strPrecoLista=oNodes.childNodes[0].nodeValue; else strPrecoLista="";
						if (strPrecoLista==null) strPrecoLista="";
					//  Atualiza o pre�o
						if (strPrecoLista=="") {
							alert("Falha na consulta do pre�o do produto " + strProduto + "!!\n" + strMsgErro);
							}
						else {
							for (j=0; j<f.c_fabricante.length; j++) {
								if ((f.c_fabricante[j].value==strFabricante)&&(f.c_produto[j].value==strProduto)) {
								//  Percorre o la�o at� o final para o caso do usu�rio ter digitado o mesmo produto em v�rias linhas
								//	(apesar de que isso n�o ser� aceito pelas consist�ncias que ser�o feitas).
									f.c_preco_lista[j].value=strPrecoLista;
									f.c_preco_lista[j].style.color="black";
									}
								}
							}
						}
					else {
					//  C�digo do Erro
						oNodes=xmlDoc.getElementsByTagName("codigo_erro")[i];
						if (oNodes.childNodes.length > 0) strCodigoErro=oNodes.childNodes[0].nodeValue; else strCodigoErro="";
						if (strCodigoErro==null) strCodigoErro="";
					//  Mensagem de Erro
						oNodes=xmlDoc.getElementsByTagName("msg_erro")[i];
						if (oNodes.childNodes.length > 0) strMsgErro=oNodes.childNodes[0].nodeValue; else strMsgErro="";
						if (strMsgErro==null) strMsgErro="";
						for (j=0; j<f.c_fabricante.length; j++) {
						//  Percorre o la�o at� o final para o caso do usu�rio ter digitado o mesmo produto em v�rias linhas
						//	(apesar de que isso n�o ser� aceito pelas consist�ncias que ser�o feitas).
							if ((f.c_fabricante[j].value==strFabricante)&&(f.c_produto[j].value==strProduto)) {
								f.c_preco_lista[j].style.color=COR_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__INEXISTENTE;
								}
							}
						if (strMsgErroAlert!="") strMsgErroAlert+="\n\n";
						strMsgErroAlert+="Falha ao consultar o pre�o do produto " + strProduto + "!!\n" + strMsgErro;
						}
					}
				}
			catch (e)
				{
				alert("Falha na consulta do pre�o!!\n"+e.message);
				}
			}
			
		if (strMsgErroAlert!="") alert(strMsgErroAlert);
		
		recalcula_total_todas_linhas(); 
		recalcula_RA();
			
		window.status="Conclu�do";
		$("#divAjaxRunning").hide();
		}
}

function recalculaCustoFinanceiroPrecoLista() {
var f, i, strListaProdutos, strUrl, strOpcaoFormaPagto;
	f=fORC;

//  O pedido foi cadastrado j� com a nova pol�tica de custo financeiro por fornecedor?
	if (f.c_custoFinancFornecTipoParcelamento.value=="") return;
	if (f.FormaPagtoBloqueado.value=="<%=Cstr(True)%>") return;

	objAjaxCustoFinancFornecConsultaPreco=GetXmlHttpObject();
	if (objAjaxCustoFinancFornecConsultaPreco==null) {
		alert("O browser N�O possui suporte ao AJAX!!");
		return;
		}
		
	strListaProdutos="";
	for (i=0; i<f.c_fabricante.length; i++) {
		if ((trim(f.c_fabricante[i].value)!="")&&(trim(f.c_produto[i].value)!="")) {
			if (strListaProdutos!="") strListaProdutos+=";";
			strListaProdutos += f.c_fabricante[i].value + "|" + f.c_produto[i].value;
			}
		}
	if (strListaProdutos=="") return;
	
//  Converte as op��es de forma de pagamento do pedido em uma op��o que possa tratada pela tabela de custo financeiro
	strOpcaoFormaPagto="";
	for (i=0; i<fORC.rb_forma_pagto.length; i++) {
		if (fORC.rb_forma_pagto[i].checked) {
			strOpcaoFormaPagto=f.rb_forma_pagto[i].value;
			break;
			}
		}
	if (strOpcaoFormaPagto=="") return;
	
	if (strOpcaoFormaPagto==COD_FORMA_PAGTO_A_VISTA) {
		f.c_custoFinancFornecTipoParcelamento.value=COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__A_VISTA;
		f.c_custoFinancFornecQtdeParcelas.value='0';
		}
	else if (strOpcaoFormaPagto==COD_FORMA_PAGTO_PARCELA_UNICA) {
		f.c_custoFinancFornecTipoParcelamento.value=COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__SEM_ENTRADA;
		f.c_custoFinancFornecQtdeParcelas.value='1';
		}
	else if (strOpcaoFormaPagto==COD_FORMA_PAGTO_PARCELADO_CARTAO) {
		f.c_custoFinancFornecTipoParcelamento.value=COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__SEM_ENTRADA;
		f.c_custoFinancFornecQtdeParcelas.value=f.c_pc_qtde.value;
		}
	else if (strOpcaoFormaPagto==COD_FORMA_PAGTO_PARCELADO_CARTAO_MAQUINETA) {
		f.c_custoFinancFornecTipoParcelamento.value=COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__SEM_ENTRADA;
		f.c_custoFinancFornecQtdeParcelas.value=f.c_pc_maquineta_qtde.value;
	}
	else if (strOpcaoFormaPagto==COD_FORMA_PAGTO_PARCELADO_COM_ENTRADA) {
		f.c_custoFinancFornecTipoParcelamento.value=COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__COM_ENTRADA;
		f.c_custoFinancFornecQtdeParcelas.value=f.c_pce_prestacao_qtde.value;
		}
	else if (strOpcaoFormaPagto==COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA) {
		f.c_custoFinancFornecTipoParcelamento.value=COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__SEM_ENTRADA;
		f.c_custoFinancFornecQtdeParcelas.value = (converte_numero(f.c_pse_demais_prest_qtde.value) + 1).toString();
		}
	else {
		f.c_custoFinancFornecTipoParcelamento.value="";
		f.c_custoFinancFornecQtdeParcelas.value="";
		}
		
	if (trim(f.c_custoFinancFornecQtdeParcelas.value)=="") return;

//  N�o consulta novamente se for a mesma consulta anterior
	if ((f.c_custoFinancFornecTipoParcelamento.value==f.c_custoFinancFornecTipoParcelamentoUltConsulta.value)&&
		(f.c_custoFinancFornecQtdeParcelas.value==f.c_custoFinancFornecQtdeParcelasUltConsulta.value)) return;
	
	f.c_custoFinancFornecTipoParcelamentoUltConsulta.value=f.c_custoFinancFornecTipoParcelamento.value;
	f.c_custoFinancFornecQtdeParcelasUltConsulta.value=f.c_custoFinancFornecQtdeParcelas.value;

	f.c_custoFinancFornecParcelamentoDescricao.value=descricaoCustoFinancFornecTipoParcelamento(f.c_custoFinancFornecTipoParcelamento.value);
	if (f.c_custoFinancFornecTipoParcelamento.value==COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__COM_ENTRADA) {
		f.c_custoFinancFornecParcelamentoDescricao.value += " (1+" + f.c_custoFinancFornecQtdeParcelas.value + ")";
		}
	else if (f.c_custoFinancFornecTipoParcelamento.value==COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__SEM_ENTRADA) {
		f.c_custoFinancFornecParcelamentoDescricao.value += " (0+" + f.c_custoFinancFornecQtdeParcelas.value + ")";
		}

	window.status="Aguarde, consultando pre�os ...";
	$("#divAjaxRunning").show();
	
	strUrl = "../Global/AjaxCustoFinancFornecConsultaPrecoBD.asp";
	strUrl+="?tipoParcelamento="+f.c_custoFinancFornecTipoParcelamento.value;
	strUrl+="&qtdeParcelas="+f.c_custoFinancFornecQtdeParcelas.value;
	strUrl+="&loja="+f.c_loja.value;
	strUrl+="&listaProdutos="+strListaProdutos;
//  Prevents server from using a cached file
	strUrl=strUrl+"&sid="+Math.random()+Math.random();
	objAjaxCustoFinancFornecConsultaPreco.onreadystatechange=trataRespostaAjaxCustoFinancFornecSincronizaPrecos;
	objAjaxCustoFinancFornecConsultaPreco.open("GET",strUrl,true);
	objAjaxCustoFinancFornecConsultaPreco.send(null);
}

// RETORNA O VALOR TOTAL DO PEDIDO A SER USADO P/ CALCULAR A FORMA DE PAGAMENTO
function fp_vl_total_pedido( ) {
var f,i,mTotVenda,mTotNF;
	f=fORC;
	mTotVenda=0;
	for (i=0; i<f.c_qtde.length; i++) mTotVenda=mTotVenda+converte_numero(f.c_qtde[i].value)*converte_numero(f.c_vl_unitario[i].value);
	mTotNF=0;
	for (i=0; i<f.c_qtde.length; i++) mTotNF=mTotNF+converte_numero(f.c_qtde[i].value)*converte_numero(f.c_vl_NF[i].value);
//  Retorna total de pre�o NF (tem valor de NF, ou seja, pedido c/ RA)?
	if (mTotNF > 0) {
		return mTotNF;
		}
//  Retorna total de pre�o de venda
	else {
		return mTotVenda;
		}
}

// PARCELA �NICA
function pu_atualiza_valor( ){
var f,vt;
	f=fORC;
	if (converte_numero(trim(f.c_pu_valor.value))>0) return;
	vt=fp_vl_total_pedido();
	f.c_pu_valor.value=formata_moeda(vt);
}

// PARCELADO NO CART�O (INTERNET)
function pc_calcula_valor_parcela( ){
var f,n,t;
	f=fORC;
	if (trim(f.c_pc_qtde.value)=='') return;
	n=converte_numero(f.c_pc_qtde.value);
	if (n<=0) return;
	t=fp_vl_total_pedido();
	p=t/n;
	f.c_pc_valor.value=formata_moeda(p);
}

// PARCELADO NO CART�O (MAQUINETA)
function pc_maquineta_calcula_valor_parcela( ){
	var f,n,t;
	f=fORC;
	if (trim(f.c_pc_maquineta_qtde.value)=='') return;
	n=converte_numero(f.c_pc_maquineta_qtde.value);
	if (n<=0) return;
	t=fp_vl_total_pedido();
	p=t/n;
	f.c_pc_maquineta_valor.value=formata_moeda(p);
}

// PARCELADO COM ENTRADA
function pce_preenche_sugestao_intervalo() {
var f;
	f=fORC;
	if (converte_numero(trim(f.c_pce_prestacao_periodo.value))>0) return;
	f.c_pce_prestacao_periodo.value='30';
}

function pce_calcula_valor_parcela( ){
var f,n,e,t;
	f=fORC;
	t=fp_vl_total_pedido();
	if (trim(f.c_pce_entrada_valor.value)=='') return;
	e=converte_numero(f.c_pce_entrada_valor.value);
	if (e<=0) return;
	if (trim(f.c_pce_prestacao_qtde.value)=='') return;
	n=converte_numero(f.c_pce_prestacao_qtde.value);
	if (n<=0) return;
	p=(t-e)/n;
	f.c_pce_prestacao_valor.value=formata_moeda(p);
}

// PARCELADO SEM ENTRADA
function pse_preenche_sugestao_intervalo() {
var f;
	f=fORC;
	if (converte_numero(trim(f.c_pse_demais_prest_periodo.value))>0) return;
	f.c_pse_demais_prest_periodo.value='30';
}

function pse_calcula_valor_parcela( ){
var f,n,e,t;
	f=fORC;
	t=fp_vl_total_pedido();
	if (trim(f.c_pse_prim_prest_valor.value)=='') return;
	e=converte_numero(f.c_pse_prim_prest_valor.value);
	if (e<=0) return;
	if (trim(f.c_pse_demais_prest_qtde.value)=='') return;
	n=converte_numero(f.c_pse_demais_prest_qtde.value);
	if (n<=0) return;
	p=(t-e)/n;
	f.c_pse_demais_prest_valor.value=formata_moeda(p);
}

function pce_sugestao_forma_pagto( ) {
var f, p, s, i, n;
	f=fORC;
	f.c_forma_pagto.value="";
	p=converte_numero(f.c_pce_prestacao_periodo.value);
	if (p<=0) return;
	n=converte_numero(f.c_pce_prestacao_qtde.value);
	if (n<=0) return;
	s='0';
	for (i=1; i<=n; i++) {
		s=s+'/';
		s=s+formata_inteiro(i*p);
		}
	f.c_forma_pagto.value=s;
}

function pse_sugestao_forma_pagto( ) {
var f, p1, p2, s, i, n;
	f=fORC;
	f.c_forma_pagto.value="";
	p1=converte_numero(f.c_pse_prim_prest_apos.value);
	if (p1<=0) return;
	p2=converte_numero(f.c_pse_demais_prest_periodo.value);
	if (p2<=0) return;
	n=converte_numero(f.c_pse_demais_prest_qtde.value);
	if (n<=0) return;
	s=formata_inteiro(p1);
	for (i=1; i<=n; i++) {
		s=s+'/';
		s=s+formata_inteiro(i*p2);
		}
	f.c_forma_pagto.value=s;
}

function recalcula_RA( ) {
var f,i,mTotVenda,mTotNF;
	f=fORC;
	mTotVenda=0;
	for (i=0; i<f.c_vl_total.length; i++) mTotVenda=mTotVenda+converte_numero(f.c_vl_total[i].value);
	mTotNF=0;
	for (i=0; i<f.c_qtde.length; i++) mTotNF=mTotNF+converte_numero(f.c_qtde[i].value)*converte_numero(f.c_vl_NF[i].value);
	f.c_total_NF.value = formata_moeda(mTotNF);
	f.c_total_RA.value = formata_moeda(mTotNF-mTotVenda);
	if (mTotNF >=mTotVenda) f.c_total_RA.style.color="green"; else f.c_total_RA.style.color="red";
}

function recalcula_total( id ) {
var idx, m, m_lista, m_unit, d, f, i, s;
	f=fORC;
	idx=parseInt(id)-1;
	if (f.c_produto[idx].value=="") return;
	m_lista=converte_numero(f.c_preco_lista[idx].value);
	m_unit=converte_numero(f.c_vl_unitario[idx].value);
	if (m_lista==0) d=0; else d=100*(m_lista-m_unit)/m_lista;
	if (d==0) s=""; else s=formata_perc_desc(d);
	if (f.c_desc[idx].value!=s) f.c_desc[idx].value=s;
	s=formata_moeda(parseInt(f.c_qtde[idx].value)*m_unit);
	if (f.c_vl_total[idx].value!=s) f.c_vl_total[idx].value=s;
	m=0;
	for (i=0; i<f.c_vl_total.length; i++) m=m+converte_numero(f.c_vl_total[i].value);
	s=formata_moeda(m);
	if (f.c_total_geral.value!=s) f.c_total_geral.value=s;
}

function recalcula_total_todas_linhas() {
var f,i,t,m_lista,m_unit,d,m,s;
	f = fORC;
	t=0;
	for (i=0; i < f.c_produto.length; i++) {
		if (trim(f.c_produto[i].value)!="") {
			m_lista=converte_numero(f.c_preco_lista[i].value);
			m_unit=converte_numero(f.c_vl_unitario[i].value);
			if (m_lista==0) d=0; else d=100*(m_lista-m_unit)/m_lista;
			if (d==0) s=""; else s=formata_perc_desc(d);
			if (f.c_desc[i].value!=s) f.c_desc[i].value=s;
			m=parseInt(f.c_qtde[i].value)*m_unit;
			f.c_vl_total[i].value=formata_moeda(m);
			t=t+m;
			}
		}
	f.c_total_geral.value=formata_moeda(t);
}

function preenche_sugestao_forma_pagto( ) {
var f, n, t, p, s;
	f=fORC;
	n=converte_numero(f.c_qtde_parcelas.value);
	t=converte_numero(f.c_total_geral.value);
	if (n > 0) {
		p=t/n;
		s = "Pagamento em " + n;
		if (n==1) s = s + " parcela de "; else s = s + " parcelas de ";
		s = s + SIMBOLO_MONETARIO + " " + formata_moeda(p);
		f.c_forma_pagto.value=s;
		}
	else f.c_forma_pagto.value="";
}

function LimparCamposEndEtg(f) {
	f.EndEtg_endereco.value = "";
	f.EndEtg_endereco_numero.value = "";
	f.EndEtg_endereco_complemento.value = "";
	f.EndEtg_bairro.value = "";
	f.EndEtg_cidade.value = "";
	f.EndEtg_uf.value = "";
	f.EndEtg_cep.value = "";
    f.EndEtg_obs.selectedIndex = 0;

    <%if blnUsarMemorizacaoCompletaEnderecos and not eh_cpf then %>
        f.EndEtg_tipo_pessoa[0].checked = false;
        f.EndEtg_tipo_pessoa[1].checked = false;
	    f.EndEtg_cnpj_cpf_PJ.value="";
        f.EndEtg_ie_PJ.value="";
        f.EndEtg_contribuinte_icms_status_PJ[0].checked = false;
        f.EndEtg_contribuinte_icms_status_PJ[1].checked = false;
        f.EndEtg_contribuinte_icms_status_PJ[2].checked = false;

        f.EndEtg_cnpj_cpf_PF.value="";
        f.EndEtg_produtor_rural_status_PF[0].checked = false;
        f.EndEtg_produtor_rural_status_PF[1].checked = false;
        f.EndEtg_ie_PF.value="";
        f.EndEtg_contribuinte_icms_status_PF[0].checked = false;
        f.EndEtg_contribuinte_icms_status_PF[1].checked = false;
        f.EndEtg_contribuinte_icms_status_PF[2].checked = false;

        f.EndEtg_nome.value="";
        f.EndEtg_ddd_res.value="";
        f.EndEtg_tel_res.value="";
        f.EndEtg_ddd_cel.value="";
        f.EndEtg_tel_cel.value="";
        f.EndEtg_ddd_com.value="";
        f.EndEtg_tel_com.value="";
        f.EndEtg_ramal_com.value="";
        f.EndEtg_ddd_com_2.value="";
        f.EndEtg_tel_com_2.value="";
        f.EndEtg_ramal_com_2.value = "";

        trataProdutorRuralEndEtg_PF(null);
        trocarEndEtgTipoPessoa(null);
    <% end if%>
}

function fORCConfirma( f ) {
var i,s,n,ni,nip,vp,vp2,ve,idx,forma_pagto_ok,vtFP,strMsgErro;
var blnConfirmaDifRAeValores=false;
	
	recalcula_total_todas_linhas();

	s = "" + f.c_obs1.value;
	if (s.length > MAX_TAM_OBS1) {
		alert('Conte�do de "Observa��es I" excede em ' + (s.length-MAX_TAM_OBS1) + ' caracteres o tamanho m�ximo de ' + MAX_TAM_OBS1 + '!!');
		f.c_obs1.focus();
		return;
		}

	s = "" + f.c_forma_pagto.value;
	if (s.length > MAX_TAM_FORMA_PAGTO) {
		alert('Conte�do de "Forma de Pagamento" excede em ' + (s.length-MAX_TAM_FORMA_PAGTO) + ' caracteres o tamanho m�ximo de ' + MAX_TAM_FORMA_PAGTO + '!!');
		f.c_forma_pagto.focus();
		return;
		}

//  Consiste a nova vers�o da forma de pagamento
	if (f.versao_forma_pagamento.value=='2') {
		vtFP=fp_vl_total_pedido();
		idx=-1;
		forma_pagto_ok=false;
	//	� Vista
		idx++;
		if (f.rb_forma_pagto[idx].checked) {
			if (trim(f.op_av_forma_pagto.value)=='') {
				alert('Indique a forma de pagamento!!');
				f.op_av_forma_pagto.focus();
				return;
				}
			forma_pagto_ok=true;
			}

	//	Parcela �nica
		idx++;
		if (f.rb_forma_pagto[idx].checked) {
			if (trim(f.op_pu_forma_pagto.value)=='') {
				alert('Indique a forma de pagamento da parcela �nica!!');
				f.op_pu_forma_pagto.focus();
				return;
				}
			if (trim(f.c_pu_valor.value)=='') {
				alert('Indique o valor da parcela �nica!!');
				f.c_pu_valor.focus();
				return;
				}
			ve=converte_numero(f.c_pu_valor.value);
			if (ve<=0) {
				alert('Valor da parcela �nica � inv�lido!!');
				f.c_pu_valor.focus();
				return;
				}
			if (trim(f.c_pu_vencto_apos.value)=='') {
				alert('Indique o intervalo de vencimento da parcela �nica!!');
				f.c_pu_vencto_apos.focus();
				return;
				}
			nip=converte_numero(f.c_pu_vencto_apos.value);
			if (nip<=0) {
				alert('Intervalo de vencimento da parcela �nica � inv�lido!!');
				f.c_pu_vencto_apos.focus();
				return;
				}
			if (ve<vtFP) {
				alert('O valor do pagamento da parcela �nica est� abaixo do valor do or�amento!!');
				f.c_pu_valor.focus();
				return;
				}
			if (blnConfirmaDifRAeValores) {
				if (ve>vtFP) {
					if (!window.confirm('O valor do pagamento da parcela �nica est� ' + SIMBOLO_MONETARIO + ' ' + formata_moeda(ve-vtFP) + ' acima do valor nominal do or�amento.\nConfirma?')) {
						f.c_pu_valor.focus();
						return;
						}
					}
				}
			forma_pagto_ok=true;
			}

	//	Parcelado no cart�o (internet)
		idx++;
		if (f.rb_forma_pagto[idx].checked) {
			if (trim(f.c_pc_qtde.value)=='') {
				alert('Indique a quantidade de parcelas!!');
				f.c_pc_qtde.focus();
				return;
				}
			n=converte_numero(f.c_pc_qtde.value);
			if (n < 1) {
				alert('Quantidade de parcelas inv�lida!!');
				f.c_pc_qtde.focus();
				return;
				}
			if (trim(f.c_pc_valor.value)=='') {
				alert('Indique o valor da parcela!!');
				f.c_pc_valor.focus();
				return;
				}
			vp2=converte_numero(f.c_pc_valor.value);
			if (vp2<=0) {
				alert('Valor de parcela inv�lido!!');
				f.c_pc_valor.focus();
				return;
				}
			vp=vtFP/n;
			if ((vp-vp2)>0.5) {
				alert('O valor da parcela diverge do valor calculado pelo sistema (' + SIMBOLO_MONETARIO + ' ' + formata_moeda(vp) + ')!!');
				f.c_pc_valor.focus();
				return;
				}
			if (blnConfirmaDifRAeValores) {
				if ((vp2-vp)>0.5) {
					if (!window.confirm('O valor da parcela est� ' + SIMBOLO_MONETARIO + ' ' + formata_moeda(vp2-vp) + ' acima do valor nominal.\nConfirma?')) {
						f.c_pc_valor.focus();
						return;
						}
					}
				}
			forma_pagto_ok=true;
			}

		//	Parcelado no cart�o (maquineta)
		idx++;
		if (f.rb_forma_pagto[idx].checked) {
			if (trim(f.c_pc_maquineta_qtde.value)=='') {
				alert('Indique a quantidade de parcelas!!');
				f.c_pc_maquineta_qtde.focus();
				return;
			}
			n=converte_numero(f.c_pc_maquineta_qtde.value);
			if (n < 1) {
				alert('Quantidade de parcelas inv�lida!!');
				f.c_pc_maquineta_qtde.focus();
				return;
			}
			if (trim(f.c_pc_maquineta_valor.value)=='') {
				alert('Indique o valor da parcela!!');
				f.c_pc_maquineta_valor.focus();
				return;
			}
			vp2=converte_numero(f.c_pc_maquineta_valor.value);
			if (vp2<=0) {
				alert('Valor de parcela inv�lido!!');
				f.c_pc_maquineta_valor.focus();
				return;
			}
			vp=vtFP/n;
			if ((vp-vp2)>0.5) {
				alert('O valor da parcela diverge do valor calculado pelo sistema (' + SIMBOLO_MONETARIO + ' ' + formata_moeda(vp) + ')!!');
				f.c_pc_maquineta_valor.focus();
				return;
			}
			if (blnConfirmaDifRAeValores) {
				if ((vp2-vp)>0.5) {
					if (!window.confirm('O valor da parcela est� ' + SIMBOLO_MONETARIO + ' ' + formata_moeda(vp2-vp) + ' acima do valor nominal.\nConfirma?')) {
						f.c_pc_maquineta_valor.focus();
						return;
					}
				}
			}
			forma_pagto_ok=true;
		}

	//	Parcelado com entrada
		idx++;
		if (f.rb_forma_pagto[idx].checked) {
			if (trim(f.op_pce_entrada_forma_pagto.value)=='') {
				alert('Indique a forma de pagamento da entrada!!');
				f.op_pce_entrada_forma_pagto.focus();
				return;
				}
			if (trim(f.c_pce_entrada_valor.value)=='') {
				alert('Indique o valor da entrada!!');
				f.c_pce_entrada_valor.focus();
				return;
				}
			ve=converte_numero(f.c_pce_entrada_valor.value);
			if (ve<=0) {
				alert('Valor da entrada inv�lido!!');
				f.c_pce_entrada_valor.focus();
				return;
				}
			if (trim(f.op_pce_prestacao_forma_pagto.value)=='') {
				alert('Indique a forma de pagamento das presta��es!!');
				f.op_pce_prestacao_forma_pagto.focus();
				return;
				}
			if (trim(f.c_pce_prestacao_qtde.value)=='') {
				alert('Indique a quantidade de presta��es!!');
				f.c_pce_prestacao_qtde.focus();
				return;
				}
			n=converte_numero(f.c_pce_prestacao_qtde.value);
			if (n<=0) {
				alert('Quantidade de presta��es inv�lida!!');
				f.c_pce_prestacao_qtde.focus();
				return;
				}
			if (trim(f.c_pce_prestacao_valor.value)=='') {
				alert('Indique o valor da presta��o!!');
				f.c_pce_prestacao_valor.focus();
				return;
				}
			vp2=converte_numero(f.c_pce_prestacao_valor.value);
			if (vp2<=0) {
				alert('Valor de presta��o inv�lido!!');
				f.c_pce_prestacao_valor.focus();
				return;
				}
			vp=(vtFP-ve)/n;
			if ((vp-vp2)>0.5) {
				alert('O valor da presta��o diverge do valor calculado pelo sistema (' + SIMBOLO_MONETARIO + ' ' + formata_moeda(vp) + ')!!');
				f.c_pce_prestacao_valor.focus();
				return;
				}
			if (blnConfirmaDifRAeValores) {
				if ((vp2-vp)>0.5) {
					if (!window.confirm('O valor da parcela est� ' + SIMBOLO_MONETARIO + ' ' + formata_moeda(vp2-vp) + ' acima do valor nominal.\nConfirma?')) {
						f.c_pce_prestacao_valor.focus();
						return;
						}
					}
				}
			if (trim(f.c_pce_prestacao_periodo.value)=='') {
				alert('Indique o intervalo de vencimento entre as parcelas!!');
				f.c_pce_prestacao_periodo.focus();
				return;
				}
			ni=converte_numero(f.c_pce_prestacao_periodo.value);
			if (ni<=0) {
				alert('Intervalo de vencimento inv�lido!!');
				f.c_pce_prestacao_periodo.focus();
				return;
				}
			forma_pagto_ok=true;
			}

	//	Parcelado sem entrada
		idx++;
		if (f.rb_forma_pagto[idx].checked) {
			if (trim(f.op_pse_prim_prest_forma_pagto.value)=='') {
				alert('Indique a forma de pagamento da 1� presta��o!!');
				f.op_pse_prim_prest_forma_pagto.focus();
				return;
				}
			if (trim(f.c_pse_prim_prest_valor.value)=='') {
				alert('Indique o valor da 1� presta��o!!');
				f.c_pse_prim_prest_valor.focus();
				return;
				}
			ve=converte_numero(f.c_pse_prim_prest_valor.value);
			if (ve<=0) {
				alert('Valor da 1� presta��o inv�lido!!');
				f.c_pse_prim_prest_valor.focus();
				return;
				}
			if (trim(f.c_pse_prim_prest_apos.value)=='') {
				alert('Indique o intervalo de vencimento da 1� parcela!!');
				f.c_pse_prim_prest_apos.focus();
				return;
				}
			nip=converte_numero(f.c_pse_prim_prest_apos.value);
			if (nip<=0) {
				alert('Intervalo de vencimento da 1� parcela � inv�lido!!');
				f.c_pse_prim_prest_apos.focus();
				return;
				}
			if (trim(f.op_pse_demais_prest_forma_pagto.value)=='') {
				alert('Indique a forma de pagamento das demais presta��es!!');
				f.op_pse_demais_prest_forma_pagto.focus();
				return;
				}
			if (trim(f.c_pse_demais_prest_qtde.value)=='') {
				alert('Indique a quantidade das demais presta��es!!');
				f.c_pse_demais_prest_qtde.focus();
				return;
				}
			n=converte_numero(f.c_pse_demais_prest_qtde.value);
			if (n<=0) {
				alert('Quantidade de presta��es inv�lida!!');
				f.c_pse_demais_prest_qtde.focus();
				return;
				}
			if (trim(f.c_pse_demais_prest_valor.value)=='') {
				alert('Indique o valor das demais presta��es!!');
				f.c_pse_demais_prest_valor.focus();
				return;
				}
			vp2=converte_numero(f.c_pse_demais_prest_valor.value);
			if (vp2<=0) {
				alert('Valor de presta��o inv�lido!!');
				f.c_pse_demais_prest_valor.focus();
				return;
				}
			vp=(vtFP-ve)/n;
			if ((vp-vp2)>0.5) {
				alert('O valor da presta��o diverge do valor calculado pelo sistema (' + SIMBOLO_MONETARIO + ' ' + formata_moeda(vp) + ')!!');
				f.c_pse_demais_prest_valor.focus();
				return;
				}
			if (blnConfirmaDifRAeValores) {
				if ((vp2-vp)>0.5) {
					if (!window.confirm('O valor da parcela est� ' + SIMBOLO_MONETARIO + ' ' + formata_moeda(vp2-vp) + ' acima do valor nominal.\nConfirma?')) {
						f.c_pse_demais_prest_valor.focus();
						return;
						}
					}
				}
			if (trim(f.c_pse_demais_prest_periodo.value)=='') {
				alert('Indique o intervalo de vencimento entre as parcelas!!');
				f.c_pse_demais_prest_periodo.focus();
				return;
				}
			ni=converte_numero(f.c_pse_demais_prest_periodo.value);
			if (ni<=0) {
				alert('Intervalo de vencimento inv�lido!!');
				f.c_pse_demais_prest_periodo.focus();
				return;
				}
			forma_pagto_ok=true;
			}

		if (!forma_pagto_ok) {
			alert('Indique a forma de pagamento!!');
			return;
			}
		}

	recalcula_RA();
	if (blnConfirmaDifRAeValores) {
		if (f.c_total_RA.value != f.c_total_RA_original.value) {
			if (!confirm("O valor do RA � de " + SIMBOLO_MONETARIO + " " + formata_moeda(converte_numero(f.c_total_RA.value))+"\nContinua?")) return;
			}
		}

	if (f.c_FlagEndEntregaEditavel.value == "S") {
		blnTemEndEntrega=false;
		if (trim(f.EndEtg_endereco.value)!="") blnTemEndEntrega=true;
		if (trim(f.EndEtg_endereco_numero.value)!="") blnTemEndEntrega=true;
		if (trim(f.EndEtg_endereco_complemento.value)!="") blnTemEndEntrega=true;
		if (trim(f.EndEtg_bairro.value)!="") blnTemEndEntrega=true;
		if (trim(f.EndEtg_cidade.value)!="") blnTemEndEntrega=true;
		if (trim(f.EndEtg_uf.value)!="") blnTemEndEntrega=true;
		if (trim(f.EndEtg_cep.value)!="") blnTemEndEntrega=true;
		if (trim(f.EndEtg_obs.value)!="") blnTemEndEntrega=true;

<%if blnUsarMemorizacaoCompletaEnderecos and not eh_cpf then %>

        if( $('input[name="EndEtg_tipo_pessoa"]:checked').val()) blnTemEndEntrega = true;

        //simplesmente testamos todos os campos, qualquer valor em qq campo significa preenchimento
        //n�o deve estar em campo oculto porque o usu�rio deve clicar no X para limpar, e o X limpa todos os campos, inclusive os n�o visiveis no momento

        //pj
        if (trim(f.EndEtg_cnpj_cpf_PJ.value) != "") blnTemEndEntrega = true;
        if (trim(f.EndEtg_ie_PJ.value) != "") blnTemEndEntrega = true;
        if( $('input[name="EndEtg_contribuinte_icms_status_PJ"]:checked').val()) blnTemEndEntrega = true;

        //pf
        if (trim(f.EndEtg_cnpj_cpf_PF.value) != "") blnTemEndEntrega = true;
        if( $('input[name="EndEtg_produtor_rural_status_PF"]:checked').val()) blnTemEndEntrega = true;
        if (trim(f.EndEtg_ie_PF.value) != "") blnTemEndEntrega = true;
        if( $('input[name="EndEtg_contribuinte_icms_status_PF"]:checked').val()) blnTemEndEntrega = true;

        //ambos
        if (trim(f.EndEtg_nome.value) != "") blnTemEndEntrega = true;

        //pj
        if (trim(f.EndEtg_ddd_com.value) != "") blnTemEndEntrega = true;
        if (trim(f.EndEtg_tel_com.value) != "") blnTemEndEntrega = true;
        if (trim(f.EndEtg_ramal_com.value) != "") blnTemEndEntrega = true;
        if (trim(f.EndEtg_ddd_com_2.value) != "") blnTemEndEntrega = true;
        if (trim(f.EndEtg_tel_com_2.value) != "") blnTemEndEntrega = true;
        if (trim(f.EndEtg_ramal_com_2.value) != "") blnTemEndEntrega = true;

        //pf
        if (trim(f.EndEtg_ddd_res.value) != "") blnTemEndEntrega = true;
        if (trim(f.EndEtg_tel_res.value) != "") blnTemEndEntrega = true;
        if (trim(f.EndEtg_ddd_cel.value) != "") blnTemEndEntrega = true;
        if (trim(f.EndEtg_tel_cel.value) != "") blnTemEndEntrega = true;
<% end if%>


	<%if r_orcamento.st_memorizacao_completa_enderecos <> 0 then %>

		if (trim(f.endereco__endereco.value) == "") {
            alert('Endere�o n�o foi preenchido corretamente!!');
            f.endereco__endereco.focus();
            return;
        }
        if (trim(f.endereco__bairro.value) == "") {
            alert('Endere�o n�o foi preenchido corretamente!!');
            f.endereco__bairro.focus();
            return;
        }

        if (trim(f.endereco__numero.value) == "") {
            alert('Endere�o n�o foi preenchido corretamente!!');
            f.endereco__numero.focus();
            return;
        }
        if (trim(f.endereco__cidade.value) == "") {
            alert('Endere�o n�o foi preenchido corretamente!!');
            f.endereco__cidade.focus();
            return;
        }

        if (trim(f.endereco__uf.value) == "") {
            alert('Endere�o n�o foi preenchido corretamente!!');
            f.endereco__uf.focus();
            return;
        }

        if (trim(f.endereco__cep.value) == "") {
            alert('Endere�o n�o foi preenchido corretamente!!');
            f.endereco__cep.focus();
            return;
        }

        if ((trim(f.cliente__email.value) != "") && (!email_ok(f.cliente__email.value))) {
            alert('E-mail inv�lido!!');
            f.cliente__email.focus();
            return;
		}

        if ((trim(f.cliente__email_xml.value) != "") && (!email_ok(f.cliente__email_xml.value))) {
            alert('E-mail xml inv�lido!!');
            f.cliente__email_xml.focus();
            return;
        }


       <% if cliente__tipo = ID_PF then %>

		    if ( (trim(f.cliente__ddd_res.value) != "" && !ddd_ok(f.cliente__ddd_res.value)) || (trim(f.cliente__ddd_res.value) == "" && trim(f.cliente__tel_res.value) != "") ) {
                alert('DDD inv�lido!!');
                f.cliente__ddd_res.focus();
                return;
            }

		    if ( (trim(f.cliente__tel_res.value) != "" && !telefone_ok(f.cliente__tel_res.value)) || (trim(f.cliente__ddd_res.value) != "" && trim(f.cliente__tel_res.value) == "") ) {
                alert('Telefone residencial inv�lido!!');
                f.cliente__tel_res.focus();
                return;
            }

		    if ( (trim(f.cliente__ddd_cel.value) != "" && !ddd_ok(f.cliente__ddd_cel.value)) || (trim(f.cliente__ddd_cel.value) == "" && trim(f.cliente__tel_cel.value) != "") ) {
                alert('Celular com DDD inv�lido!!');
                f.cliente__ddd_cel.focus();
                return;
            }

		    if ( (trim(f.cliente__tel_cel.value) != "" && !telefone_ok(f.cliente__tel_cel.value)) || (trim(f.cliente__ddd_cel.value) != "" && trim(f.cliente__tel_cel.value) == "") ) {
                alert('Telefone celular inv�lido!!');
                f.cliente__tel_cel.focus();
                return;
            }


		    if ( (trim(f.cliente__ddd_com.value) != "" && !ddd_ok(f.cliente__ddd_com.value)) || (trim(f.cliente__ddd_com.value) == "" && trim(f.cliente__tel_com.value) != "") ) {
                alert('DDD comercial inv�lido!!');
                f.cliente__ddd_com.focus();
                return;
            }

		    if ( (trim(f.cliente__tel_com.value) != "" && !telefone_ok(f.cliente__tel_com.value)) || (trim(f.cliente__ddd_com.value) != "" && trim(f.cliente__tel_com.value) == "") ) {
                alert('Telefone comercial inv�lido!!');
                f.cliente__tel_com.focus();
                return;
            }

		    if (trim(f.cliente__ddd_com.value) == "" && trim(f.cliente__ramal_com.value) != "") {
                alert('DDD comercial inv�lido!!');
                f.cliente__ddd_com.focus();
                return;
            }

		    if (trim(f.cliente__tel_com.value) == "" && trim(f.cliente__ramal_com.value) != "") {
                alert('Telefone comercial inv�lido!!');
                f.cliente__tel_com.focus();
                return;
            }

            if (trim(f.cliente__tel_res.value) == "" && trim(f.cliente__tel_cel.value) == "" && trim(f.cliente__tel_com.value) == "") {
                alert('Necess�rio preencher ao menos um telefone!!');
                f.cliente__ddd_cel.focus();
                return;
            }



            if (f.rb_produtor_rural[1].checked) {
                if (!f.rb_contribuinte_icms[1].checked) {
                    alert('Para ser cadastrado como Produtor Rural, � necess�rio ser contribuinte do ICMS e possuir n� de IE!!');
                    return;
                }
                if ((!f.rb_contribuinte_icms[0].checked) && (!f.rb_contribuinte_icms[1].checked) && (!f.rb_contribuinte_icms[2].checked)) {
                    alert('Informe se o cliente � contribuinte do ICMS, n�o contribuinte ou isento!!');
                    return;
                }
                if ((f.rb_contribuinte_icms[1].checked) && (trim(f.cliente__ie.value) == "")) {
                    alert('Se o cliente � contribuinte do ICMS a inscri��o estadual deve ser preenchida!!');
                    f.cliente__ie.focus();
                    return;
                }
                if ((f.rb_contribuinte_icms[0].checked) && (f.cliente__ie.value.toUpperCase().indexOf('ISEN') >= 0)) {
                    alert('Se cliente � n�o contribuinte do ICMS, n�o pode ter o valor ISENTO no campo de Inscri��o Estadual!!');
                    f.cliente__ie.focus();
                    return;
                }
                if ((f.rb_contribuinte_icms[1].checked) && (f.cliente__ie.value.toUpperCase().indexOf('ISEN') >= 0)) {
                    alert('Se cliente � contribuinte do ICMS, n�o pode ter o valor ISENTO no campo de Inscri��o Estadual!!');
                    f.cliente__ie.focus();
                    return;
                }
                if (f.rb_contribuinte_icms[2].checked) {
                    if (f.cliente__ie.value != "") {
                        alert("Se o Contribuinte ICMS � isento, o campo IE deve ser vazio!");
                        f.cliente__ie.focus();
                        return;
                    }
                }
            }

		<% else %>

            if ((trim(f.cliente__email.value) != "") && (!email_ok(f.cliente__email.value))) {
                alert('E-mail inv�lido!!');
                f.cliente__email.focus();
                return;
            }

            if ((trim(f.cliente__email_xml.value) != "") && (!email_ok(f.cliente__email_xml.value))) {
                alert('E-mail (XML) inv�lido!!');
                f.cliente__email_xml.focus();
                return;
            }

               <% if CStr(r_orcamento.loja) <> CStr(NUMERO_LOJA_ECOMMERCE_AR_CLUBE) then %>
                // PARA CLIENTE PJ, � OBRIGAT�RIO O PREENCHIMENTO DO E-MAIL
                if ((trim(f.cliente__email.value) == "") && (trim(f.cliente__email_xml.value) == "")) {
                    alert("� obrigat�rio informar um endere�o de e-mail");
                    f.cliente__email.focus();
                    return;
                }
                <% end if %>

            if ((f.rb_contribuinte_icms[1].checked) && (trim(f.cliente__ie.value) == "")) {
                alert('Se o cliente � contribuinte do ICMS a inscri��o estadual deve ser preenchida!!');
                f.cliente__ie.focus();
                return;
            }
            if ((f.rb_contribuinte_icms[0].checked) && (f.cliente__ie.value.toUpperCase().indexOf('ISEN') >= 0)) {
                alert('Se cliente � n�o contribuinte do ICMS, n�o pode ter o valor ISENTO no campo de Inscri��o Estadual!!');
                f.cliente__ie.focus();
                return;
            }
            if ((f.rb_contribuinte_icms[1].checked) && (f.cliente__ie.value.toUpperCase().indexOf('ISEN') >= 0)) {
                alert('Se cliente � contribuinte do ICMS, n�o pode ter o valor ISENTO no campo de Inscri��o Estadual!!');
                f.cliente__ie.focus();
                return;
            }
            if (f.rb_contribuinte_icms[2].checked) {
                if (f.cliente__ie.value != "") {
                    alert("Se o Contribuinte ICMS � isento, o campo IE deve ser vazio!");
                    f.cliente__ie.focus();
                    return;
                }
            }
		    if ( (trim(f.cliente__ddd_com.value) != "" && !ddd_ok(f.cliente__ddd_com.value)) || (trim(f.cliente__ddd_com.value) == "" && trim(f.cliente__tel_com.value) != "") ) {
                alert('DDD comercial inv�lido!!');
                f.cliente__ddd_com.focus();
                return;
            }

		    if (trim(f.cliente__ddd_com.value) == "" && trim(f.cliente__ramal_com.value) != "") {
                alert('DDD comercial inv�lido!!');
                f.cliente__ddd_com.focus();
                return;
            }

		    if ( (trim(f.cliente__tel_com.value) != "" && !telefone_ok(f.cliente__tel_com.value)) || (trim(f.cliente__ddd_com.value) != "" && trim(f.cliente__tel_com.value) == "") ) {
                alert('Telefone comercial inv�lido!!');
                f.cliente__tel_com.focus();
                return;
            }

		    if (trim(f.cliente__tel_com.value) == "" && trim(f.cliente__ramal_com.value) != "") {
                alert('Telefone comercial inv�lido!!');
                f.cliente__tel_com.focus();
                return;
            }

		    if ( (trim(f.cliente__ddd_com_2.value) != "" && !ddd_ok(f.cliente__ddd_com_2.value)) || (trim(f.cliente__ddd_com_2.value) == "" && trim(f.cliente__tel_com_2.value) != "") ) {
                alert('DDD comercial 2 inv�lido!!');
                f.cliente__ddd_com_2.focus();
                return;
            }

		    if (trim(f.cliente__ddd_com_2.value) == "" && trim(f.cliente__ramal_com_2.value) != "") {
                alert('DDD comercial 2 inv�lido!!');
                f.cliente__ddd_com_2.focus();
                return;
            }

		    if ( (trim(f.cliente__tel_com_2.value) != "" && !telefone_ok(f.cliente__tel_com_2.value)) || (trim(f.cliente__ddd_com_2.value) != "" && trim(f.cliente__tel_com_2.value) == "") ) {
                alert('Telefone comercial 2 inv�lido!!');
                f.cliente__tel_com_2.focus();
                return;
            }

		    if (trim(f.cliente__tel_com_2.value) == "" && trim(f.cliente__ramal_com_2.value) != "") {
                alert('Telefone comercial 2 inv�lido!!');
                f.cliente__tel_com_2.focus();
                return;
            }

            if (trim(f.cliente__tel_com.value) == "" && trim(f.cliente__tel_com_2.value) == "") {
                alert('Necess�rio preencher ao menos um telefone!!');
                f.cliente__ddd_com.focus();
                return;
            }

		<% end if%>

		
<% end if%>



		if (blnTemEndEntrega) {
		    var blnEndEtg_obs
		    blnEndEtg_obs = false;
		    if ((f.EndEtg_endereco.value != f.Verifica_End_Entrega.value) || (f.EndEtg_endereco_numero.value != f.Verifica_num.value) || (f.EndEtg_cidade.value != f.Verifica_Cidade.value) || (f.EndEtg_uf.value != f.Verifica_UF.value) || (f.EndEtg_cep.value != f.Verifica_CEP.value) || (f.EndEtg_obs.value != f.Verifica_Justificativa.value)){
		        blnEndEtg_obs = true;
		    }
			if (trim(f.EndEtg_endereco.value)=="") {
				alert('Endere�o de entrega n�o foi preenchido corretamente!!');
				f.EndEtg_endereco.focus();
				return;
				}

			if (trim(f.EndEtg_endereco_numero.value)=="") {
				alert('N�mero do endere�o de entrega n�o foi preenchido corretamente!!');
				f.EndEtg_endereco_numero.focus();
				return;
				}

			if (trim(f.EndEtg_bairro.value)=="") {
				alert('Bairro do endere�o de entrega n�o foi preenchido corretamente!!');
				f.EndEtg_bairro.focus();
				return;
				}

			if (trim(f.EndEtg_cidade.value)=="") {
				alert('Cidade do endere�o de entrega n�o foi preenchido corretamente!!');
				f.EndEtg_cidade.focus();
				return;
				}
			if ((trim(f.EndEtg_obs.value)=="")  && blnEndEtg_obs == true) {
			    alert('Justificativa do endere�o de entrega n�o foi preenchido corretamente!!');
			    f.EndEtg_obs.focus();
			    return;
			    }
			s=trim(f.EndEtg_uf.value);
			if ((s=="")||(!uf_ok(s))) {
				alert('UF do endere�o de entrega n�o foi preenchido corretamente!!');
				f.EndEtg_uf.focus();
				return;
				}
				
			if (!cep_ok(f.EndEtg_cep.value)) {
				alert('CEP do endere�o de entrega n�o foi preenchido corretamente!!');
				f.EndEtg_cep.focus();
				return;
				}



<%if blnUsarMemorizacaoCompletaEnderecos and Not eh_cpf then%>
                var EndEtg_tipo_pessoa = $('input[name="EndEtg_tipo_pessoa"]:checked').val();
                if (!EndEtg_tipo_pessoa)
                    EndEtg_tipo_pessoa = "";
                if (EndEtg_tipo_pessoa != "PJ" && EndEtg_tipo_pessoa != "PF") {
                    alert('Necess�rio escolher Pessoa Jur�dica ou Pessoa F�sica no Endere�o de entrega!!');
                    f.EndEtg_tipo_pessoa.focus();
                    return;
                }

                if (EndEtg_tipo_pessoa == "PJ") {
                    //Campos PJ: 

                    if (f.EndEtg_cnpj_cpf_PJ.value == "" || !cnpj_ok(f.EndEtg_cnpj_cpf_PJ.value)) {
                        alert('Endere�o de entrega: CNPJ inv�lido!!');
                        f.EndEtg_cnpj_cpf_PJ.focus();
                        return;
                    }

                    if ($('input[name="EndEtg_contribuinte_icms_status_PJ"]:checked').length == 0) {
                        alert('Endere�o de entrega: informe se o cliente � contribuinte do ICMS, n�o contribuinte ou isento!!');
                        f.EndEtg_contribuinte_icms_status_PJ.focus();
                        return;
                    }

                    if ((f.EndEtg_contribuinte_icms_status_PJ[1].checked) && (trim(f.EndEtg_ie_PJ.value) == "")) {
                        alert('Endere�o de entrega: se o cliente � contribuinte do ICMS a inscri��o estadual deve ser preenchida!!');
                        f.EndEtg_ie_PJ.focus();
                        return;
                    }
                    if ((f.EndEtg_contribuinte_icms_status_PJ[0].checked) && (f.EndEtg_ie_PJ.value.toUpperCase().indexOf('ISEN') >= 0)) {
                        alert('Endere�o de entrega: se cliente � n�o contribuinte do ICMS, n�o pode ter o valor ISENTO no campo de Inscri��o Estadual!!');
                        f.EndEtg_ie_PJ.focus();
                        return;
                    }
                    if ((f.EndEtg_contribuinte_icms_status_PJ[1].checked) && (f.EndEtg_ie_PJ.value.toUpperCase().indexOf('ISEN') >= 0)) {
                        alert('Endere�o de entrega: se cliente � contribuinte do ICMS, n�o pode ter o valor ISENTO no campo de Inscri��o Estadual!!');
                        f.EndEtg_ie_PJ.focus();
                        return;
                    }

                    if (trim(f.EndEtg_nome.value) == "") {
                        alert('Preencha a raz�o social no endere�o de entrega!!');
                        f.EndEtg_nome.focus();
                        return;
                    }

                    /*
                    telefones PJ:
                    EndEtg_ddd_com
                    EndEtg_tel_com
                    EndEtg_ramal_com
                    EndEtg_ddd_com_2
                    EndEtg_tel_com_2
                    EndEtg_ramal_com_2
    */

                    if (!ddd_ok(f.EndEtg_ddd_com.value)) {
                        alert('Endere�o de entrega: DDD inv�lido!!');
                        f.EndEtg_ddd_com.focus();
                        return;
                    }
                    if (!telefone_ok(f.EndEtg_tel_com.value)) {
                        alert('Endere�o de entrega: telefone inv�lido!!');
                        f.EndEtg_tel_com.focus();
                        return;
                    }
                    if ((f.EndEtg_ddd_com.value == "") && (f.EndEtg_tel_com.value != "")) {
                        alert('Endere�o de entrega: preencha o DDD do telefone.');
                        f.EndEtg_ddd_com.focus();
                        return;
                    }
                    if ((f.EndEtg_tel_com.value == "") && (f.EndEtg_ddd_com.value != "")) {
                        alert('Endere�o de entrega: preencha o telefone.');
                        f.EndEtg_tel_com.focus();
                        return;
                    }


                    if (!ddd_ok(f.EndEtg_ddd_com_2.value)) {
                        alert('Endere�o de entrega: DDD inv�lido!!');
                        f.EndEtg_ddd_com_2.focus();
                        return;
                    }
                    if (!telefone_ok(f.EndEtg_tel_com_2.value)) {
                        alert('Endere�o de entrega: telefone inv�lido!!');
                        f.EndEtg_tel_com_2.focus();
                        return;
                    }
                    if ((f.EndEtg_ddd_com_2.value == "") && (f.EndEtg_tel_com_2.value != "")) {
                        alert('Endere�o de entrega: preencha o DDD do telefone.');
                        f.EndEtg_ddd_com_2.focus();
                        return;
                    }
                    if ((f.EndEtg_tel_com_2.value == "") && (f.EndEtg_ddd_com_2.value != "")) {
                        alert('Endere�o de entrega: preencha o telefone.');
                        f.EndEtg_tel_com_2.focus();
                        return;
                    }

                }
                else {
                    //campos PF

                    if (f.EndEtg_cnpj_cpf_PF.value == "" || !cpf_ok(f.EndEtg_cnpj_cpf_PF.value)) {
                        alert('Endere�o de entrega: CPF inv�lido!!');
                        f.EndEtg_cnpj_cpf_PF.focus();
                        return;
                    }

                    if ((!f.EndEtg_produtor_rural_status_PF[0].checked) && (!f.EndEtg_produtor_rural_status_PF[1].checked)) {
                        alert('Endere�o de entrega: informe se o cliente � produtor rural ou n�o!!');
                        return;
                    }
                    if (!f.EndEtg_produtor_rural_status_PF[0].checked) {
                        if (!f.EndEtg_contribuinte_icms_status_PF[1].checked) {
                            alert('Endere�o de entrega: para ser cadastrado como Produtor Rural, � necess�rio ser contribuinte do ICMS e possuir n� de IE!!');
                            return;
                        }
                        if ((!f.EndEtg_contribuinte_icms_status_PF[0].checked) && (!f.EndEtg_contribuinte_icms_status_PF[1].checked) && (!f.EndEtg_contribuinte_icms_status_PF[2].checked)) {
                            alert('Endere�o de entrega: informe se o cliente � contribuinte do ICMS, n�o contribuinte ou isento!!');
                            return;
                        }
                        if ((f.EndEtg_contribuinte_icms_status_PF[1].checked) && (trim(f.EndEtg_ie_PF.value) == "")) {
                            alert('Endere�o de entrega: se o cliente � contribuinte do ICMS a inscri��o estadual deve ser preenchida!!');
                            f.EndEtg_ie_PF.focus();
                            return;
                        }
                        if ((f.EndEtg_contribuinte_icms_status_PF[0].checked) && (f.EndEtg_ie_PF.value.toUpperCase().indexOf('ISEN') >= 0)) {
                            alert('Endere�o de entrega: se cliente � n�o contribuinte do ICMS, n�o pode ter o valor ISENTO no campo de Inscri��o Estadual!!');
                            f.EndEtg_ie_PF.focus();
                            return;
                        }
                        if ((f.EndEtg_contribuinte_icms_status_PF[1].checked) && (f.EndEtg_ie_PF.value.toUpperCase().indexOf('ISEN') >= 0)) {
                            alert('Endere�o de entrega: se cliente � contribuinte do ICMS, n�o pode ter o valor ISENTO no campo de Inscri��o Estadual!!');
                            f.EndEtg_ie_PF.focus();
                            return;
                        }

                        if (f.EndEtg_contribuinte_icms_status_PF[2].checked) {
                            if (f.EndEtg_ie_PF.value != "") {
                                alert("Endere�o de entrega: se o Contribuinte ICMS � isento, o campo IE deve ser vazio!");
                                f.EndEtg_ie_PF.focus();
                                return;
                            }
                        }
                    }
            

                    if (trim(f.EndEtg_nome.value) == "") {
                        alert('Preencha o nome no endere�o de entrega!!');
                        f.EndEtg_nome.focus();
                        return;
                    }

                    /*
                    telefones PF:
                    EndEtg_ddd_res
                    EndEtg_tel_res
                    EndEtg_ddd_cel
                    EndEtg_tel_cel
                    */
                    if (!ddd_ok(f.EndEtg_ddd_res.value)) {
                        alert('Endere�o de entrega: DDD inv�lido!!');
                        f.EndEtg_ddd_res.focus();
                        return;
                    }
                    if (!telefone_ok(f.EndEtg_tel_res.value)) {
                        alert('Endere�o de entrega: telefone inv�lido!!');
                        f.EndEtg_tel_res.focus();
                        return;
                    }
                    if ((trim(f.EndEtg_ddd_res.value) != "") || (trim(f.EndEtg_tel_res.value) != "")) {
                        if (trim(f.EndEtg_ddd_res.value) == "") {
                            alert('Endere�o de entrega: preencha o DDD!!');
                            f.EndEtg_ddd_res.focus();
                            return;
                        }
                        if (trim(f.EndEtg_tel_res.value) == "") {
                            alert('Endere�o de entrega: preencha o telefone!!');
                            f.EndEtg_tel_res.focus();
                            return;
                        }
                    }

                    if (!ddd_ok(f.EndEtg_ddd_cel.value)) {
                        alert('Endere�o de entrega: DDD inv�lido!!');
                        f.EndEtg_ddd_cel.focus();
                        return;
                    }
                    if (!telefone_ok(f.EndEtg_tel_cel.value)) {
                        alert('Endere�o de entrega: telefone inv�lido!!');
                        f.EndEtg_tel_cel.focus();
                        return;
                    }
                    if ((f.EndEtg_ddd_cel.value == "") && (f.EndEtg_tel_cel.value != "")) {
                        alert('Endere�o de entrega: preencha o DDD do celular.');
                        f.EndEtg_tel_cel.focus();
                        return;
                    }
                    if ((f.EndEtg_tel_cel.value == "") && (f.EndEtg_ddd_cel.value != "")) {
                        alert('Endere�o de entrega: preencha o n�mero do celular.');
                        f.EndEtg_tel_cel.focus();
                        return;
                    }


                }
<%end if%>

			}
		}

	strMsgErro="";
	for (i=0; i < f.c_produto.length; i++) {
		if (trim(f.c_produto[i].value)!="") {
			if (f.c_preco_lista[i].style.color.toLowerCase()==COR_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__INEXISTENTE.toLowerCase()) {
				strMsgErro+="\n" + f.c_produto[i].value + " - " + f.c_descricao[i].value;
				}
			}
		}
	if (strMsgErro!="") {
		strMsgErro="A forma de pagamento " + KEY_ASPAS + f.c_custoFinancFornecParcelamentoDescricao.value.toLowerCase() + KEY_ASPAS + " n�o est� dispon�vel para o(s) produto(s):"+strMsgErro;
		alert(strMsgErro);
		return;
		}

    if (f.rb_etg_imediata[0].checked) {
        if (trim(f.c_data_previsao_entrega.value) == "") {
            alert("Informe a data de previs�o de entrega!");
            f.c_data_previsao_entrega.focus();
            return;
        }

        if (!isDate(f.c_data_previsao_entrega)) {
            alert("Data de previs�o de entrega � inv�lida!");
            f.c_data_previsao_entrega.focus();
            return;
        }

        if (retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(f.c_data_previsao_entrega.value)) <= retorna_so_digitos(formata_ddmmyyyy_yyyymmdd('<%=formata_data(Date)%>'))) {
            alert("Data de previs�o de entrega deve ser uma data futura!");
            f.c_data_previsao_entrega.focus();
            return;
        }
    }

    //campos do endere�o de entrega que precisam de transformacao
	transferirCamposEndEtg(f);

    

	f.action="OrcamentoAtualiza.asp";
	dCONFIRMA.style.visibility="hidden";
	window.status = "Aguarde ...";
	f.submit();
}

function transferirCamposEndEtg(formulario) {
<%if blnUsarMemorizacaoCompletaEnderecos and Not eh_cpf then%>
    //Transferimos os dados do endere�o de entrega dos campos certos. 
    //Temos dois conjuntos de campos (para PF e PJ) porque o layout � muito diferente.
    var pj = $('input[name="EndEtg_tipo_pessoa"]:checked').val() == "PJ";
    if (pj) {
        formulario.EndEtg_cnpj_cpf.value = formulario.EndEtg_cnpj_cpf_PJ.value;
        formulario.EndEtg_ie.value = formulario.EndEtg_ie_PJ.value;
        formulario.EndEtg_contribuinte_icms_status.value = $('input[name="EndEtg_contribuinte_icms_status_PJ"]:checked').val();
        if (!$('input[name="EndEtg_contribuinte_icms_status_PJ"]:checked').val())
            formulario.EndEtg_contribuinte_icms_status.value = "";
    }
    else {
        formulario.EndEtg_cnpj_cpf.value = formulario.EndEtg_cnpj_cpf_PF.value;
        formulario.EndEtg_ie.value = formulario.EndEtg_ie_PF.value;
        formulario.EndEtg_contribuinte_icms_status.value = $('input[name="EndEtg_contribuinte_icms_status_PF"]:checked').val();
        if (!$('input[name="EndEtg_contribuinte_icms_status_PF"]:checked').val())
            formulario.EndEtg_contribuinte_icms_status.value = "";
        formulario.EndEtg_produtor_rural_status.value = $('input[name="EndEtg_produtor_rural_status_PF"]:checked').val();
        if (!$('input[name="EndEtg_produtor_rural_status_PF"]:checked').val())
            formulario.EndEtg_produtor_rural_status.value = "";
	}
	

    //os campos a mais s�o enviados junto. Deixamos enviar...
<%end if%>
}

//para mudar o tipo do endere�o de entrega
function trocarEndEtgTipoPessoa(novoTipo) {
<%if blnUsarMemorizacaoCompletaEnderecos then%>
    if (novoTipo && $('input[name="EndEtg_tipo_pessoa"]:disabled').length == 0)
        setarValorRadio($('input[name="EndEtg_tipo_pessoa"]'), novoTipo);

    var pf = $('input[name="EndEtg_tipo_pessoa"]:checked').val() == "PF";

    //se nao tiver nada selecionado queremos tratar cono pj
    if (!pf) {
        $(".Mostrar_EndEtg_pf").css("display", "none");
        $(".Mostrar_EndEtg_pj").css("display", "");
        $("#Label_EndEtg_nome").text("RAZ�O SOCIAL");
    }
    else {
        //display block prejudica as tabelas
        $(".Mostrar_EndEtg_pf").css("display", "");
        $(".Mostrar_EndEtg_pj").css("display", "none");
        $("#Label_EndEtg_nome").text("NOME");
    }
<%else%>
    //oculta todos
    $(".Mostrar_EndEtg_pf").css("display", "none");
    $(".Mostrar_EndEtg_pj").css("display", "none");
    $(".Habilitar_EndEtg_outroendereco").css("display", "none");
<%end if%>
}

function trataContribuinteIcmsEndEtg_PJ(novoTipo)
{
    if (novoTipo && $('input[name="EndEtg_contribuinte_icms_status_PJ"]:disabled').length == 0)
        setarValorRadio($('input[name="EndEtg_contribuinte_icms_status_PJ"]'),novoTipo);
}
function trataContribuinteIcmsEndEtg_PF(novoTipo)
{
    if (novoTipo && $('input[name="EndEtg_contribuinte_icms_status_PF"]:disabled').length == 0)
        setarValorRadio($('input[name="EndEtg_contribuinte_icms_status_PF"]'),novoTipo);
}

function trataProdutorRural() {
    //ao clicar na op��o Produtor Rural, exibir/ocultar os campos apropriados
    if (!fORC.rb_produtor_rural[1].checked) {
        $("#t_contribuinte_icms").css("display", "none");
    }
    else {
        $("#t_contribuinte_icms").css("display", "block");
    }
}

function trataProdutorRuralEndEtg_PF(novoTipo) {
    //ao clicar na op��o Produtor Rural, exibir/ocultar os campos apropriados (endere�o de entrega)
    if (novoTipo && $('input[name="EndEtg_produtor_rural_status_PF"]:disabled').length == 0)
        setarValorRadio($('input[name="EndEtg_produtor_rural_status_PF"]'), novoTipo);

    var sim = $('input[name="EndEtg_produtor_rural_status_PF"]:checked').val() == "<%=COD_ST_CLIENTE_PRODUTOR_RURAL_SIM%>";

    //contribuinte ICMS sempre aparece para PJ
    if(sim) {
        $(".Mostrar_EndEtg_contribuinte_icms_PF").css("display", "");
    }
    else {
        $(".Mostrar_EndEtg_contribuinte_icms_PF").css("display", "none");
    }
}

function trataProdutorRuralEndEtg_PJ(novoTipo) {
    if (novoTipo && $('input[name="EndEtg_produtor_rural_status_PJ"]:disabled').length == 0)
        setarValorRadio($('input[name="EndEtg_produtor_rural_status_PJ"]'), novoTipo);
}

//definir um valor como ativo em um radio 
function setarValorRadio(array, valor)
{
    for (var i = 0; i < array.length; i++)
    {
        var este = array[i];
        if (este.value == valor)
            este.checked = true;
    }
}

</script>

<script type="text/javascript">
	function exibeJanelaCEP_Etg() {
		$.mostraJanelaCEP("EndEtg_cep", "EndEtg_uf", "EndEtg_cidade", "EndEtg_bairro", "EndEtg_endereco", "EndEtg_endereco_numero", "EndEtg_endereco_complemento");
	}

    function exibeJanelaCEP() {
        $.mostraJanelaCEP("endereco__cep", "endereco__uf", "endereco__cidade", "endereco__bairro", "endereco__endereco", "endereco__numero", "endereco__complemento");
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
<link href="<%=URL_FILE__E_JANELABUSCACEP_CSS%>" rel="stylesheet" type="text/css">
<link href="<%=URL_FILE__JQUERY_UI_CSS%>" rel="stylesheet" type="text/css">

<style TYPE="text/css">
#rb_etg_imediata, #rb_bem_uso_consumo {
	margin: 0pt 2pt 1pt 15pt;
	vertical-align: top;
	}
#rb_status {
	margin: 0pt 2pt 1pt 15pt;
	vertical-align: top;
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
</style>


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
<!-- ****************************************************** -->
<!-- **********  P�GINA PARA EDITAR O OR�AMENTO  ********** -->
<!-- ****************************************************** -->
<body id="corpoPagina" onload="processaFormaPagtoDefault();trataProdutorRural();">
<center>

<!-- #include file = "../global/JanelaBuscaCEP.htm"    -->

<form method="post" action="Pedido.asp" id="fPED" name="fPED">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="pedido_selecionado" id="pedido_selecionado" value=''>
</form>

<form id="fORC" name="fORC" method="post">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="orcamento_selecionado" id="orcamento_selecionado" value='<%=orcamento_selecionado%>'>
<input type="hidden" name="c_FlagEndEntregaEditavel" id="c_FlagEndEntregaEditavel" value='<%=strFlagEndEntregaEditavel%>'>
<input type="hidden" name="GarantiaIndicadorStatusOriginal" id="GarantiaIndicadorStatusOriginal" value='<%=r_orcamento.GarantiaIndicadorStatus%>'>
<input type="hidden" name="blnGarantiaIndicadorBloqueado" id="blnGarantiaIndicadorBloqueado" value='<%=Cstr(blnGarantiaIndicadorBloqueado)%>'>

<input type="hidden" name="FormaPagtoBloqueado" id="FormaPagtoBloqueado" value='<%=Cstr(blnFormaPagtoBloqueado)%>'>
<input type="hidden" name="c_loja" id="c_loja" value='<%=r_orcamento.loja%>'>
<input type="hidden" name="c_custoFinancFornecTipoParcelamentoOriginal" id="c_custoFinancFornecTipoParcelamentoOriginal" value='<%=r_orcamento.custoFinancFornecTipoParcelamento%>'>
<input type="hidden" name="c_custoFinancFornecQtdeParcelasOriginal" id="c_custoFinancFornecQtdeParcelasOriginal" value='<%=r_orcamento.custoFinancFornecQtdeParcelas%>'>
<input type="hidden" name="c_custoFinancFornecTipoParcelamento" id="c_custoFinancFornecTipoParcelamento" value='<%=r_orcamento.custoFinancFornecTipoParcelamento%>'>
<input type="hidden" name="c_custoFinancFornecQtdeParcelas" id="c_custoFinancFornecQtdeParcelas" value='<%=r_orcamento.custoFinancFornecQtdeParcelas%>'>
<input type="hidden" name="c_custoFinancFornecTipoParcelamentoUltConsulta" id="c_custoFinancFornecTipoParcelamentoUltConsulta" value='<%=r_orcamento.custoFinancFornecTipoParcelamento%>'>
<input type="hidden" name="c_custoFinancFornecQtdeParcelasUltConsulta" id="c_custoFinancFornecQtdeParcelasUltConsulta" value='<%=r_orcamento.custoFinancFornecQtdeParcelas%>'>
<input type="hidden" name="c_custoFinancFornecParcelamentoDescricao" id="c_custoFinancFornecParcelamentoDescricao" value=''>

<!-- AJAX EM ANDAMENTO -->
<div id="divAjaxRunning" style="display:none;"><img src="../Imagem/ajax_loader_gray_256.gif" class="AjaxImgLoader"/></div>


<!--  I D E N T I F I C A � � O   D O   O R � A M E N T O -->
<%=MontaHeaderIdentificacaoOrcamento(orcamento_selecionado, r_orcamento, 649)%>
<br>

<!--  L O J A   -->
<table width="649" class="Q" cellspacing="0">
	<tr>
<%	s = ""
	set r_loja = New cl_LOJA
	if x_loja_bd(r_orcamento.loja, r_loja) then
		with r_loja
			if Trim(.razao_social) <> "" then
				s = Trim(.razao_social)
			else
				s = Trim(.nome)
				end if
			end with
		end if
	strTextoIndicador = ""
	if r_orcamento.orcamentista <> "" then
		strTextoIndicador = r_orcamento.orcamentista
		if r_orcamentista_e_indicador.desempenho_nota <> "" then
			strTextoIndicador = strTextoIndicador & " (" & r_orcamentista_e_indicador.desempenho_nota & ")"
			end if
		end if
%>
	<td class="MD" align="left"><p class="Rf">LOJA</p><p class="C"><%=s%>&nbsp;</p></td>
	<td width="145" class="MD" align="left"><p class="Rf">OR�AMENTISTA</p><p class="C"><%=strTextoIndicador%>&nbsp;</p></td>
	<td width="145" align="left"><p class="Rf">VENDEDOR</p><p class="C"><%=r_orcamento.vendedor%>&nbsp;</p></td>
	</tr>
	</table>

<br>


<!--  ENDERE�O DO CLIENTE  -->
<% if r_orcamento.st_memorizacao_completa_enderecos = 0 then %>
	<!--  CLIENTE   -->
	<table width="649" class="Q" cellspacing="0">
	<tr>
<%	s = ""
	if xcliente_bd_resultado then
%>
<%	if cliente__tipo = ID_PF then s_aux="CPF" else s_aux="CNPJ"
	s = cnpj_cpf_formata(cliente__cnpj_cpf) 
%>
		<td align="left" width="50%" class="MD"><p class="Rf"><%=s_aux%></p>
		
			<p class="C"><%=s%>&nbsp;</p>
		
		</td>
		<%
		if cliente__tipo = ID_PF then s = Trim(cliente__rg) else s = Trim(cliente__ie)
			if cliente__tipo = ID_PF then 
%>
	<td align="left"><p class="Rf">RG</p><p class="C"><%=s%>&nbsp;</p></td>
<% else %>
	<td align="left"><p class="Rf">IE</p><p class="C"><%=s%>&nbsp;</p></td>
<% end if %>
		</tr>
	<%
		if Trim(cliente__nome) <> "" then
			s = Trim(cliente__nome)
			end if
		end if
	
	if cliente__tipo = ID_PF then s_aux="NOME DO CLIENTE" else s_aux="RAZ�O SOCIAL DO CLIENTE"
%>
    <tr>
	<td class="MC" align="left" colspan="2"><p class="Rf"><%=s_aux%></p>
	
		<p class="C"><%=s%>&nbsp;</p>
	
		</td>
	</tr>
	</table>
	
	<!--  ENDERE�O DO CLIENTE  -->
	<table width="649" class="QS" cellspacing="0">
	<tr>
<%
	s = formata_endereco(cliente__endereco, cliente__endereco_numero, cliente__endereco_complemento, cliente__bairro, cliente__cidade, cliente__uf, cliente__cep)
%>		
		<td align="left"><p class="Rf">ENDERE�O</p><p class="C"><%=s%>&nbsp;</p></td>
	</tr>
</table>
	
	<!--  TELEFONE DO CLIENTE  -->
	<table width="649" class="QS" cellspacing="0">
	<tr>
<%	s = ""
	if Trim(cliente__tel_res) <> "" then
		s = telefone_formata(Trim(cliente__tel_res))
		s_aux=Trim(cliente__ddd_res)
		if s_aux<>"" then s = "(" & s_aux & ") " & s
		end if
	
	s2 = ""
	if Trim(cliente__tel_com) <> "" then
		s2 = telefone_formata(Trim(cliente__tel_com))
		s_aux = Trim(cliente__ddd_com)
		if s_aux<>"" then s2 = "(" & s_aux & ") " & s2
		s_aux = Trim(cliente__ramal_com)
		if s_aux<>"" then s2 = s2 & "  (R. " & s_aux & ")"
		end if
	if Trim(cliente__tel_cel) <> "" then
		s3 = telefone_formata(Trim(cliente__tel_cel))
		s_aux = Trim(cliente__ddd_cel)
		if s_aux<>"" then s3 = "(" & s_aux & ") " & s3
		end if
	if Trim(cliente__tel_com_2) <> "" then
		s4 = telefone_formata(Trim(cliente__tel_com_2))
		s_aux = Trim(cliente__ddd_com_2)
		if s_aux<>"" then s4 = "(" & s_aux & ") " & s4
		s_aux = Trim(cliente__ramal_com_2)
		if s_aux<>"" then s4 = s4 & "  (R. " & s_aux & ")"
		end if
	
%>

<% if cliente__tipo = ID_PF then %>
	<td class="MD" width="33%" align="left"><p class="Rf">TELEFONE RESIDENCIAL</p><p class="C"><%=s%>&nbsp;</p></td>
	<td class="MD" width="33%" align="left"><p class="Rf">TELEFONE COMERCIAL</p><p class="C"><%=s2%>&nbsp;</p></td>
		<td align="left"><p class="Rf">CELULAR</p><p class="C"><%=s3%>&nbsp;</p></td>

<% else %>
	<td class="MD" width="50%" align="left"><p class="Rf">TELEFONE</p><p class="C"><%=s2%>&nbsp;</p></td>
	<td width="50%" align="left"><p class="Rf">TELEFONE</p><p class="C"><%=s4%>&nbsp;</p></td>

<% end if %>

	</tr>
</table>
	
	<!--  E-MAIL DO CLIENTE  -->
	<table width="649" class="QS" cellspacing="0">
	<tr>
		<td align="left"><p class="Rf">E-MAIL</p><p class="C"><%=Trim(cliente__email)%>&nbsp;</p></td>
	</tr>
</table>

<% else %>

			<!--  CLIENTE   -->
	<table width="649" class="Q" cellspacing="0">
	<tr>
<%	s = ""
	if xcliente_bd_resultado then
%>
<%	if cliente__tipo = ID_PF then s_aux="CPF" else s_aux="CNPJ"
	s = cnpj_cpf_formata(cliente__cnpj_cpf) 
%>
		<td align="left" class="MD" width="210"><p class="Rf"><%=s_aux%></p>
		
			<!--<p class="C"><%=s%>&nbsp;</p>-->
			<input id="endereco__cpf_cnpj" name="endereco__cpf_cnpj" readonly="readonly" class="TA" maxlength="72" style="width:310px;" value="<%=s%>">
		
		</td>
		<%
		if cliente__tipo = ID_PF then s = Trim(cliente__rg) else s = Trim(cliente__ie)
			if cliente__tipo = ID_PF then 
		%>
	<td align="left" class="MD"><p class="Rf">RG</p><input id="cliente__rg" name="cliente__rg" class="TA" maxlength="72" style="width:310px;" value="<%=s%>"></td>
	</tr>
	</table>


	<table width="649" class="QS" cellspacing="0">
		<tr>
			<td align="left"><p class="R">PRODUTOR RURAL</p><p class="C">
				<%s=cliente__produtor_rural_status%>
				<%if s = converte_numero(COD_ST_CLIENTE_PRODUTOR_RURAL_NAO) then s_aux="checked" else s_aux=""%>
				
				<input type="radio" id="rb_produtor_rural_nao" name="rb_produtor_rural" value="<%=COD_ST_CLIENTE_PRODUTOR_RURAL_NAO%>" <%=s_aux%> onclick="trataProdutorRural();"><span class="C" style="cursor:default" onclick="fORC.rb_produtor_rural[0].click();">N�o</span>
				<%if s = converte_numero(COD_ST_CLIENTE_PRODUTOR_RURAL_SIM) then s_aux="checked" else s_aux=""%>
				
				<input type="radio" id="rb_produtor_rural_sim" name="rb_produtor_rural" value="<%=COD_ST_CLIENTE_PRODUTOR_RURAL_SIM%>" <%=s_aux%> onclick="trataProdutorRural();"><span class="C" style="cursor:default" onclick="fORC.rb_produtor_rural[1].click();">Sim</span></p>
			</td>
		</tr>
	</table>


	

	<table width="649" class="QS" cellspacing="0" id="t_contribuinte_icms" onload="trataProdutorRural();">
		<tr>
			<%s=cliente__ie%>
			<td width="210" class="MD" align="left"><p class="R">IE</p><p class="C">
				<input id="cliente__ie" name="cliente__ie" class="TA" maxlength="72" style="width:310px;" value="<%=s%>" /></p>
			</td>
			<td align="left"><p class="R">CONTRIBUINTE ICMS</p><p class="C">
				<%s=cliente__icms%>
				<%if s = converte_numero(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO) then s_aux="checked" else s_aux=""%>
				<% intIdx = 0 %>
				<input type="radio" id="rb_contribuinte_icms_nao" name="rb_contribuinte_icms" value="<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO%>" <%=s_aux%>><span class="C" style="cursor:default" onclick="fORC.rb_contribuinte_icms[<%=Cstr(intIdx)%>].click();">N�o</span>
				<%if s = converte_numero(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM) then s_aux="checked" else s_aux=""%>
				<% intIdx = intIdx + 1 %>
				<input type="radio" id="rb_contribuinte_icms_sim" name="rb_contribuinte_icms" value="<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM%>" <%=s_aux%>><span class="C" style="cursor:default" onclick="fORC.rb_contribuinte_icms[<%=Cstr(intIdx)%>].click();">Sim</span>
				<%if s = converte_numero(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_ISENTO) then s_aux="checked" else s_aux=""%>
				<% intIdx = intIdx + 1 %>
				<input type="radio" id="rb_contribuinte_icms_isento" name="rb_contribuinte_icms" value="<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_ISENTO%>" <%=s_aux%>><span class="C" style="cursor:default" onclick="fORC.rb_contribuinte_icms[<%=Cstr(intIdx)%>].click();">Isento</span></p>
			</td>
		</tr>
	</table>
	

<% else %>
	<td class="MD" width="215" align="left"><p class="Rf">IE</p><input id="cliente__ie" name="cliente__ie" class="TA" maxlength="72" style="width:310px;" value="<%=s%>"></td>
	</tr>
	<tr>
		<td class="MC" align="left" colspan="2"><p class="R">CONTRIBUINTE ICMS</p><p class="C">

				<%
                    s = " "
                    if cliente__icms = converte_numero(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO) then
                        s = " checked "
                    end if
                %>
			
			<input type="radio" id="rb_contribuinte_icms_nao" name="rb_contribuinte_icms" value="<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO%>" <%=s%>><span class="C" style="cursor:default" onclick="fPED.rb_contribuinte_icms[1].click();">N�o</span>
				<%
                    s = " "
                    if cliente__icms = converte_numero(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM) then
                        s = " checked "
                    end if
                %>
			<input type="radio" id="rb_contribuinte_icms_sim" name="rb_contribuinte_icms" value="<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM%>" <%=s%>><span class="C" style="cursor:default" onclick="fPED.rb_contribuinte_icms[2].click();">Sim</span>
				<%
                    s = " "
                    if cliente__icms = converte_numero(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_ISENTO) then
                        s = " checked "
                    end if
                %>
			<input type="radio" id="rb_contribuinte_icms_isento" name="rb_contribuinte_icms" value="<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_ISENTO%>" <%=s%>><span class="C" style="cursor:default" onclick="fPED.rb_contribuinte_icms[3].click();">Isento</span></p>
			
		</td>
	</tr>
	</table>
<% end if %>
		
		
	<table width="649" class="QS" cellspacing="0">
		
	<%
		if Trim(cliente__nome) <> "" then
			s = Trim(cliente__nome)
			end if
		end if
	
	if cliente__tipo = ID_PF then s_aux="NOME DO CLIENTE" else s_aux="RAZ�O SOCIAL DO CLIENTE"
%>
    <tr>
	<td class="MD" align="left" colspan="2"><p class="Rf"><%=s_aux%></p>
	
		
		<input id="cliente__nome" name="cliente__nome" class="TA" value="<%=s%>" maxlength="60" style="width:635px;" />
				
	
		</td>
	</tr>
	</table>
	
	<!--  ENDERE�O DO CLIENTE  -->
	<table width="649" class="QS" cellspacing="0">
	    <tr>           
		    <td colspan="2" class="MB" align="left"><p class="Rf">ENDERE�O</p><input id="endereco__endereco" name="endereco__endereco" class="TA" maxlength="60" style="width:635px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fPED.EndEtg_endereco_numero.focus(); filtra_nome_identificador();" value="<%=cliente__endereco%>"></td>
	    </tr>
	    <tr>
		    <td class="MDB" align="left"><p class="Rf">N�</p><input id="endereco__numero" name="endereco__numero" class="TA" maxlength="20" style="width:310px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fPED.EndEtg_endereco_complemento.focus(); filtra_nome_identificador();" value="<%=cliente__endereco_numero%>"></td>
		    <td class="MB" align="left"><p class="Rf">COMPLEMENTO</p><input id="endereco__complemento" name="endereco__complemento" class="TA" maxlength="60" style="width:310px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fPED.EndEtg_bairro.focus(); filtra_nome_identificador();" value="<%=cliente__endereco_complemento%>"></td>
	    </tr>
	    <tr>
		    <td class="MDB" align="left"><p class="Rf">BAIRRO</p><input id="endereco__bairro" name="endereco__bairro" class="TA" maxlength="72" style="width:310px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fPED.EndEtg_cidade.focus(); filtra_nome_identificador();" value="<%=cliente__bairro%>"></td>
		    <td class="MB" align="left"><p class="Rf">CIDADE</p><input id="endereco__cidade" name="endereco__cidade" class="TA" maxlength="60" style="width:310px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fPED.EndEtg_uf.focus(); filtra_nome_identificador();" value="<%=cliente__cidade%>"></td>
	    </tr>
	    <tr>
		    <td width="50%" class="MD" align="left"><p class="Rf">UF</p><input id="endereco__uf" name="endereco__uf" class="TA" maxlength="2" size="3" onkeypress="if (digitou_enter(true) && tem_info(this.value) && uf_ok(this.value)) fPED.EndEtg_cep.focus();" onblur="this.value=trim(this.value); if (!uf_ok(this.value)) {alert('UF inv�lida!!');this.focus();} else this.value=ucase(this.value);" value="<%=cliente__uf%>"></td>
		    <td>
			    <table width="100%" cellspacing="0" cellpadding="0">
			    <tr>
			    <td width="50%" align="left"><p class="Rf">CEP</p><input id="endereco__cep" name="endereco__cep" readonly tabindex=-1 class="TA" maxlength="9" size="11" onkeypress="if (digitou_enter(true) && cep_ok(this.value)) filtra_cep();" onblur="if (!cep_ok(this.value)) {alert('CEP inv�lido!!');this.focus();} else this.value=cep_formata(this.value);" value='<%=cep_formata(cliente__cep)%>'></td>
			    <td align="center">
				    <% if blnPesquisaCEPAntiga then %>
				    <button type="button" name="bPesqCepEndEtg" id="bPesqCepEndEtg" style='width:130px;font-size:10pt;' class="Botao" onclick="AbrePesquisaCepEndEtg();">Pesquisar CEP</button>
				    <% end if %>
				    <% if blnPesquisaCEPAntiga and blnPesquisaCEPNova then Response.Write "&nbsp;" %>
				    <% if blnPesquisaCEPNova then %>
				    <button type="button" name="bPesqCepEndEtg" id="bPesqCepEndEtg" style='width:130px;font-size:10pt;' class="Botao" onclick="exibeJanelaCEP();">Pesquisar CEP</button>
				    <% end if %>
				    <a name="bLimparEndEtg" id="bLimparEndEtg" href="javascript:LimparCamposEndEtg(fPED)" title="limpa o endere�o de entrega">
					    <img src="../botao/botao_x_red.gif" style="vertical-align:bottom;margin-bottom:1px;" width="20" height="20" border="0"></a>
			    </td>
			    </tr>
			    </table>
		    </td>
	    </tr>
    </table>

		<% if cliente__tipo = ID_PF then %>
			<!--  TELEFONE DO CLIENTE  -->
			<table width="649" class="QS" cellspacing="0">
	            <tr>
					<td class="MD" width="20%" align="left"><p class="R">DDD</p><p class="C">
						<input id="cliente__ddd_res" name="cliente__ddd_res" class="TA" value="<%=cliente__ddd_res%>" maxlength="4" size="5" onkeypress="if (digitou_enter(true) && ddd_ok(this.value)) fPED.EndEtg_tel_res.focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inv�lido!!');this.focus();}"></p>
					</td>
					<td class="MD" align="left"><p class="R">TELEFONE RESIDENCIAL</p><p class="C">
						<input id="cliente__tel_res" name="cliente__tel_res" class="TA" value="<%=telefone_formata(cliente__tel_res)%>" maxlength="11" size="12" onkeypress="if (digitou_enter(true) && telefone_ok(this.value)) fPED.EndEtg_ddd_cel.focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('Telefone inv�lido!!');this.focus();} else this.value=telefone_formata(this.value);"></p>
					</td>
	            </tr>
			</table>
			<table width="649" class="QS" cellspacing="0">
				<tr>
					<td class="MD" width="20%" align="left"><p class="R">DDD</p><p class="C">
						<input id="cliente__ddd_cel" name="cliente__ddd_cel" class="TA" value="<%=cliente__ddd_cel%>" maxlength="4" size="5" onkeypress="if (digitou_enter(true) && ddd_ok(this.value)) fPED.EndEtg_tel_cel.focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inv�lido!!');this.focus();}"></p>
					</td>
					<td align="left" class="MD"><p class="R">CELULAR</p><p class="C">
						<input id="cliente__tel_cel" name="cliente__tel_cel" class="TA" value="<%=telefone_formata(cliente__tel_cel)%>" maxlength="9" size="12" onkeypress="if (digitou_enter(true) && telefone_ok(this.value)) fPED.EndEtg_obs.focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('N�mero de celular inv�lido!!');this.focus();} else this.value=telefone_formata(this.value);"></p>
					</td>
	            </tr>
			</table>
			<table width="649" class="QS" cellspacing="0">
	            <tr>
					<td class="MD" width="20%" align="left"><p class="R">DDD</p><p class="C">
						<input id="cliente__ddd_com" name="cliente__ddd_com" class="TA" value="<%=cliente__ddd_com%>" maxlength="4" size="5" onkeypress="if (digitou_enter(true) && ddd_ok(this.value)) fPED.EndEtg_tel_cel.focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inv�lido!!');this.focus();}"></p>
					</td>
					<td class="MD" align="left"><p class="R">COMERCIAL</p><p class="C">
						<input id="cliente__tel_com" name="cliente__tel_com" class="TA" value="<%=telefone_formata(cliente__tel_com)%>" maxlength="9" size="12" onkeypress="if (digitou_enter(true) && telefone_ok(this.value)) fPED.EndEtg_obs.focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('N�mero de telefone comercial inv�lido!!');this.focus();} else this.value=telefone_formata(this.value);"></p>
					</td>
					<td align="left"><p class="R">RAMAL</p><p class="C">
						<input id="cliente__ramal_com" name="cliente__ramal_com" class="TA" value="<%=cliente__ramal_com%>" maxlength="4" size="6" onkeypress="if (digitou_enter(true)) fPED.EndEtg_ddd_com_2.focus(); filtra_numerico();"></p>
					</td>
				</tr>
				
			</table>	

		<% else %>
			<!--  TELEFONE DO CLIENTE  -->
			<table width="649" class="QS" cellspacing="0">
	            <tr>
					<td class="MD" width="20%" align="left"><p class="R">DDD</p><p class="C">
						<input id="cliente__ddd_com" name="cliente__ddd_com" class="TA" value="<%=cliente__ddd_com%>" maxlength="4" size="5" onkeypress="if (digitou_enter(true) && ddd_ok(this.value)) fPED.EndEtg_tel_com.focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inv�lido!!');this.focus();}"></p></td>
					<td class="MD" align="left"><p class="R">TELEFONE </p><p class="C">
						<input id="cliente__tel_com" name="cliente__tel_com" class="TA" value="<%=telefone_formata(cliente__tel_com)%>" maxlength="11" size="12" onkeypress="if (digitou_enter(true) && telefone_ok(this.value)) fPED.EndEtg_ramal_com.focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('Telefone inv�lido!!');this.focus();} else this.value=telefone_formata(this.value);"></p></td>
					<td align="left"><p class="R">RAMAL</p><p class="C">
						<input id="cliente__ramal_com" name="cliente__ramal_com" class="TA" value="<%=cliente__ramal_com%>" maxlength="4" size="6" onkeypress="if (digitou_enter(true)) fPED.EndEtg_ddd_com_2.focus(); filtra_numerico();"></p>
					</td>
	            </tr>
	            <tr>
	                <td class="MD MC" width="20%" align="left"><p class="R">DDD</p><p class="C">
						<input id="cliente__ddd_com_2" name="cliente__ddd_com_2" class="TA" value="<%=cliente__ddd_com_2%>" maxlength="4" size="5" onkeypress="if (digitou_enter(true) && ddd_ok(this.value)) fPED.EndEtg_tel_com_2.focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inv�lido!!!');this.focus();}" /></p>  
	                </td>
	                <td class="MD MC" align="left"><p class="R">TELEFONE</p><p class="C">
						<input id="cliente__tel_com_2" name="cliente__tel_com_2" class="TA" value="<%=telefone_formata(cliente__tel_com_2)%>" maxlength="9" size="12" onkeypress="if (digitou_enter(true) && telefone_ok(this.value)) fPED.EndEtg_ramal_com_2.focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('Telefone inv�lido!!');this.focus();} else this.value=telefone_formata(this.value);"></p>
	                </td>
	                <td align="left" class="MC"><p class="R">RAMAL</p><p class="C">
						<input id="cliente__ramal_com_2" name="cliente__ramal_com_2" class="TA" value="<%=cliente__ramal_com_2%>" maxlength="4" size="6" onkeypress="if (digitou_enter(true)) fPED.EndEtg_obs.focus(); filtra_numerico();" /></p>
	                </td>
	            </tr>
            </table>
		<% end if %>

	<!--  E-MAIL DO CLIENTE  -->
	<table width="649" class="QS" cellspacing="0">
		 <tr>           
		    <td colspan="2" class="Rf" align="left"><p class="Rf">E-MAIL</p>
				<input id="cliente__email" name="cliente__email" class="TA" maxlength="60" style="width:635px;" value="<%=cliente__email%>" onkeypress="Sfiltra_email();" />

		    </td>
	    </tr>
	</table>

	 <!-- ************   E-MAIL (XML)  ************ -->
    <table width="649" class="QS" cellspacing="0">
	    <tr>
	    <td width="100%" align="left"><p class="R">E-MAIL (XML)</p><p class="C">
		    <input id="cliente__email_xml" name="cliente__email_xml" value="<%=cliente__email_xml%>" class="TA" maxlength="60" size="74" onkeypress="if (digitou_enter(true)) fORC.rb_end_entrega_nao.focus(); filtra_email();"></p></td>
	    </tr>
    </table>

<%end if%>


<% if strFlagEndEntregaEditavel = "N" then %>
<!--  ENDERE�O DE ENTREGA  -->
<%	
	s = pedido_formata_endereco_entrega(r_orcamento, r_cliente)
%>		
<table width="649" class="QS" cellspacing="0" style="table-layout:fixed">
	<tr>
		<td align="left"><p class="Rf">ENDERE�O DE ENTREGA</p><p class="C"><%=s%>&nbsp;</p></td>
	</tr>
     <%	if r_orcamento.EndEtg_cod_justificativa <> "" then %>	
    <tr>
		<td align="left" style="word-wrap:break-word"><p class="C"><%=obtem_descricao_tabela_t_codigo_descricao(GRUPO_T_CODIGO_DESCRICAO__ENDETG_JUSTIFICATIVA,r_orcamento.EndEtg_cod_justificativa)%>&nbsp;</p></td>
	</tr>
    <%end if %>
</table>
<% else %>


<% if blnUsarMemorizacaoCompletaEnderecos then %>
    <!--  ************  TIPO DO ENDERE�O DE ENTREGA: PF/PJ (SOMENTE SE O CLIENTE FOR PJ)   ************ -->

    <%if eh_cpf then%>
        <!-- ************   ENDERE�O DE ENTREGA PARA CLIENTE PF   ************ -->
        <!-- Pegamos todos os atuais. Sem campos edit�veis. Pegamos os atuais do cadastro do cliente, n�o do pedido em si. -->
        <input type="hidden" id="EndEtg_tipo_pessoa" name="EndEtg_tipo_pessoa" value="PF"/>
        <input type="hidden" id="EndEtg_cnpj_cpf" name="EndEtg_cnpj_cpf" value="<%=r_cliente.cnpj_cpf%>"/>
        <input type="hidden" id="EndEtg_ie" name="EndEtg_ie" value="<%=r_cliente.ie%>"/>
        <input type="hidden" id="EndEtg_contribuinte_icms_status" name="EndEtg_contribuinte_icms_status" value="<%=r_cliente.contribuinte_icms_status%>"/>
        <input type="hidden" id="EndEtg_rg" name="EndEtg_rg" value="<%=r_cliente.rg%>"/>
        <input type="hidden" id="EndEtg_produtor_rural_status" name="EndEtg_produtor_rural_status" value="<%=r_cliente.produtor_rural_status%>"/>
        <input type="hidden" id="EndEtg_email" name="EndEtg_email" value="<%=r_cliente.email%>"/>
        <input type="hidden" id="EndEtg_email_xml" name="EndEtg_email_xml" value="<%=r_cliente.email_xml%>"/>
        <input type="hidden" id="EndEtg_nome" name="EndEtg_nome" value="<%=r_cliente.nome%>"/>


    <%else%>
        <table width="649" class="QS Habilitar_EndEtg_outroendereco" cellspacing="0">
	        <tr>
		        <td align="left">
		        <p class="R">ENDERE�O DE ENTREGA</p><p class="C">
                    <%
                        s = " "
                        if r_orcamento.EndEtg_tipo_pessoa = ID_PJ then
                            s = " checked "
                        end if
                    %>
			        <input type="radio" id="EndEtg_tipo_pessoa_PJ" name="EndEtg_tipo_pessoa" value="PJ" onclick="trocarEndEtgTipoPessoa(null);" <%=s%> >
			        <span class="C" style="cursor:default" onclick="trocarEndEtgTipoPessoa('PJ');">Pessoa Jur�dica</span>
			        &nbsp;
                    <%
                        s = " "
                        if r_orcamento.EndEtg_tipo_pessoa = ID_PF then
                            s = " checked "
                        end if
                    %>
			        <input type="radio" id="EndEtg_tipo_pessoa_PF" name="EndEtg_tipo_pessoa" value="PF" onclick="trocarEndEtgTipoPessoa(null);" <%=s%> >
			        <span class="C" style="cursor:default" onclick="trocarEndEtgTipoPessoa('PF');">Pessoa F�sica</span>
		        </p>
		        </td>
	        </tr>
        </table>

                <!-- ************   PJ: CNPJ/CONTRIBUINTE ICMS/IE - DO ENDERE�O DE ENTREGA DE PJ ************ -->
                <!-- ************   PF: CPF/PRODUTOR RURAL/CONTRIBUINTE ICMS/IE - DO ENDERE�O DE ENTREGA DE PJ  ************ -->
                <!-- fizemos dois conjuntos diferentes de campos porque a ordem � muito diferente -->
                <!-- EndEtg_rg EndEtg_email e EndEtg_email_xml vem diretamente do t_CLIENTE -->

        <input type="hidden" id="EndEtg_cnpj_cpf" name="EndEtg_cnpj_cpf" />
        <input type="hidden" id="EndEtg_ie" name="EndEtg_ie" />
        <input type="hidden" id="EndEtg_contribuinte_icms_status" name="EndEtg_contribuinte_icms_status" />
        <input type="hidden" id="EndEtg_rg" name="EndEtg_rg" value="<%=r_cliente.rg%>"/>
        <input type="hidden" id="EndEtg_produtor_rural_status" name="EndEtg_produtor_rural_status" />
        <input type="hidden" id="EndEtg_email" name="EndEtg_email" value="<%=r_cliente.email%>"/>
        <input type="hidden" id="EndEtg_email_xml" name="EndEtg_email_xml" value="<%=r_cliente.email_xml%>"/>


        <table width="649" class="QS Habilitar_EndEtg_outroendereco Mostrar_EndEtg_pj" cellspacing="0">
	        <tr>
		        <td width="210" align="left">
	        <p class="R">CNPJ</p><p class="C">

	        <input id="EndEtg_cnpj_cpf_PJ" name="EndEtg_cnpj_cpf_PJ" class="TA" value="<%=r_orcamento.EndEtg_cnpj_cpf%>" size="22" style="text-align:center; color:#0000ff"></p></td>

	        <td class="MDE" width="215" align="left"><p class="R">IE</p><p class="C">
		        <input id="EndEtg_ie_PJ" name="EndEtg_ie_PJ" class="TA" type="text" maxlength="20" size="25" value="<%=r_orcamento.EndEtg_ie%>" onkeypress="if (digitou_enter(true)) fORC.EndEtg_nome.focus(); filtra_nome_identificador();"></p>

	        </td>

	        <td align="left" class="Mostrar_EndEtg_contribuinte_icms_PJ"><p class="R">CONTRIBUINTE ICMS</p><p class="C">
                <%
                    s = " "
                    if r_orcamento.EndEtg_contribuinte_icms_status = converte_numero(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO) then
                        s = " checked "
                    end if
                %>
		        <input type="radio"  <%=s%> id="EndEtg_contribuinte_icms_status_PJ_nao" name="EndEtg_contribuinte_icms_status_PJ" value="<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO%>" ><span class="C" style="cursor:default" onclick="trataContribuinteIcmsEndEtg_PJ('<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO%>');">N�o</span>
                <%
                    s = " "
                    if r_orcamento.EndEtg_contribuinte_icms_status = converte_numero(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM) then
                        s = " checked "
                    end if
                %>
		        <input type="radio"  <%=s%> id="EndEtg_contribuinte_icms_status_PJ_sim" name="EndEtg_contribuinte_icms_status_PJ" value="<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM%>" ><span class="C" style="cursor:default" onclick="trataContribuinteIcmsEndEtg_PJ('<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM%>');">Sim</span>
                <%
                    s = " "
                    if r_orcamento.EndEtg_contribuinte_icms_status = converte_numero(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_ISENTO) then
                        s = " checked "
                    end if
                %>
		        <input type="radio"  <%=s%> id="EndEtg_contribuinte_icms_status_PJ_isento" name="EndEtg_contribuinte_icms_status_PJ" value="<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_ISENTO%>" ><span class="C" style="cursor:default" onclick="trataContribuinteIcmsEndEtg_PJ('<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_ISENTO%>');">Isento</span></p></td>
	        </tr>
        </table>

        <table width="649" class="QS Habilitar_EndEtg_outroendereco Mostrar_EndEtg_pf" cellspacing="0">
	        <tr>
		        <td width="210" align="left">
	        <p class="R">CPF</p><p class="C">
	        <input id="EndEtg_cnpj_cpf_PF" name="EndEtg_cnpj_cpf_PF" class="TA" value="<%=r_orcamento.EndEtg_cnpj_cpf%>" size="22" style="text-align:center; color:#0000ff"></p></td>

	        <td align="left" class="ME" style="min-width: 110px;" ><p class="R">PRODUTOR RURAL</p><p class="C">
                <%
                    s = " "
                    if r_orcamento.EndEtg_produtor_rural_status = converte_numero(COD_ST_CLIENTE_PRODUTOR_RURAL_NAO) then
                        s = " checked "
                    end if
                %>
		        <input type="radio"  <%=s%> id="EndEtg_produtor_rural_status_PF_nao" name="EndEtg_produtor_rural_status_PF" value="<%=COD_ST_CLIENTE_PRODUTOR_RURAL_NAO%>" onclick="trataProdutorRuralEndEtg_PF(null);"><span class="C" style="cursor:default" onclick="trataProdutorRuralEndEtg_PF('<%=COD_ST_CLIENTE_PRODUTOR_RURAL_NAO%>');">N�o</span>
                <%
                    s = " "
                    if r_orcamento.EndEtg_produtor_rural_status = converte_numero(COD_ST_CLIENTE_PRODUTOR_RURAL_SIM) then
                        s = " checked "
                    end if
                %>
		        <input type="radio"  <%=s%> id="EndEtg_produtor_rural_status_PF_sim" name="EndEtg_produtor_rural_status_PF" value="<%=COD_ST_CLIENTE_PRODUTOR_RURAL_SIM%>" onclick="trataProdutorRuralEndEtg_PF(null);"><span class="C" style="cursor:default" onclick="trataProdutorRuralEndEtg_PF('<%=COD_ST_CLIENTE_PRODUTOR_RURAL_SIM%>')">Sim</span></p></td>

	        <td align="left" class="MDE Mostrar_EndEtg_contribuinte_icms_PF"><p class="R">IE</p><p class="C">
		        <input id="EndEtg_ie_PF" name="EndEtg_ie_PF" class="TA" type="text" maxlength="20" size="13" value="<%=r_orcamento.EndEtg_ie%>" onkeypress="if (digitou_enter(true)) fORC.EndEtg_nome.focus(); filtra_nome_identificador();"></p>
	        </td>

	        <td align="left" class="Mostrar_EndEtg_contribuinte_icms_PF" ><p class="R">CONTRIBUINTE ICMS</p><p class="C">
                <%
                    s = " "
                    if r_orcamento.EndEtg_contribuinte_icms_status = converte_numero(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO) then
                        s = " checked "
                    end if
                %>
		        <input type="radio"  <%=s%> id="EndEtg_contribuinte_icms_status_PF_nao" name="EndEtg_contribuinte_icms_status_PF" value="<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO%>" ><span class="C" style="cursor:default" onclick="trataContribuinteIcmsEndEtg_PF('<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO%>');">N�o</span>
                <%
                    s = " "
                    if r_orcamento.EndEtg_contribuinte_icms_status = converte_numero(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM) then
                        s = " checked "
                    end if
                %>
		        <input type="radio"  <%=s%> id="EndEtg_contribuinte_icms_status_PF_sim" name="EndEtg_contribuinte_icms_status_PF" value="<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM%>" ><span class="C" style="cursor:default" onclick="trataContribuinteIcmsEndEtg_PF('<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM%>');">Sim</span>
                <%
                    s = " "
                    if r_orcamento.EndEtg_contribuinte_icms_status = converte_numero(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_ISENTO) then
                        s = " checked "
                    end if
                %>
		        <input type="radio"  <%=s%> id="EndEtg_contribuinte_icms_status_PF_isento" name="EndEtg_contribuinte_icms_status_PF" value="<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_ISENTO%>" ><span class="C" style="cursor:default" onclick="trataContribuinteIcmsEndEtg_PF('<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_ISENTO%>');">Isento</span></p>
	        </td>
	        </tr>
        </table>



        <!-- ************   ENDERE�O DE ENTREGA: NOME  ************ -->
        <table width="649" class="QS" cellspacing="0">
	        <tr>
	        <td width="100%" align="left"><p class="R" id="Label_EndEtg_nome">RAZ�O SOCIAL</p><p class="C">
		        <input id="EndEtg_nome" name="EndEtg_nome" class="TA" value="<%=r_orcamento.EndEtg_nome%>" maxlength="60" size="85" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fORC.EndEtg_endereco.focus(); filtra_nome_identificador();"></p></td>
	        </tr>
        </table>

    <%end if%>
<%end if%> <% 'blnUsarMemorizacaoCompletaEnderecos %>

<table width="649" class="QS" cellspacing="0">
	<tr>
        <%
            s = "ENDERE�O"
            if eh_cpf then
                s = "ENDERE�O DE ENTREGA"
                end if
            %>
		<td colspan="2" class="MB" align="left"><p class="Rf"><%=s%></p><input id="EndEtg_endereco" name="EndEtg_endereco" class="TA" maxlength="60" style="width:635px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fORC.EndEtg_endereco_numero.focus(); filtra_nome_identificador();" value="<%=r_orcamento.EndEtg_endereco%>"></td>
	</tr>
	<tr>
		<td class="MDB" align="left"><p class="Rf">N�</p><input id="EndEtg_endereco_numero" name="EndEtg_endereco_numero" class="TA" maxlength="20" style="width:310px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fORC.EndEtg_endereco_complemento.focus(); filtra_nome_identificador();" value="<%=r_orcamento.EndEtg_endereco_numero%>"></td>
		<td class="MB" align="left"><p class="Rf">COMPLEMENTO</p><input id="EndEtg_endereco_complemento" name="EndEtg_endereco_complemento" class="TA" maxlength="60" style="width:310px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fORC.EndEtg_bairro.focus(); filtra_nome_identificador();" value="<%=r_orcamento.EndEtg_endereco_complemento%>"></td>
	</tr>
	<tr>
		<td class="MDB" align="left"><p class="Rf">BAIRRO</p><input id="EndEtg_bairro" name="EndEtg_bairro" class="TA" maxlength="72" style="width:310px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fORC.EndEtg_cidade.focus(); filtra_nome_identificador();" value="<%=r_orcamento.EndEtg_bairro%>"></td>
		<td class="MB" align="left"><p class="Rf">CIDADE</p><input id="EndEtg_cidade" name="EndEtg_cidade" class="TA" maxlength="60"  style="width:310px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fORC.EndEtg_uf.focus(); filtra_nome_identificador();" value="<%=r_orcamento.EndEtg_cidade%>"></td>
	</tr>
	<tr>
		<td width="50%" class="MD" align="left"><p class="Rf">UF</p><input id="EndEtg_uf" name="EndEtg_uf" class="TA" maxlength="2" size="3" onkeypress="if (digitou_enter(true) && tem_info(this.value) && uf_ok(this.value)) fORC.EndEtg_cep.focus();" onblur="this.value=trim(this.value); if (!uf_ok(this.value)) {alert('UF inv�lida!!');this.focus();} else this.value=ucase(this.value);" value="<%=r_orcamento.EndEtg_uf%>"></td>
		<td align="left">
			<table width="100%" cellspacing="0" cellpadding="0">
			<tr>
			<td width="50%" align="left"><p class="Rf">CEP</p><input id="EndEtg_cep" name="EndEtg_cep" readonly tabindex=-1 class="TA" maxlength="9" size="11" onkeypress="if (digitou_enter(true) && cep_ok(this.value)) filtra_cep();" onblur="if (!cep_ok(this.value)) {alert('CEP inv�lido!!');this.focus();} else this.value=cep_formata(this.value);" value='<%=cep_formata(r_orcamento.EndEtg_cep)%>'></td>
			<td align="center">
				<% if blnPesquisaCEPAntiga then %>
				<button type="button" name="bPesqCepEndEtg" id="bPesqCepEndEtg" style='width:130px;font-size:10pt;' class="Botao" onclick="AbrePesquisaCepEndEtg();">Pesquisar CEP</button>
				<% end if %>
				<% if blnPesquisaCEPAntiga and blnPesquisaCEPNova then Response.Write "&nbsp;" %>
				<% if blnPesquisaCEPNova then %>
				<button type="button" name="bPesqCepEndEtg" id="bPesqCepEndEtg" style='width:130px;font-size:10pt;' class="Botao" onclick="exibeJanelaCEP_Etg();">Pesquisar CEP</button>
				<% end if %>
				<a name="bLimparEndEtg" id="bLimparEndEtg" href="javascript:LimparCamposEndEtg(fORC)" title="limpa o endere�o de entrega">
					<img src="../botao/botao_x_red.gif" style="vertical-align:bottom;margin-bottom:1px;" width="20" height="20" border="0"></a>
			</td>
			</tr>
			</table>
		</td>
	</tr>
    </table>


    <% if blnUsarMemorizacaoCompletaEnderecos then %>
        <%if eh_cpf then%>

            <!-- ************   ENDERE�O DE ENTREGA PARA PF: TELEFONES   ************ -->
            <!-- Pegamos todos os atuais. Sem campos edit�veis. Pegamos os atuais do cadastro do cliente, n�o do pedido em si. -->
            <input type="hidden" id="EndEtg_ddd_res" name="EndEtg_ddd_res" value="<%=r_cliente.ddd_res%>"/>
            <input type="hidden" id="EndEtg_tel_res" name="EndEtg_tel_res" value="<%=r_cliente.tel_res%>"/>
            <input type="hidden" id="EndEtg_ddd_cel" name="EndEtg_ddd_cel" value="<%=r_cliente.ddd_cel%>"/>
            <input type="hidden" id="EndEtg_tel_cel" name="EndEtg_tel_cel" value="<%=r_cliente.tel_cel%>"/>
            <input type="hidden" id="EndEtg_ddd_com" name="EndEtg_ddd_com" value="<%=r_cliente.ddd_com%>"/>
            <input type="hidden" id="EndEtg_tel_com" name="EndEtg_tel_com" value="<%=r_cliente.tel_com%>"/>
            <input type="hidden" id="EndEtg_ramal_com" name="EndEtg_ramal_com" value="<%=r_cliente.ramal_com%>"/>
            <input type="hidden" id="EndEtg_ddd_com_2" name="EndEtg_ddd_com_2" value="<%=r_cliente.ddd_com_2%>"/>
            <input type="hidden" id="EndEtg_tel_com_2" name="EndEtg_tel_com_2" value="<%=r_cliente.tel_com_2%>"/>
            <input type="hidden" id="EndEtg_ramal_com_2" name="EndEtg_ramal_com_2" value="<%=r_cliente.ramal_com_2%>"/>

        <%else%>
        
            <!-- ************   ENDERE�O DE ENTREGA: TELEFONE RESIDENCIAL   ************ -->
            <table width="649" class="QS Mostrar_EndEtg_pf Habilitar_EndEtg_outroendereco" cellspacing="0">
	            <tr>
	            <td class="MD" width="20%" align="left"><p class="R">DDD</p><p class="C">
		            <input id="EndEtg_ddd_res" name="EndEtg_ddd_res" class="TA" value="<%=r_orcamento.EndEtg_ddd_res%>" maxlength="4" size="5" onkeypress="if (digitou_enter(true) && ddd_ok(this.value)) fORC.EndEtg_tel_res.focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inv�lido!!');this.focus();}"></p></td>
	            <td align="left"><p class="R">TELEFONE RESIDENCIAL</p><p class="C">
		            <input id="EndEtg_tel_res" name="EndEtg_tel_res" class="TA" value="<%=r_orcamento.EndEtg_tel_res%>" maxlength="11" size="12" onkeypress="if (digitou_enter(true) && telefone_ok(this.value)) fORC.EndEtg_ddd_cel.focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('Telefone inv�lido!!');this.focus();} else this.value=telefone_formata(this.value);"></p></td>
	            </tr>
	            <tr>
	            <td class="MD MC" width="20%" align="left"><p class="R">DDD</p><p class="C">
		            <input id="EndEtg_ddd_cel" name="EndEtg_ddd_cel" class="TA" value="<%=r_orcamento.EndEtg_ddd_cel%>" maxlength="4" size="5" onkeypress="if (digitou_enter(true) && ddd_ok(this.value)) fORC.EndEtg_tel_cel.focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inv�lido!!');this.focus();}"></p></td>
	            <td align="left" class="MC"><p class="R">CELULAR</p><p class="C">
		            <input id="EndEtg_tel_cel" name="EndEtg_tel_cel" class="TA" value="<%=r_orcamento.EndEtg_tel_cel%>" maxlength="9" size="12" onkeypress="if (digitou_enter(true) && telefone_ok(this.value)) fORC.EndEtg_obs.focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('N�mero de celular inv�lido!!');this.focus();} else this.value=telefone_formata(this.value);"></p></td>
	            </tr>
            </table>
	
            <!-- ************   ENDERE�O DE ENTREGA: TELEFONE COMERCIAL   ************ -->
            <table width="649" class="QS Mostrar_EndEtg_pj Habilitar_EndEtg_outroendereco" cellspacing="0">
	            <tr>
	            <td class="MD" width="20%" align="left"><p class="R">DDD</p><p class="C">
		            <input id="EndEtg_ddd_com" name="EndEtg_ddd_com" class="TA" value="<%=r_orcamento.EndEtg_ddd_com%>" maxlength="4" size="5" onkeypress="if (digitou_enter(true) && ddd_ok(this.value)) fORC.EndEtg_tel_com.focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inv�lido!!');this.focus();}"></p></td>
	            <td class="MD" align="left"><p class="R">TELEFONE </p><p class="C">
		            <input id="EndEtg_tel_com" name="EndEtg_tel_com" class="TA" value="<%=r_orcamento.EndEtg_tel_com%>" maxlength="11" size="12" onkeypress="if (digitou_enter(true) && telefone_ok(this.value)) fORC.EndEtg_ramal_com.focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('Telefone inv�lido!!');this.focus();} else this.value=telefone_formata(this.value);"></p></td>
	            <td align="left"><p class="R">RAMAL</p><p class="C">
		            <input id="EndEtg_ramal_com" name="EndEtg_ramal_com" class="TA" value="<%=r_orcamento.EndEtg_ramal_com%>" maxlength="4" size="6" onkeypress="if (digitou_enter(true)) fORC.EndEtg_ddd_com_2.focus(); filtra_numerico();"></p></td>
	            </tr>
	            <tr>
	                <td class="MD MC" width="20%" align="left"><p class="R">DDD</p><p class="C">
	                <input id="EndEtg_ddd_com_2" name="EndEtg_ddd_com_2" class="TA" value="<%=r_orcamento.EndEtg_ddd_com_2%>" maxlength="4" size="5" onkeypress="if (digitou_enter(true) && ddd_ok(this.value)) fORC.EndEtg_tel_com_2.focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inv�lido!!!');this.focus();}" /></p>  
	                </td>
	                <td class="MD MC" align="left"><p class="R">TELEFONE</p><p class="C">
	                <input id="EndEtg_tel_com_2" name="EndEtg_tel_com_2" class="TA" value="<%=r_orcamento.EndEtg_tel_com_2%>" maxlength="9" size="12" onkeypress="if (digitou_enter(true) && telefone_ok(this.value)) fORC.EndEtg_ramal_com_2.focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('Telefone inv�lido!!');this.focus();} else this.value=telefone_formata(this.value);"></p>
	                </td>
	                <td align="left" class="MC"><p class="R">RAMAL</p><p class="C">
	                <input id="EndEtg_ramal_com_2" name="EndEtg_ramal_com_2" class="TA" value="<%=r_orcamento.EndEtg_ramal_com_2%>" maxlength="4" size="6" onkeypress="if (digitou_enter(true)) fORC.EndEtg_obs.focus(); filtra_numerico();" /></p>
	                </td>
	            </tr>
            </table>

        <% end if %>
    <%end if%> <% 'blnUsarMemorizacaoCompletaEnderecos %>


    <!-- ************   JUSTIFIQUE O ENDERE�O   ************ -->
    <table id="obs_endereco" width="649" class="QS" cellspacing="0">
	    <tr >
            <td colspan="2" class="MC" align="left"><p class="Rf">JUSTIFIQUE O ENDERE�O</p><p class="C">
            <select id="EndEtg_obs" name="EndEtg_obs" style="margin-right:225px;">			
			 <%=codigo_descricao_monta_itens_select_all(GRUPO_T_CODIGO_DESCRICAO__ENDETG_JUSTIFICATIVA, "")%>
		   </select></p></td>
	</tr>
</table>
	

<% end if %>
<!--  R E L A � � O   D E   P R O D U T O S  -->
<br>
<br>
<table class="Qx" cellspacing="0">
	<tr bgColor="#FFFFFF">
	<td class="MB" align="left" valign="bottom"><span class="PLTe">Fabr</span></td>
	<td class="MB" align="left" valign="bottom"><span class="PLTe">Produto</span></td>
	<td class="MB" align="left" valign="bottom"><span class="PLTe">Descri��o</span></td>
	<td class="MB" align="left" valign="bottom"><span class="PLTe">Observa��es</span></td>
	<td class="MB" align="right" valign="bottom"><span class="PLTd">Qtde</span></td>
	<td class="MB" align="right" valign="bottom"><span class="PLTd">Pre�o</span></td>
	<td class="MB" align="right" valign="bottom"><span class="PLTd">VL Lista</span></td>
	<td class="MB" align="right" valign="bottom"><span class="PLTd">Desc</span></td>
	<td class="MB" align="right" valign="bottom"><span class="PLTd">VL Venda</span></td>
	<td class="MB" align="right" valign="bottom"><span class="PLTd">VL Total</span></td>
	</tr>

<% m_TotalDestePedido=0
   m_TotalDestePedidoComRA=0
   n = Lbound(v_item)-1
   for i=1 to MAX_ITENS 
	 s_readonly = "readonly tabindex=-1"
	 s_readonly_valor = "readonly tabindex=-1"
	 n = n+1
	 if n <= Ubound(v_item) then
		with v_item(n)
			s_fabricante=.fabricante
			s_produto=.produto
			s_descricao=.descricao
			s_descricao_html=produto_formata_descricao_em_html(.descricao_html)
			s_obs=.obs
			s_qtde=.qtde
			s_preco_lista=formata_moeda(.preco_lista)
			if .desc_dado=0 then s_desc_dado="" else s_desc_dado=formata_perc_desc(.desc_dado)
			s_vl_unitario=formata_moeda(.preco_venda)
			s_preco_NF=formata_moeda(.preco_NF)
			m_TotalItem=.qtde * .preco_venda
			m_TotalItemComRA=.qtde * .preco_NF
			s_vl_TotalItem=formata_moeda(m_TotalItem)
			m_TotalDestePedido=m_TotalDestePedido + m_TotalItem
			m_TotalDestePedidoComRA=m_TotalDestePedidoComRA + m_TotalItemComRA
			if operacao_permitida(OP_CEN_EDITA_ITEM_DO_ORCAMENTO, s_lista_operacoes_permitidas) And (r_orcamento.st_orcamento<>ST_ORCAMENTO_CANCELADO) then s_readonly_valor = ""
			s_readonly = ""
			end with
	 else
		s_fabricante=""
		s_produto=""
		s_descricao=""
		s_descricao_html=""
		s_obs=""
		s_qtde=""
		s_preco_lista=""
		s_desc_dado=""
		s_vl_unitario=""
		s_preco_NF=""
		s_vl_TotalItem=""
		end if
%>
	<tr>
	<td class="MDBE" align="left"><input name="c_fabricante" id="c_fabricante" class="PLLe" style="width:26px;"
		value='<%=s_fabricante%>' readonly tabindex=-1></td>
	<td class="MDB" align="left"><input name="c_produto" id="c_produto" class="PLLe" style="width:55px;"
		value='<%=s_produto%>' readonly tabindex=-1></td>
	<td class="MDB" align="left" style="width:277px;">
		<span class="PLLe"><%=s_descricao_html%></span>
		<input type="hidden" name="c_descricao" id="c_descricao" value='<%=s_descricao%>'>
	</td>
	<td class="MDB" align="left"><input name="c_obs" id="c_obs" maxlength="10" class="PLLe" style="width:80px;"
		onkeypress="if (digitou_enter(true)) {if ((<%=Cstr(i)%>==fORC.c_obs.length)||(trim(fORC.c_produto[<%=Cstr(i)%>].value)=='')) fORC.c_obs1.focus(); else fORC.c_obs[<%=Cstr(i)%>].focus();} filtra_nome_identificador();" onblur="this.value=trim(this.value);"
		value='<%=s_obs%>' <%=s_readonly%>></td>
	<td class="MDB" align="right"><input name="c_qtde" id="c_qtde" class="PLLd" style="width:27px;"
		value='<%=s_qtde%>' readonly tabindex=-1></td>
	<td class="MDB" align="right"><input name="c_vl_NF" id="c_vl_NF" class="PLLd" style="width:62px;"
		onkeypress="if (digitou_enter(true)) fORC.c_vl_unitario[<%=Cstr(i-1)%>].focus(); filtra_moeda_positivo();" onblur="this.value=formata_moeda(this.value); recalcula_RA();"
		value='<%=s_preco_NF%>' <%=s_readonly_valor%>></td>
	<td class="MDB" align="right"><input name="c_preco_lista" id="c_preco_lista" class="PLLd" style="width:62px;"
		value='<%=s_preco_lista%>' readonly tabindex=-1></td>
	<td class="MDB" align="right"><input name="c_desc" id="c_desc" class="PLLd" style="width:36px;"
		value='<%=s_desc_dado%>' readonly tabindex=-1></td>
	<td class="MDB" align="right"><input name="c_vl_unitario" id="c_vl_unitario" class="PLLd" style="width:62px;"
		onkeypress="if (digitou_enter(true)) {if ((<%=Cstr(i)%>==fORC.c_vl_unitario.length)||(trim(fORC.c_produto[<%=Cstr(i)%>].value)=='')) fORC.c_obs1.focus(); else fORC.c_vl_NF[<%=Cstr(i)%>].focus();} filtra_moeda_positivo();" onblur="this.value=formata_moeda(this.value); recalcula_total(<%=Cstr(i)%>); recalcula_RA();"
		value='<%=s_vl_unitario%>' <%=s_readonly_valor%>></td>
	<td class="MDB" align="right"><input name="c_vl_total" id="c_vl_total" class="PLLd" style="width:70px;" 
		value='<%=s_vl_TotalItem%>' readonly tabindex=-1></td>
	</tr>
<% next %>

	<tr>
	<td colspan="4" align="left">
		<table cellspacing="0" cellpadding="0" width='100%' style="margin-top:4px;">
			<tr>
			<td width="60%" align="left">&nbsp;</td>
			<td align="right">
				<table cellspacing="0" cellpadding="0" style="margin-right:2px;">
					<tr>
						<td class="MTBE" align="left"><span class="PLTe">&nbsp;RA</span></td>
						<td class="MTBD" align="right"><input name="c_total_RA" id="c_total_RA" class="PLLd" style="width:70px;color:<%if m_total_RA >=0 then Response.Write " green" else Response.Write " red"%>;" 
							value='<%=formata_moeda(m_total_RA)%>' readonly tabindex=-1></td>
					</tr>
				</table>
			</td>
			<td align="right">
				<table cellspacing="0" cellpadding="0" style="margin-right:2px;">
					<tr>
						<td class="MTBE" align="left" nowrap><span class="PLTe">&nbsp;COM(%)</span></td>
						<td class="MTBD" align="left"><input name="c_perc_RT" id="c_perc_RT" class="PLLd" style="width:30px;color:blue;" 
							value='<%=formata_perc_RT(r_orcamento.perc_RT)%>' maxlength="5" 
							onkeypress="if (digitou_enter(true)) fORC.c_obs1.focus(); filtra_percentual();"
							onblur="this.value=formata_perc_RT(this.value); if ((converte_numero(this.value)>100)||(converte_numero(this.value)<0)) {alert('Percentual inv�lido!!');this.focus();}"
							<%if r_orcamento.st_orcamento=ST_ORCAMENTO_CANCELADO then Response.Write " readonly tabindex=-1"%>
							></td>
					</tr>
				</table>
			</td>
			</tr>
		</table>
	</td>
	<td class="MD" align="left">&nbsp;</td>
	<td class="MDB" align="right">
		<input name="c_total_NF" id="c_total_NF" class="PLLd" style="width:70px;color:blue;" 
				value='<%=formata_moeda(m_TotalDestePedidoComRA)%>' readonly tabindex=-1>
	</td>
	<td colspan="3" class="MD" align="left">&nbsp;</td>
	<td class="MDB" align="right"><input name="c_total_geral" id="c_total_geral" class="PLLd" style="width:70px;color:blue;" 
		value='<%=formata_moeda(m_TotalDestePedido)%>' readonly tabindex=-1></td>
	</tr>
</table>

<input type="hidden" name="c_total_RA_original" id="c_total_RA_original" value='<%=formata_moeda(m_total_RA)%>'>


<% if r_orcamento.tipo_parcelamento = 0 then %>
<input type="hidden" name="versao_forma_pagamento" id="versao_forma_pagamento" value='1'>
<!--  TRATA VERS�O ANTIGA DA FORMA DE PAGAMENTO   -->
<br>
<table class="Q" cellspacing="0" style="width:649px;">
	<tr>
		<td class="MB" colspan="5" align="left"><p class="Rf">Observa��es I</p>
			<textarea name="c_obs1" id="c_obs1" class="PLLe" rows="<%=Cstr(MAX_LINHAS_OBS1)%>" 
				style="width:642px;margin-left:2pt;" onkeypress="limita_tamanho(this,MAX_TAM_OBS1);" onblur="this.value=trim(this.value);"
				><%=r_orcamento.obs_1%></textarea>
		</td>
	</tr>
	<tr>
		<td class="MB" colspan="5" align="left"><p class="Rf">Observa��es II</p>
			<input name="c_obs2" id="c_obs2" class="PLLe" maxlength="10" style="width:85px;margin-left:2pt;" onkeypress="if (digitou_enter(true)) fORC.c_qtde_parcelas.focus(); filtra_nome_identificador();" onblur="this.value=trim(this.value);"
				value='<%=r_orcamento.obs_2%>'>
		</td>
	</tr>
	<tr>
		<td class="MDB" nowrap width="10%" align="left"><p class="Rf">Parcelas</p>
			<table cellspacing="0" cellpadding="0" width="100%"><tr>
				<td align="left"><input name="c_qtde_parcelas" id="c_qtde_parcelas" class="PLLc" maxlength="2" style="width:60px;" onkeypress="if (digitou_enter(true)) fORC.c_forma_pagto.focus(); filtra_numerico();"
						value='<%if (r_orcamento.qtde_parcelas<>0) Or (r_orcamento.forma_pagto<>"") then Response.write Cstr(r_orcamento.qtde_parcelas)%>'></td>
			</tr></table>
		</td>
		<td class="MDB" nowrap align="left" valign="top"><p class="Rf">Entrega Imediata</p>
			<input type="radio" id="rb_etg_imediata" name="rb_etg_imediata" 
				value="<%=COD_ETG_IMEDIATA_NAO%>" <%if Cstr(r_orcamento.st_etg_imediata)=Cstr(COD_ETG_IMEDIATA_NAO) then Response.Write " checked"%>><span class="C" style="cursor:default" onclick="fORC.rb_etg_imediata[0].click();">N�o</span>
			<input type="radio" id="rb_etg_imediata" name="rb_etg_imediata" 
				value="<%=COD_ETG_IMEDIATA_SIM%>" <%if Cstr(r_orcamento.st_etg_imediata)=Cstr(COD_ETG_IMEDIATA_SIM) then Response.Write " checked"%>><span class="C" style="cursor:default" onclick="fORC.rb_etg_imediata[1].click();">Sim</span>
		</td>
		<td class="MDB" nowrap align="left" valign="top"><p class="Rf">Bem de Uso/Consumo</p>
			<input type="radio" id="rb_bem_uso_consumo" name="rb_bem_uso_consumo" 
				value="<%=COD_ST_BEM_USO_CONSUMO_NAO%>" <%if Cstr(r_orcamento.StBemUsoConsumo)=Cstr(COD_ST_BEM_USO_CONSUMO_NAO) then Response.Write " checked"%>><span class="C" style="cursor:default" onclick="fORC.rb_bem_uso_consumo[0].click();">N�o</span>
			<input type="radio" id="rb_bem_uso_consumo" name="rb_bem_uso_consumo" 
				value="<%=COD_ST_BEM_USO_CONSUMO_SIM%>" <%if Cstr(r_orcamento.StBemUsoConsumo)=Cstr(COD_ST_BEM_USO_CONSUMO_SIM) then Response.Write " checked"%>><span class="C" style="cursor:default" onclick="fORC.rb_bem_uso_consumo[1].click();">Sim</span>
		</td>
		<td class="MDB" align="left" nowrap><p class="Rf">Instalador Instala</p>
			<% if blnInstaladorInstalaBloqueado then strDisabled=" disabled" else strDisabled=""%>
			<input type="radio" id="rb_instalador_instala" name="rb_instalador_instala" 
				<%=strDisabled%>
				value="<%=COD_INSTALADOR_INSTALA_NAO%>" <%if Cstr(r_orcamento.InstaladorInstalaStatus)=Cstr(COD_INSTALADOR_INSTALA_NAO) then Response.Write " checked"%>><span class="C" style="cursor:default" onclick="fORC.rb_instalador_instala[0].click();">N�o</span>
			<input type="radio" id="rb_instalador_instala" name="rb_instalador_instala" 
				<%=strDisabled%>
				value="<%=COD_INSTALADOR_INSTALA_SIM%>" <%if Cstr(r_orcamento.InstaladorInstalaStatus)=Cstr(COD_INSTALADOR_INSTALA_SIM) then Response.Write " checked"%>><span class="C" style="cursor:default" onclick="fORC.rb_instalador_instala[1].click();">Sim</span>
		</td>
		<td class="MB" nowrap align="left" valign="top"><p class="Rf">Garantia Indicador</p>
			<% if blnGarantiaIndicadorBloqueado then strDisabled=" disabled" else strDisabled=""%>
			<input type="radio" id="rb_garantia_indicador" name="rb_garantia_indicador" 
				<%=strDisabled%>
				value="<%=COD_GARANTIA_INDICADOR_STATUS__NAO%>" <%if Cstr(r_orcamento.GarantiaIndicadorStatus)=Cstr(COD_GARANTIA_INDICADOR_STATUS__NAO) then Response.Write " checked"%>><span class="C" style="cursor:default" onclick="fORC.rb_garantia_indicador[0].click();">N�o</span>
			<input type="radio" id="rb_garantia_indicador" name="rb_garantia_indicador" 
				<%=strDisabled%>
				value="<%=COD_GARANTIA_INDICADOR_STATUS__SIM%>" <%if Cstr(r_orcamento.GarantiaIndicadorStatus)=Cstr(COD_GARANTIA_INDICADOR_STATUS__SIM) then Response.Write " checked"%>><span class="C" style="cursor:default" onclick="fORC.rb_garantia_indicador[1].click();">Sim</span>
		</td>
	</tr>
	<tr>
		<td colspan="5" align="left"><p class="Rf">Forma de Pagamento</p>
			<textarea name="c_forma_pagto" id="c_forma_pagto" class="PLLe" rows="<%=Cstr(MAX_LINHAS_FORMA_PAGTO)%>"
				style="width:642px;margin-left:2pt;" onkeypress="limita_tamanho(this,MAX_TAM_FORMA_PAGTO);" onblur="this.value=trim(this.value);"
				><%=r_orcamento.forma_pagto%></textarea>
		</td>
	</tr>
</table>
<% else %>
<!--  TRATA NOVA VERS�O DA FORMA DE PAGAMENTO   -->
<input type="hidden" name='versao_forma_pagamento' id="versao_forma_pagamento" value='2'>
<br>
<table class="Q" cellspacing="0" style="width:649px;">
	<tr>
		<td class="MB" colspan="5" align="left"><p class="Rf">Observa��es I</p>
			<textarea name="c_obs1" id="c_obs1" class="PLLe" rows="<%=Cstr(MAX_LINHAS_OBS1)%>" 
				style="width:642px;margin-left:2pt;" onkeypress="limita_tamanho(this,MAX_TAM_OBS1);" onblur="this.value=trim(this.value);"
				><%=r_orcamento.obs_1%></textarea>
		</td>
	</tr>
	<tr>
		<td class="MB" align="left" colspan="5">
			<p class="Rf">Previs�o de Entrega</p>
			<input name="c_data_previsao_entrega" id="c_data_previsao_entrega" class="PLLe" maxlength="10" style="width:90px;margin-left:2pt"
				value="<%=formata_data(r_orcamento.PrevisaoEntregaData)%>" />
		</td>
	</tr>
	<tr>
		<td class="MD" align="left" nowrap><p class="Rf">N� Nota Fiscal</p>
			<input name="c_obs2" id="c_obs2" class="PLLe" maxlength="10" style="width:85px;margin-left:2pt;" readonly tabindex="-1" onkeypress="if (digitou_enter(true)) fORC.c_qtde_parcelas.focus(); filtra_nome_identificador();" onblur="this.value=trim(this.value);"
				value='<%=r_orcamento.obs_2%>'>
		</td>
		<td class="MD" nowrap align="left" valign="top"><p class="Rf">Entrega Imediata</p>
			<input type="radio" id="rb_etg_imediata" name="rb_etg_imediata" 
				value="<%=COD_ETG_IMEDIATA_NAO%>" <%if Cstr(r_orcamento.st_etg_imediata)=Cstr(COD_ETG_IMEDIATA_NAO) then Response.Write " checked"%>><span class="C" style="cursor:default" onclick="fORC.rb_etg_imediata[0].click();">N�o</span>
			<input type="radio" id="rb_etg_imediata" name="rb_etg_imediata" 
				value="<%=COD_ETG_IMEDIATA_SIM%>" <%if Cstr(r_orcamento.st_etg_imediata)=Cstr(COD_ETG_IMEDIATA_SIM) then Response.Write " checked"%>><span class="C" style="cursor:default" onclick="fORC.rb_etg_imediata[1].click();">Sim</span>
		</td>
		<td class="MD" nowrap align="left" valign="top"><p class="Rf">Bem de Uso/Consumo</p>
			<input type="radio" id="rb_bem_uso_consumo" name="rb_bem_uso_consumo" 
				value="<%=COD_ST_BEM_USO_CONSUMO_NAO%>" <%if Cstr(r_orcamento.StBemUsoConsumo)=Cstr(COD_ST_BEM_USO_CONSUMO_NAO) then Response.Write " checked"%>><span class="C" style="cursor:default" onclick="fORC.rb_bem_uso_consumo[0].click();">N�o</span>
			<input type="radio" id="rb_bem_uso_consumo" name="rb_bem_uso_consumo" 
				value="<%=COD_ST_BEM_USO_CONSUMO_SIM%>" <%if Cstr(r_orcamento.StBemUsoConsumo)=Cstr(COD_ST_BEM_USO_CONSUMO_SIM) then Response.Write " checked"%>><span class="C" style="cursor:default" onclick="fORC.rb_bem_uso_consumo[1].click();">Sim</span>
		</td>
		<td class="MD" align="left" nowrap><p class="Rf">Instalador Instala</p>
			<% if blnInstaladorInstalaBloqueado then strDisabled=" disabled" else strDisabled=""%>
			<input type="radio" id="rb_instalador_instala" name="rb_instalador_instala" 
				<%=strDisabled%>
				value="<%=COD_INSTALADOR_INSTALA_NAO%>" <%if Cstr(r_orcamento.InstaladorInstalaStatus)=Cstr(COD_INSTALADOR_INSTALA_NAO) then Response.Write " checked"%>><span class="C" style="cursor:default" onclick="fORC.rb_instalador_instala[0].click();">N�o</span>
			<input type="radio" id="rb_instalador_instala" name="rb_instalador_instala" 
				<%=strDisabled%>
				value="<%=COD_INSTALADOR_INSTALA_SIM%>" <%if Cstr(r_orcamento.InstaladorInstalaStatus)=Cstr(COD_INSTALADOR_INSTALA_SIM) then Response.Write " checked"%>><span class="C" style="cursor:default" onclick="fORC.rb_instalador_instala[1].click();">Sim</span>
		</td>
		<td nowrap align="left" valign="top"><p class="Rf">Garantia Indicador</p>
			<% if blnGarantiaIndicadorBloqueado then strDisabled=" disabled" else strDisabled=""%>
			<input type="radio" id="rb_garantia_indicador" name="rb_garantia_indicador" 
				<%=strDisabled%>
				value="<%=COD_GARANTIA_INDICADOR_STATUS__NAO%>" <%if Cstr(r_orcamento.GarantiaIndicadorStatus)=Cstr(COD_GARANTIA_INDICADOR_STATUS__NAO) then Response.Write " checked"%>><span class="C" style="cursor:default" onclick="fORC.rb_garantia_indicador[0].click();">N�o</span>
			<input type="radio" id="rb_garantia_indicador" name="rb_garantia_indicador" 
				<%=strDisabled%>
				value="<%=COD_GARANTIA_INDICADOR_STATUS__SIM%>" <%if Cstr(r_orcamento.GarantiaIndicadorStatus)=Cstr(COD_GARANTIA_INDICADOR_STATUS__SIM) then Response.Write " checked"%>><span class="C" style="cursor:default" onclick="fORC.rb_garantia_indicador[1].click();">Sim</span>
		</td>
	</tr>
</table>
<br>
<table class="Q" style="width:649px;" cellspacing="0">
  <tr>
	<td align="left">
	  <p class="Rf">Forma de Pagamento</p>
	</td>
  </tr>  
  <tr>
	<td align="left">
	  <table width="100%" cellspacing="0" cellpadding="4" border="0">
		<!--  � VISTA  -->
		<tr>
		  <td align="left">
			<table cellspacing="0" cellpadding="1" border="0">
			  <tr>
				<td align="left">
				  <% intIdx = 0 %>
				  <input type="radio" id="rb_forma_pagto" name="rb_forma_pagto" 
						value="<%=COD_FORMA_PAGTO_A_VISTA%>"
						<% if Cstr(r_orcamento.tipo_parcelamento) = COD_FORMA_PAGTO_A_VISTA then Response.Write " checked"%>
						onclick="recalculaCustoFinanceiroPrecoLista();"
				  ><span class="C" style="cursor:default" onclick="fORC.rb_forma_pagto[<%=Cstr(intIdx)%>].click();">� Vista</span>
				</td>
				<td align="left">&nbsp;</td>
				<td align="left">
				  <select id="op_av_forma_pagto" name="op_av_forma_pagto" onclick="fORC.rb_forma_pagto[<%=Cstr(intIdx)%>].click();">
					<%	if operacao_permitida(OP_CEN_EDITA_FORMA_PAGTO_SEM_APLICAR_RESTRICOES, s_lista_operacoes_permitidas) then
							if Cstr(r_orcamento.tipo_parcelamento) = COD_FORMA_PAGTO_A_VISTA then 
								Response.Write forma_pagto_av_monta_itens_select_incluindo_default(r_orcamento.av_forma_pagto)
							else
								Response.Write forma_pagto_av_monta_itens_select(Null)
								end if
						else
							if Cstr(r_orcamento.tipo_parcelamento) = COD_FORMA_PAGTO_A_VISTA then 
								Response.Write forma_pagto_liberada_av_monta_itens_select_incluindo_default(r_orcamento.av_forma_pagto, r_orcamento.orcamentista, r_cliente.tipo)
							else
								Response.Write forma_pagto_liberada_av_monta_itens_select(Null, r_orcamento.orcamentista, r_cliente.tipo)
								end if
							end if
					%>
				  </select>
				</td>
			  </tr>
			</table>
		  </td>
		</tr>
		<!--  PARCELA �NICA  -->
		<tr>
		  <td class="MC" align="left">
			<table cellspacing="0" cellpadding="1" border="0">
			  <tr>
				<td colspan="3" align="left">
				  <% intIdx = intIdx+1 %>
				  <input type="radio" id="rb_forma_pagto" name="rb_forma_pagto" 
						value="<%=COD_FORMA_PAGTO_PARCELA_UNICA%>"
						<% if Cstr(r_orcamento.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELA_UNICA then Response.Write " checked"%>
						onclick="pu_atualiza_valor();recalculaCustoFinanceiroPrecoLista();"
						><span class="C" style="cursor:default" onclick="fORC.rb_forma_pagto[<%=Cstr(intIdx)%>].click();">Parcela �nica</span>
				</td>
				<td align="left">&nbsp;</td>
				<td align="left">
				  <select id="op_pu_forma_pagto" name="op_pu_forma_pagto" onclick="fORC.rb_forma_pagto[<%=Cstr(intIdx)%>].click();">
					<%	if operacao_permitida(OP_CEN_EDITA_FORMA_PAGTO_SEM_APLICAR_RESTRICOES, s_lista_operacoes_permitidas) then
							if Cstr(r_orcamento.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELA_UNICA then
								Response.Write forma_pagto_da_parcela_unica_monta_itens_select_incluindo_default(r_orcamento.pu_forma_pagto)
							else
								Response.Write forma_pagto_da_parcela_unica_monta_itens_select(Null)
								end if
						else
							if Cstr(r_orcamento.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELA_UNICA then
								Response.Write forma_pagto_liberada_da_parcela_unica_monta_itens_select_incluindo_default(r_orcamento.pu_forma_pagto, r_orcamento.orcamentista, r_cliente.tipo)
							else
								Response.Write forma_pagto_liberada_da_parcela_unica_monta_itens_select(Null, r_orcamento.orcamentista, r_cliente.tipo)
								end if
							end if
					%>
				  </select>
				  <span style="width:10px;">&nbsp;</span>
				  <span class="C" style="margin-right:0pt;"><%=SIMBOLO_MONETARIO%></span
				  ><input name="c_pu_valor" id="c_pu_valor" class="Cd" maxlength="18" style="width:90px;" onclick="fORC.rb_forma_pagto[<%=Cstr(intIdx)%>].click();" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fORC.c_pu_vencto_apos.focus(); filtra_moeda_positivo();" onblur="this.value=formata_moeda(this.value);" 
					<% if Cstr(r_orcamento.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELA_UNICA then %>
						value="<%=formata_moeda(r_orcamento.pu_valor)%>"
					<% else %>
						value=""
					<% end if %>
				  ><span style="width:10px;">&nbsp;</span
				  ><span class="C">vencendo ap�s</span
				  ><input name="c_pu_vencto_apos" id="c_pu_vencto_apos" class="Cc" maxlength="2" style="width:30px;" onclick="fORC.rb_forma_pagto[<%=Cstr(intIdx)%>].click();" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fORC.c_forma_pagto.focus(); filtra_numerico();" 
					<% if Cstr(r_orcamento.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELA_UNICA then %>
						value="<%=Cstr(r_orcamento.pu_vencto_apos)%>"
					<% else %>
						value=""
					<% end if %>
				  ><span class="C">dias</span>
				</td>
			  </tr>
			</table>
		  </td>
		</tr>
		<!--  PARCELADO NO CART�O (INTERNET)  -->
		<% if operacao_permitida(OP_CEN_EDITA_FORMA_PAGTO_SEM_APLICAR_RESTRICOES, s_lista_operacoes_permitidas) Or _
				(Cstr(r_orcamento.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_CARTAO) Or _
				(Not is_restricao_ativa_forma_pagto(r_orcamento.orcamentista, ID_FORMA_PAGTO_CARTAO, r_cliente.tipo)) then %>
		<tr>
		<% else %>
		<tr style="display:none;">
		<% end if %>
		  <td class="MC" align="left">
			<table cellspacing="0" cellpadding="1" border="0">
			  <tr>
				<td align="left">
				  <% intIdx = intIdx+1 %>
				  <input type="radio" id="rb_forma_pagto" name="rb_forma_pagto" 
						value="<%=COD_FORMA_PAGTO_PARCELADO_CARTAO%>"
						<% if Cstr(r_orcamento.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_CARTAO then Response.Write " checked"%>
						onclick="recalculaCustoFinanceiroPrecoLista();"
				   ><span class="C" style="cursor:default" onclick="fORC.rb_forma_pagto[<%=Cstr(intIdx)%>].click();">Parcelado no Cart�o (internet)</span>
				</td>
				<td align="left">&nbsp;</td>
				<td align="left">
				  <input name="c_pc_qtde" id="c_pc_qtde" class="Cc" maxlength="2" style="width:30px;" onclick="fORC.rb_forma_pagto[<%=Cstr(intIdx)%>].click();" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fORC.c_pc_valor.focus(); filtra_numerico();" onblur="pc_calcula_valor_parcela();recalculaCustoFinanceiroPrecoLista();" 
					<% if Cstr(r_orcamento.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_CARTAO then %>
						value="<%=Cstr(r_orcamento.pc_qtde_parcelas)%>"
					<% else %>
						value=""
					<% end if %>
				  >
				</td>
				<td align="left"><span class="C" style="margin-right:0pt;">&nbsp;X&nbsp;&nbsp;&nbsp;<%=SIMBOLO_MONETARIO%></span></td>
				<td align="left">
				  <input name="c_pc_valor" id="c_pc_valor" class="Cd" maxlength="18" style="width:90px;" onclick="fORC.rb_forma_pagto[<%=Cstr(intIdx)%>].click();" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fORC.c_forma_pagto.focus(); filtra_moeda_positivo();" onblur="this.value=formata_moeda(this.value);" 
					<% if Cstr(r_orcamento.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_CARTAO then %>
						value="<%=formata_moeda(r_orcamento.pc_valor_parcela)%>"
					<% else %>
						value=""
					<% end if %>
				  >
				</td>
			  </tr>
			</table>
		  </td>
		</tr>
		<!--  PARCELADO NO CART�O (MAQUINETA)  -->
		<% if operacao_permitida(OP_CEN_EDITA_FORMA_PAGTO_SEM_APLICAR_RESTRICOES, s_lista_operacoes_permitidas) Or _
				(Cstr(r_orcamento.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_CARTAO_MAQUINETA) Or _
				(Not is_restricao_ativa_forma_pagto(r_orcamento.orcamentista, ID_FORMA_PAGTO_CARTAO_MAQUINETA, r_cliente.tipo)) then %>
		<tr>
		<% else %>
		<tr style="display:none;">
		<% end if %>
		  <td class="MC" align="left">
			<table cellspacing="0" cellpadding="1" border="0">
			  <tr>
				<td align="left">
				  <% intIdx = intIdx+1 %>
				  <input type="radio" id="rb_forma_pagto" name="rb_forma_pagto" 
						value="<%=COD_FORMA_PAGTO_PARCELADO_CARTAO_MAQUINETA%>"
						<% if Cstr(r_orcamento.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_CARTAO_MAQUINETA then Response.Write " checked"%>
						onclick="recalculaCustoFinanceiroPrecoLista();"
				   ><span class="C" style="cursor:default" onclick="fORC.rb_forma_pagto[<%=Cstr(intIdx)%>].click();">Parcelado no Cart�o (maquineta)</span>
				</td>
				<td align="left">&nbsp;</td>
				<td align="left">
				  <input name="c_pc_maquineta_qtde" id="c_pc_maquineta_qtde" class="Cc" maxlength="2" style="width:30px;" onclick="fORC.rb_forma_pagto[<%=Cstr(intIdx)%>].click();" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fORC.c_pc_maquineta_valor.focus(); filtra_numerico();" onblur="pc_maquineta_calcula_valor_parcela();recalculaCustoFinanceiroPrecoLista();" 
					<% if Cstr(r_orcamento.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_CARTAO_MAQUINETA then %>
						value="<%=Cstr(r_orcamento.pc_maquineta_qtde_parcelas)%>"
					<% else %>
						value=""
					<% end if %>
				  >
				</td>
				<td align="left"><span class="C" style="margin-right:0pt;">&nbsp;X&nbsp;&nbsp;&nbsp;<%=SIMBOLO_MONETARIO%></span></td>
				<td align="left">
				  <input name="c_pc_maquineta_valor" id="c_pc_maquineta_valor" class="Cd" maxlength="18" style="width:90px;" onclick="fORC.rb_forma_pagto[<%=Cstr(intIdx)%>].click();" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fORC.c_forma_pagto.focus(); filtra_moeda_positivo();" onblur="this.value=formata_moeda(this.value);" 
					<% if Cstr(r_orcamento.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_CARTAO_MAQUINETA then %>
						value="<%=formata_moeda(r_orcamento.pc_maquineta_valor_parcela)%>"
					<% else %>
						value=""
					<% end if %>
				  >
				</td>
			  </tr>
			</table>
		  </td>
		</tr>
		<!--  PARCELADO COM ENTRADA  -->
		<tr>
		  <td class="MC" align="left">
			<table cellspacing="0" cellpadding="1" border="0">
			  <tr>
				<td colspan="3" align="left">
				  <% intIdx = intIdx+1 %>
				  <input type="radio" id="rb_forma_pagto" name="rb_forma_pagto" 
						value="<%=COD_FORMA_PAGTO_PARCELADO_COM_ENTRADA%>"
						<% if Cstr(r_orcamento.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_COM_ENTRADA then Response.Write " checked"%>
						onclick="pce_preenche_sugestao_intervalo();recalculaCustoFinanceiroPrecoLista();"
						><span class="C" style="cursor:default" onclick="fORC.rb_forma_pagto[<%=Cstr(intIdx)%>].click();">Parcelado com Entrada</span>
				</td>
			  </tr>
			  <tr>
				<td style="width:60px;" align="left">&nbsp;</td>
				<td align="right"><span class="C">Entrada&nbsp;</span></td>
				<td align="left">
				  <select id="op_pce_entrada_forma_pagto" name="op_pce_entrada_forma_pagto" onclick="fORC.rb_forma_pagto[<%=Cstr(intIdx)%>].click();">
					<%	if operacao_permitida(OP_CEN_EDITA_FORMA_PAGTO_SEM_APLICAR_RESTRICOES, s_lista_operacoes_permitidas) then
							if Cstr(r_orcamento.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_COM_ENTRADA then
								Response.Write forma_pagto_da_entrada_monta_itens_select_incluindo_default(r_orcamento.pce_forma_pagto_entrada)
							else
								Response.Write forma_pagto_da_entrada_monta_itens_select(Null)
								end if
						else
							if Cstr(r_orcamento.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_COM_ENTRADA then
								Response.Write forma_pagto_liberada_da_entrada_monta_itens_select_incluindo_default(r_orcamento.pce_forma_pagto_entrada, r_orcamento.orcamentista, r_cliente.tipo)
							else
								Response.Write forma_pagto_liberada_da_entrada_monta_itens_select(Null, r_orcamento.orcamentista, r_cliente.tipo)
								end if
							end if
					%>
				  </select>
				  <span style="width:10px;">&nbsp;</span>
				  <span class="C" style="margin-right:0pt;"><%=SIMBOLO_MONETARIO%></span
				  ><input name="c_pce_entrada_valor" id="c_pce_entrada_valor" class="Cd" maxlength="18" style="width:90px;" onclick="fORC.rb_forma_pagto[<%=Cstr(intIdx)%>].click();" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fORC.op_pce_prestacao_forma_pagto.focus(); filtra_moeda_positivo();" onblur="this.value=formata_moeda(this.value);pce_calcula_valor_parcela();" 
					<% if Cstr(r_orcamento.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_COM_ENTRADA then %>
						value="<%=formata_moeda(r_orcamento.pce_entrada_valor)%>"
					<% else %>
						value=""
					<% end if %>
				  >
				</td>
			  </tr>
			  <tr>
				<td style="width:60px;" align="left">&nbsp;</td>
				<td align="right"><span class="C">Presta��es&nbsp;</span></td>
				<td align="left">
				  <select id="op_pce_prestacao_forma_pagto" name="op_pce_prestacao_forma_pagto" onclick="fORC.rb_forma_pagto[<%=Cstr(intIdx)%>].click();">
					<%	if operacao_permitida(OP_CEN_EDITA_FORMA_PAGTO_SEM_APLICAR_RESTRICOES, s_lista_operacoes_permitidas) then
							if Cstr(r_orcamento.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_COM_ENTRADA then
								Response.Write forma_pagto_da_prestacao_monta_itens_select_incluindo_default(r_orcamento.pce_forma_pagto_prestacao)
							else
								Response.Write forma_pagto_da_prestacao_monta_itens_select(Null)
								end if
						else
							if Cstr(r_orcamento.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_COM_ENTRADA then
								Response.Write forma_pagto_liberada_da_prestacao_monta_itens_select_incluindo_default(r_orcamento.pce_forma_pagto_prestacao, r_orcamento.orcamentista, r_cliente.tipo)
							else
								Response.Write forma_pagto_liberada_da_prestacao_monta_itens_select(Null, r_orcamento.orcamentista, r_cliente.tipo)
								end if
							end if
					%>
				  </select>
				  <span style="width:10px;">&nbsp;</span>
				  <input name="c_pce_prestacao_qtde" id="c_pce_prestacao_qtde" class="Cc" maxlength="2" style="width:30px;" onclick="fORC.rb_forma_pagto[<%=Cstr(intIdx)%>].click();" onblur="pce_calcula_valor_parcela();recalculaCustoFinanceiroPrecoLista();" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fORC.c_pce_prestacao_valor.focus(); filtra_numerico();" 
					<% if Cstr(r_orcamento.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_COM_ENTRADA then %>
						value="<%=Cstr(r_orcamento.pce_prestacao_qtde)%>"
					<% else %>
						value=""
					<% end if %>
				  ><span class="C" style="margin-right:0pt;">&nbsp;X&nbsp;&nbsp;&nbsp;<%=SIMBOLO_MONETARIO%></span
				  ><input name="c_pce_prestacao_valor" id="c_pce_prestacao_valor" class="Cd" maxlength="18" style="width:90px;" onclick="fORC.rb_forma_pagto[<%=Cstr(intIdx)%>].click();" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fORC.c_pce_prestacao_periodo.focus(); filtra_moeda_positivo();" onblur="this.value=formata_moeda(this.value);" 
					<% if Cstr(r_orcamento.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_COM_ENTRADA then %>
						value="<%=formata_moeda(r_orcamento.pce_prestacao_valor)%>"
					<% else %>
						value=""
					<% end if %>
				  >
				</td>
			  </tr>
			  <tr>
				<td style="width:60px;" align="left">&nbsp;</td>
				<td colspan="2" align="left"><span class="C">Parcelas vencendo a cada</span
				><input name="c_pce_prestacao_periodo" id="c_pce_prestacao_periodo" class="Cc" maxlength="2" style="width:30px;" onclick="fORC.rb_forma_pagto[<%=Cstr(intIdx)%>].click();" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fORC.c_forma_pagto.focus(); filtra_numerico();" 
					<% if Cstr(r_orcamento.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_COM_ENTRADA then %>
						value="<%=Cstr(r_orcamento.pce_prestacao_periodo)%>"
					<% else %>
						value=""
					<% end if %>
				><span class="C">dias</span
				><span style="width:10px;">&nbsp;</span
				><span class="notPrint"><input name="b_pce_SugereFormaPagto" id="b_pce_SugereFormaPagto" type="button" class="Button" onclick="pce_sugestao_forma_pagto();" value="sugest�o autom�tica" title="preenche o campo 'Forma de Pagamento' com uma sugest�o de texto"></span
				></td>
			  </tr>
			</table>
		  </td>
		</tr>
		<!--  PARCELADO SEM ENTRADA  -->
		<tr>
		  <td class="MC" align="left">
			<table cellspacing="0" cellpadding="1" border="0">
			  <tr>
				<td colspan="3" align="left">
				  <% intIdx = intIdx+1 %>
				  <input type="radio" id="rb_forma_pagto" name="rb_forma_pagto" 
						value="<%=COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA%>"
						<% if Cstr(r_orcamento.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA then Response.Write " checked"%>
						onclick="pse_preenche_sugestao_intervalo();recalculaCustoFinanceiroPrecoLista();"
						><span class="C" style="cursor:default" onclick="fORC.rb_forma_pagto[<%=Cstr(intIdx)%>].click();">Parcelado sem Entrada</span>
				</td>
			  </tr>
			  <tr>
				<td style="width:60px;" align="left">&nbsp;</td>
				<td align="right"><span class="C">1� Presta��o&nbsp;</span></td>
				<td align="left">
				  <select id="op_pse_prim_prest_forma_pagto" name="op_pse_prim_prest_forma_pagto" onclick="fORC.rb_forma_pagto[<%=Cstr(intIdx)%>].click();">
					<%	if operacao_permitida(OP_CEN_EDITA_FORMA_PAGTO_SEM_APLICAR_RESTRICOES, s_lista_operacoes_permitidas) then
							if Cstr(r_orcamento.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA then
								Response.Write forma_pagto_da_prestacao_monta_itens_select_incluindo_default(r_orcamento.pse_forma_pagto_prim_prest)
							else
								Response.Write forma_pagto_da_prestacao_monta_itens_select(Null)
								end if
						else
							if Cstr(r_orcamento.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA then
								Response.Write forma_pagto_liberada_da_prestacao_monta_itens_select_incluindo_default(r_orcamento.pse_forma_pagto_prim_prest, r_orcamento.orcamentista, r_cliente.tipo)
							else
								Response.Write forma_pagto_liberada_da_prestacao_monta_itens_select(Null, r_orcamento.orcamentista, r_cliente.tipo)
								end if
							end if
					%>
				  </select>
				  <span style="width:10px;">&nbsp;</span>
				  <span class="C" style="margin-right:0pt;"><%=SIMBOLO_MONETARIO%></span
				  ><input name="c_pse_prim_prest_valor" id="c_pse_prim_prest_valor" class="Cd" maxlength="18" style="width:90px;" onclick="fORC.rb_forma_pagto[<%=Cstr(intIdx)%>].click();" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fORC.c_pse_prim_prest_apos.focus(); filtra_moeda_positivo();" onblur="this.value=formata_moeda(this.value); pse_calcula_valor_parcela();" 
					<% if Cstr(r_orcamento.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA then %>
						value="<%=formata_moeda(r_orcamento.pse_prim_prest_valor)%>"
					<% else %>
						value=""
					<% end if %>
				  ><span style="width:10px;">&nbsp;</span
				  ><span class="C">vencendo ap�s</span
				  ><input name="c_pse_prim_prest_apos" id="c_pse_prim_prest_apos" class="Cc" maxlength="2" style="width:30px;" onclick="fORC.rb_forma_pagto[<%=Cstr(intIdx)%>].click();" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fORC.op_pse_demais_prest_forma_pagto.focus(); filtra_numerico();" 
					<% if Cstr(r_orcamento.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA then %>
						value="<%=Cstr(r_orcamento.pse_prim_prest_apos)%>"
					<% else %>
						value=""
					<% end if %>
				  ><span class="C">dias</span>
				</td>
			  </tr>
			  <tr>
				<td style="width:60px;" align="left">&nbsp;</td>
				<td align="right"><span class="C">Demais Presta��es&nbsp;</span></td>
				<td align="left">
				  <select id="op_pse_demais_prest_forma_pagto" name="op_pse_demais_prest_forma_pagto" onclick="fORC.rb_forma_pagto[<%=Cstr(intIdx)%>].click();">
					<%	if operacao_permitida(OP_CEN_EDITA_FORMA_PAGTO_SEM_APLICAR_RESTRICOES, s_lista_operacoes_permitidas) then
							if Cstr(r_orcamento.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA then
								Response.Write forma_pagto_da_prestacao_monta_itens_select_incluindo_default(r_orcamento.pse_forma_pagto_demais_prest)
							else
								Response.Write forma_pagto_da_prestacao_monta_itens_select(Null)
								end if
						else
							if Cstr(r_orcamento.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA then
								Response.Write forma_pagto_liberada_da_prestacao_monta_itens_select_incluindo_default(r_orcamento.pse_forma_pagto_demais_prest, r_orcamento.orcamentista, r_cliente.tipo)
							else
								Response.Write forma_pagto_liberada_da_prestacao_monta_itens_select(Null, r_orcamento.orcamentista, r_cliente.tipo)
								end if
							end if
					%>
				  </select>
				  <span style="width:10px;">&nbsp;</span>
				  <input name="c_pse_demais_prest_qtde" id="c_pse_demais_prest_qtde" class="Cc" maxlength="2" style="width:30px;" onclick="fORC.rb_forma_pagto[<%=Cstr(intIdx)%>].click();" onblur="pse_calcula_valor_parcela();recalculaCustoFinanceiroPrecoLista();" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fORC.c_pse_demais_prest_valor.focus(); filtra_numerico();" 
					<% if Cstr(r_orcamento.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA then %>
						value="<%=Cstr(r_orcamento.pse_demais_prest_qtde)%>"
					<% else %>
						value=""
					<% end if %>
				  >
				  <span class="C" style="margin-right:0pt;">&nbsp;X&nbsp;&nbsp;&nbsp;<%=SIMBOLO_MONETARIO%></span
				  ><input name="c_pse_demais_prest_valor" id="c_pse_demais_prest_valor" class="Cd" maxlength="18" style="width:90px;" onclick="fORC.rb_forma_pagto[<%=Cstr(intIdx)%>].click();" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fORC.c_pse_demais_prest_periodo.focus(); filtra_moeda_positivo();" onblur="this.value=formata_moeda(this.value); " 
					<% if Cstr(r_orcamento.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA then %>
						value="<%=formata_moeda(r_orcamento.pse_demais_prest_valor)%>"
					<% else %>
						value=""
					<% end if %>
				  >
				</td>
			  </tr>
			  <tr>
				<td style="width:60px;" align="left">&nbsp;</td>
				<td colspan="2" align="left"><span class="C">Parcelas vencendo a cada</span
				><input name="c_pse_demais_prest_periodo" id="c_pse_demais_prest_periodo" class="Cc" maxlength="2" style="width:30px;" onclick="fORC.rb_forma_pagto[<%=Cstr(intIdx)%>].click();" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fORC.c_forma_pagto.focus(); filtra_numerico();" 
					<% if Cstr(r_orcamento.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA then %>
						value="<%=Cstr(r_orcamento.pse_demais_prest_periodo)%>"
					<% else %>
						value=""
					<% end if %>
				><span class="C">dias</span
				><span style="width:10px;">&nbsp;</span
				><span class="notPrint"><input name="b_pse_SugereFormaPagto" id="b_pse_SugereFormaPagto" type="button" class="Button" onclick="pse_sugestao_forma_pagto();" value="sugest�o autom�tica" title="preenche o campo 'Forma de Pagamento' com uma sugest�o de texto"></span
				></td>
			  </tr>
			</table>
		  </td>
		</tr>
	  </table>
	</td>
  </tr>
  <tr>
	<td class="MC" align="left">
	  <p class="Rf">Descri��o da Forma de Pagamento</p>
		<textarea name="c_forma_pagto" id="c_forma_pagto" class="PLLe" rows="<%=Cstr(MAX_LINHAS_FORMA_PAGTO)%>"
			style="width:641px;margin-left:2pt;" onkeypress="limita_tamanho(this,MAX_TAM_FORMA_PAGTO);" onblur="this.value=trim(this.value);"
			><%=r_orcamento.forma_pagto%></textarea>
	</td>
  </tr>
</table>
<% end if %>


<!-- ************   SEPARADOR   ************ -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc" align="left">&nbsp;</td></tr>
</table>
<br>
<input type="hidden" name="Verifica_End_Entrega" id="Verifica_End_Entrega" value=''>
<input type="hidden" name="Verifica_num" id="Verifica_num" value=''>
<input type="hidden" name="Verifica_Cidade" id="Verifica_Cidade" value=''>
<input type="hidden" name="Verifica_UF" id="Verifica_UF" value=''>
<input type="hidden" name="Verifica_CEP" id="Verifica_CEP" value=''>
<input type="hidden" name="Verifica_Justificativa" id="Verifica_Justificativa" value=''>

<!-- ************   BOT�ES   ************ -->
<table class="notPrint" width="649" cellspacing="0">
<tr>
	<td align="left"><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()" title="volta para p�gina anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a>
	</td>
	<td align="right"><div name="dCONFIRMA" id="dCONFIRMA">
		<a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fORCConfirma(fORC)" title="confirma as altera��es">
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