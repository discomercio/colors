<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/Global.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ==============================
'	  PedidoNovoProdCompostoMask.asp
'     ==============================
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

	class cl_MAP_ITEM
		dim sku
		dim qty_ordered
		dim price
		dim name
		end class

	dim i, iv, idx_map_item, s, usuario, loja, cliente_selecionado, r_cliente, msg_erro
	usuario = Trim(Session("usuario_atual"))
	loja = Trim(Session("loja_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 
	If (loja = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

	cliente_selecionado = Trim(request("cliente_selecionado"))
	if (cliente_selecionado = "") then Response.Redirect("aviso.asp?id=" & ERR_CLIENTE_NAO_ESPECIFICADO)
	
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))
	
	dim cn, tMAP_XML, tMAP_ITEM
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	set r_cliente = New cl_CLIENTE
	if Not x_cliente_bd(cliente_selecionado, r_cliente) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
	
	if Trim(r_cliente.endereco_numero) = "" then
		Response.Redirect("aviso.asp?id=" & ERR_CAD_CLIENTE_ENDERECO_NUMERO_NAO_PREENCHIDO)
	elseif Len(Trim(r_cliente.endereco)) > CLng(MAX_TAMANHO_CAMPO_ENDERECO) then
		Response.Redirect("aviso.asp?id=" & ERR_CAD_CLIENTE_ENDERECO_EXCEDE_TAMANHO_MAXIMO)
		end if
		
	dim rb_end_entrega, EndEtg_endereco, EndEtg_endereco_numero, EndEtg_endereco_complemento
	dim EndEtg_bairro, EndEtg_cidade, EndEtg_uf, EndEtg_cep,EndEtg_obs
	rb_end_entrega = Trim(Request.Form("rb_end_entrega"))
	EndEtg_endereco = Trim(Request.Form("EndEtg_endereco"))
	EndEtg_endereco_numero = Trim(Request.Form("EndEtg_endereco_numero"))
	EndEtg_endereco_complemento = Trim(Request.Form("EndEtg_endereco_complemento"))
	EndEtg_bairro = Trim(Request.Form("EndEtg_bairro"))
	EndEtg_cidade = Trim(Request.Form("EndEtg_cidade"))
	EndEtg_uf = Trim(Request.Form("EndEtg_uf"))
	EndEtg_cep = Trim(Request.Form("EndEtg_cep"))
	EndEtg_obs = Trim(Request.Form("EndEtg_obs"))

	dim alerta
	alerta = ""

	dim s_produto, s_qtde
	dim s_nome_cliente, c_mag_cpf_cnpj_identificado
	dim operacao_origem, c_numero_magento, operationControlTicket, sessionToken, id_magento_api_pedido_xml
	operacao_origem = Trim(Request("operacao_origem"))
	c_numero_magento = ""
	operationControlTicket = ""
	sessionToken = ""
	id_magento_api_pedido_xml = ""
	c_mag_cpf_cnpj_identificado = ""
	if operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO then
		c_numero_magento = Trim(Request("c_numero_magento"))
		operationControlTicket = Trim(Request("operationControlTicket"))
		sessionToken = Trim(Request("sessionToken"))
		id_magento_api_pedido_xml = Trim(Request("id_magento_api_pedido_xml"))
		end if
	
	if operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO then
		If Not cria_recordset_otimista(tMAP_XML, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
		If Not cria_recordset_otimista(tMAP_ITEM, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
		end if

	dim v_map_item
	redim v_map_item(0)
	set v_map_item(UBound(v_map_item)) = new cl_MAP_ITEM
	v_map_item(UBound(v_map_item)).sku = ""
	idx_map_item = LBound(v_map_item)

	if alerta = "" then
		if operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO then
			s = "SELECT " & _
					"*" & _
				" FROM t_MAGENTO_API_PEDIDO_XML" & _
				" WHERE" & _
					" (id = " & id_magento_api_pedido_xml & ")"
			if tMAP_XML.State <> 0 then tMAP_XML.Close
			tMAP_XML.open s, cn
			if tMAP_XML.Eof then
				alerta = "Falha ao tentar localizar no banco de dados o registro com os dados do pedido Magento consultados via API (id = " & id_magento_api_pedido_xml & ")"
			else
				s_nome_cliente = UCase(ec_dados_formata_nome(tMAP_XML("customer_firstname"), tMAP_XML("customer_middlename"), tMAP_XML("customer_lastname"), 60))
				c_mag_cpf_cnpj_identificado = retorna_so_digitos(Trim("" & tMAP_XML("cpfCnpjIdentificado")))
				end if

			s = "SELECT " & _
					"*" & _
				" FROM t_MAGENTO_API_PEDIDO_XML_DECODE_ITEM" & _
				" WHERE" & _
					" (id_magento_api_pedido_xml = " & id_magento_api_pedido_xml & ")" & _
					" AND (product_type <> 'configurable')" & _
				" ORDER BY" & _
					" id"
			if tMAP_ITEM.State <> 0 then tMAP_ITEM.Close
			tMAP_ITEM.open s, cn
			if tMAP_ITEM.Eof then
				alerta = "Falha ao tentar localizar no banco de dados os itens do pedido Magento nº " & c_numero_magento & " (operationControlTicket = " & operationControlTicket & ")"
			else
				do while Not tMAP_ITEM.Eof
					if Trim("" & v_map_item(UBound(v_map_item)).sku) <> "" then
						redim preserve v_map_item(UBound(v_map_item)+1)
						set v_map_item(UBound(v_map_item)) = new cl_MAP_ITEM
						end if

					v_map_item(UBound(v_map_item)).sku = Trim("" & tMAP_ITEM("sku"))
					v_map_item(UBound(v_map_item)).qty_ordered = CLng(tMAP_ITEM("qty_ordered"))
					v_map_item(UBound(v_map_item)).price = converte_numero(formata_moeda(tMAP_ITEM("price")))
					v_map_item(UBound(v_map_item)).name = Trim("" & tMAP_ITEM("name"))

					tMAP_ITEM.MoveNext
					loop
				end if

			'POSICIONA O ÍNDICE NA PRIMEIRA POSIÇÃO QUE TENHA DADOS
			for iv=LBound(v_map_item) to UBound(v_map_item)
				if Trim("" & v_map_item(iv).sku) <> "" then
					idx_map_item = iv
					exit for
					end if
				next
			end if
		end if

'	CONSISTÊNCIAS P/ EMISSÃO DE NFe
	dim s_lista_sugerida_municipios
	dim v_lista_sugerida_municipios
	dim iCounterLista, iNumeracaoLista
	dim s_tabela_municipios_IBGE
	s_tabela_municipios_IBGE = ""
	if alerta = "" then
	'	DDD VÁLIDO?
		if Not ddd_ok(r_cliente.ddd_res) then
			if alerta <> "" then alerta = alerta & "<br><br>" & String(80,"=") & "<br><br>"
			alerta = alerta & "DDD do telefone residencial é inválido!!"
			end if
			
		if Not ddd_ok(r_cliente.ddd_com) then
			if alerta <> "" then alerta = alerta & "<br><br>" & String(80,"=") & "<br><br>"
			alerta = alerta & "DDD do telefone comercial é inválido!!"
			end if
			
	'	I.E. É VÁLIDA?
		if r_cliente.tipo = ID_PJ then
            if r_cliente.ie <> "" then
			    if Not isInscricaoEstadualValida(r_cliente.ie, r_cliente.uf) then
				    if alerta <> "" then alerta = alerta & "<br><br>" & String(80,"=") & "<br><br>"
				    alerta=alerta & "Corrija a IE (Inscrição Estadual) com um número válido!!" & _
						    "<br>" & "Certifique-se de que a UF informada corresponde à UF responsável pelo registro da IE."
				    end if
            end if
		end if

	'	MUNICÍPIO DE ACORDO C/ TABELA DO IBGE?
		if Not consiste_municipio_IBGE_ok(r_cliente.cidade, r_cliente.uf, s_lista_sugerida_municipios, msg_erro) then
			if alerta <> "" then alerta = alerta & "<br><br>" & String(80,"=") & "<br><br>"
			if msg_erro <> "" then
				alerta = alerta & msg_erro
			else
				alerta = alerta & "Município '" & r_cliente.cidade & "' não consta na relação de municípios do IBGE para a UF de '" & r_cliente.uf & "'!!"
				if s_lista_sugerida_municipios <> "" then
					alerta = alerta & "<br>" & _
									  "Localize o município na lista abaixo e verifique se a grafia está correta!!"
					v_lista_sugerida_municipios = Split(s_lista_sugerida_municipios, chr(13))
					iNumeracaoLista=0
					for iCounterLista=LBound(v_lista_sugerida_municipios) to UBound(v_lista_sugerida_municipios)
						if Trim("" & v_lista_sugerida_municipios(iCounterLista)) <> "" then
							iNumeracaoLista=iNumeracaoLista+1
							s_tabela_municipios_IBGE = s_tabela_municipios_IBGE & _
												"	<tr>" & chr(13) & _
												"		<td align='right'>" & chr(13) & _
												"			<span class='N'>&nbsp;" & Cstr(iNumeracaoLista) & "." & "</span>" & chr(13) & _
												"		</td>" & chr(13) & _
												"		<td align='left'>" & chr(13) & _
												"			<span class='N'>" & Trim("" & v_lista_sugerida_municipios(iCounterLista)) & "</span>" & chr(13) & _
												"		</td>" & chr(13) & _
												"	</tr>" & chr(13)
							end if
						next
					
					if s_tabela_municipios_IBGE <> "" then
						s_tabela_municipios_IBGE = _
								"<table cellspacing='0' cellpadding='1'>" & chr(13) & _
								"	<tr>" & chr(13) & _
								"		<td align='center'>" & chr(13) & _
								"			<p class='N'>" & "Relação de municípios de '" & r_cliente.uf & "' que se iniciam com a letra '" & Ucase(left(r_cliente.cidade,1)) & "'" & "</p>" & chr(13) & _
								"		</td>" & chr(13) & _
								"	</tr>" & chr(13) & _
								"	<tr>" & chr(13) & _
								"		<td align='center'>" & chr(13) &_
								"			<table cellspacing='0' border='1'>" & chr(13) & _
												s_tabela_municipios_IBGE & _
								"			</table>" & chr(13) & _
								"		</td>" & chr(13) & _
								"	</tr>" & chr(13) & _
								"</table>" & chr(13)
						end if
					end if
				end if
			end if
		end if

	if alerta = "" then
		if rb_end_entrega = "S" then
		'	MUNICÍPIO DE ACORDO C/ TABELA DO IBGE?
			if Not consiste_municipio_IBGE_ok(EndEtg_cidade, EndEtg_uf, s_lista_sugerida_municipios, msg_erro) then
				if alerta <> "" then alerta = alerta & "<br><br>" & String(80,"=") & "<br><br>"
				if msg_erro <> "" then
					alerta = alerta & msg_erro
				else
					alerta = alerta & "Município '" & EndEtg_cidade & "' não consta na relação de municípios do IBGE para a UF de '" & EndEtg_uf & "'!!"
					if s_lista_sugerida_municipios <> "" then
						alerta = alerta & "<br>" & _
										  "Localize o município na lista abaixo e verifique se a grafia está correta!!"
						v_lista_sugerida_municipios = Split(s_lista_sugerida_municipios, chr(13))
						iNumeracaoLista=0
						for iCounterLista=LBound(v_lista_sugerida_municipios) to UBound(v_lista_sugerida_municipios)
							if Trim("" & v_lista_sugerida_municipios(iCounterLista)) <> "" then
								iNumeracaoLista=iNumeracaoLista+1
								s_tabela_municipios_IBGE = s_tabela_municipios_IBGE & _
													"	<tr>" & chr(13) & _
													"		<td align='right'>" & chr(13) & _
													"			<span class='N'>&nbsp;" & Cstr(iNumeracaoLista) & "." & "</span>" & chr(13) & _
													"		</td>" & chr(13) & _
													"		<td align='left'>" & chr(13) & _
													"			<span class='N'>" & Trim("" & v_lista_sugerida_municipios(iCounterLista)) & "</span>" & chr(13) & _
													"		</td>" & chr(13) & _
													"	</tr>" & chr(13)
								end if
							next

						if s_tabela_municipios_IBGE <> "" then
							s_tabela_municipios_IBGE = _
									"<table cellspacing='0' cellpadding='1'>" & chr(13) & _
									"	<tr>" & chr(13) & _
									"		<td align='center'>" & chr(13) & _
									"			<p class='N'>" & "Relação de municípios de '" & EndEtg_uf & "' que se iniciam com a letra '" & Ucase(left(EndEtg_cidade,1)) & "'" & "</p>" & chr(13) & _
									"		</td>" & chr(13) & _
									"	</tr>" & chr(13) & _
									"	<tr>" & chr(13) & _
									"		<td align='center'>" & chr(13) &_
									"			<table cellspacing='0' border='1'>" & chr(13) & _
													s_tabela_municipios_IBGE & _
									"			</table>" & chr(13) & _
									"		</td>" & chr(13) & _
									"	</tr>" & chr(13) & _
									"</table>" & chr(13)
							end if
						end if
					end if
				end if 'if Not consiste_municipio_IBGE_ok()
			end if 'if rb_end_entrega = "S"
		end if 'if alerta = ""

	dim s_campo_inicial
	if isLojaHabilitadaProdCompostoECommerce(loja) then
		s_campo_inicial = "c_produto"
	else
		s_campo_inicial = "c_fabricante"
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
	<title>LOJA</title>
	</head>



<script src="<%=URL_FILE__JQUERY%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__AJAX_JS%>" language="JavaScript" type="text/javascript"></script>

<script type="text/javascript">
	$(function () {
		$(".tdTitFabr").hide();
		$(".tdDadosFabr").hide();
		$(".tdDadosProd").addClass("ME");
		$("#divAjaxRunning").css('filter', 'alpha(opacity=60)'); // TRANSPARÊNCIA NO IE8

		//Every resize of window
		$(window).resize(function () {
			sizeDivAjaxRunning();
		});

		//Every scroll of window
		$(window).scroll(function () {
			sizeDivAjaxRunning();
		});

	<%if operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO then%>
		for (var i = 0; i < fPED.c_produto.length; i++) {
			if (fPED.c_produto[i].value!="")
			{
				fPED.c_produto[i].value = normaliza_produto(fPED.c_produto[i].value);
				consultaAjaxJQueryDadosProduto(i);
			}
		}
	<% end if %>
	});

	//Dynamically assign height
	function sizeDivAjaxRunning() {
		var newTop = $(window).scrollTop() + "px";
		$("#divAjaxRunning").css("top", newTop);
	}
</script>

<script language="JavaScript" type="text/javascript">
function trataRespostaAjaxJQueryConsultaDadosProduto(xmlResposta) {
var f, strResp, i, j, xmlDoc, oNodes;
var strFabricante,strProduto, strStatus, strPrecoLista, strDescricao, strTabelaOrigem, strMsgErro;
	f=fPED;
	strResp = xmlResposta;
	if (strResp=="") {
		alert("Falha ao consultar a descrição!!");
		window.status="Concluído";
		$("#divAjaxRunning").hide();
		return;
		}
		
	if (strResp!="") {
		try {
			xmlDoc = xmlResposta.documentElement;
			for (i = 0; i < xmlDoc.getElementsByTagName("ItemConsulta").length; i++) {
				//  Fabricante
				oNodes = xmlDoc.getElementsByTagName("fabricante")[i];
				if (oNodes.childNodes.length > 0) strFabricante = oNodes.childNodes[0].nodeValue; else strFabricante = "";
				if (strFabricante == null) strFabricante = "";
				//  Produto
				oNodes = xmlDoc.getElementsByTagName("produto")[i];
				if (oNodes.childNodes.length > 0) strProduto = oNodes.childNodes[0].nodeValue; else strProduto = "";
				if (strProduto == null) strProduto = "";
				// Tabela Origem
				oNodes = xmlDoc.getElementsByTagName("tabela_origem")[i];
				if (oNodes.childNodes.length > 0) strTabelaOrigem = oNodes.childNodes[0].nodeValue; else strTabelaOrigem = "";
				if (strTabelaOrigem == null) strTabelaOrigem = "";
				//  Status
				oNodes = xmlDoc.getElementsByTagName("status")[i];
				if (oNodes.childNodes.length > 0) strStatus = oNodes.childNodes[0].nodeValue; else strStatus = "";
				if (strStatus == null) strStatus = "";
				if (strStatus == "OK") {
					//  Descrição
					oNodes = xmlDoc.getElementsByTagName("descricao")[i];
					if (oNodes.childNodes.length > 0) strDescricao = oNodes.childNodes[0].nodeValue; else strDescricao = "";
					if (strDescricao == null) strDescricao = "";
					if (strDescricao != "") {
						for (j = 0; j < f.c_fabricante.length; j++) {
							if (
							((f.c_fabricante[j].value == strFabricante) && (f.c_produto[j].value == strProduto))
							||
							((f.c_fabricante[j].value == "") && (f.c_produto[j].value == strProduto))
							) {
								//  Percorre o laço até o final para o caso do usuário ter digitado o mesmo produto em várias linhas
								//	(apesar de que isso não será aceito pelas consistências que serão feitas).
								if (f.c_fabricante[j].value == "") f.c_fabricante[j].value = strFabricante;
								f.c_descricao[j].value = strDescricao;
								f.c_fabricante[j].style.color = "black";
								f.c_produto[j].style.color = "black";
							}
						}
					}
					//  Preço
					oNodes = xmlDoc.getElementsByTagName("precoLista")[i];
					if (oNodes.childNodes.length > 0) strPrecoLista = oNodes.childNodes[0].nodeValue; else strPrecoLista = "";
					if (strPrecoLista == null) strPrecoLista = "";
					//  Atualiza o preço
					if ((strPrecoLista == "") && (strTabelaOrigem.toUpperCase() != "T_EC_PRODUTO_COMPOSTO")) {
						alert("Falha na consulta do preço do produto " + strProduto);
					}
					else {
						for (j = 0; j < f.c_fabricante.length; j++) {
							if (
								((f.c_fabricante[j].value == strFabricante) && (f.c_produto[j].value == strProduto))
								||
								((f.c_fabricante[j].value == "") && (f.c_produto[j].value == strProduto))
								) {
								//  Percorre o laço até o final para o caso do usuário ter digitado o mesmo produto em várias linhas
								//	(apesar de que isso não será aceito pelas consistências que serão feitas).
								f.c_preco_lista[j].value = strPrecoLista;
							}
						}
					}
				}
				else {
					//  Mensagem de Erro
					oNodes = xmlDoc.getElementsByTagName("msg_erro")[i];
					if (oNodes.childNodes.length > 0) strMsgErro = oNodes.childNodes[0].nodeValue; else strMsgErro = "";
					if (strMsgErro == null) strMsgErro = "";
					for (j = 0; j < f.c_fabricante.length; j++) {
						//  Percorre o laço até o final para o caso do usuário ter digitado o mesmo produto em várias linhas
						//	(apesar de que isso não será aceito pelas consistências que serão feitas).
						if ((f.c_fabricante[j].value == strFabricante) && (f.c_produto[j].value == strProduto)) {
							f.c_fabricante[j].style.color = COR_AJAX_CONSULTA_DADOS_PRODUTO__INEXISTENTE;
							f.c_produto[j].style.color = COR_AJAX_CONSULTA_DADOS_PRODUTO__INEXISTENTE;
						}
					}
					alert("Falha ao consultar os dados do produto " + strProduto + "\n" + strMsgErro);
				}
			}
		}
		catch (e) {
			alert("Falha na consulta dos dados do produto!!\n" + e.message);
		}
		}
	window.status="Concluído";
	$("#divAjaxRunning").hide();
}

// Esta função foi alterada para executar a requisição Ajax usando jQuery ao invés do objeto XMLHttpRequest
// para atender às requisições em paralelo na inicialização da página quando se trata de pedido de e-commerce.
function consultaAjaxJQueryDadosProduto(intIndice) {
	var f, i, strProdutoSelecionado, strUrl;
	f=fPED;
	if (trim(f.c_produto[intIndice].value)=="") return;

	f.c_fabricante[intIndice].value = "";
	strProdutoSelecionado=f.c_fabricante[intIndice].value + "|" + f.c_produto[intIndice].value;
	
	window.status="Aguarde, consultando descrição ...";
	$("#divAjaxRunning").show();

	$.ajax({
		type: "GET",
		url: "../Global/AjaxConsultaDadosProdutoBD.asp",
		data: "listaProdutos=" + strProdutoSelecionado + "&loja=" + f.c_loja.value,
		cache: false,
		async: true,
		success: function(response) {
			if (response != "") {
				trataRespostaAjaxJQueryConsultaDadosProduto(response);
			}
		},
		error: function(response) {
			alert("Erro ao pesquisar lista de locais!");
		}
	});
}

function trataLimpaLinha(intIndice) {
var f, c, blnProdutoVazio;
	f=fPED;
	blnProdutoVazio=false;
	c = f.c_fabricante[intIndice];
	if ($(c).is(":visible")){
		if ((trim(f.c_fabricante[intIndice].value)=="")&&(trim(f.c_produto[intIndice].value)=="")) blnProdutoVazio=true;
	}
	else {
		if (trim(f.c_produto[intIndice].value)=="") blnProdutoVazio=true;
	}

	if (blnProdutoVazio) {
		f.c_fabricante[intIndice].value="";
		f.c_produto[intIndice].value="";
		f.c_qtde[intIndice].value="";
		f.c_descricao[intIndice].value="";
		f.c_preco_lista[intIndice].value="";
		}
}

function LimparLinha(f, intIdx) {
	f.c_fabricante[intIdx].value = "";
	f.c_produto[intIdx].value = "";
	f.c_qtde[intIdx].value = "";
	f.c_descricao[intIdx].value = "";
	f.c_preco_lista[intIdx].value = "";
	f.c_produto[intIdx].focus();
}

function fPEDConfirma( f ) {
var i, b, ha_item;
	ha_item=false;
	for (i=0; i < f.c_produto.length; i++) {
		b=false;
		if (trim(f.c_fabricante[i].value)!="") b=true;
		if (trim(f.c_produto[i].value)!="") b=true;
		if (trim(f.c_qtde[i].value)!="") b=true;
		
		if (b) {
			ha_item=true;
			if (trim(f.c_produto[i].value)=="") {
				alert("Informe o código do produto!!");
				f.c_produto[i].focus();
				return;
				}
			if (trim(f.c_qtde[i].value)=="") {
				alert("Informe a quantidade!!");
				f.c_qtde[i].focus();
				return;
				}
			if (parseInt(f.c_qtde[i].value)<=0) {
				alert("Quantidade inválida!!");
				f.c_qtde[i].focus();
				return;
				}
			}
		}
		
	if (!ha_item) {
		alert("Não há produtos na lista!!");
		f.c_fabricante[0].focus();
		return;
		}
		
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
<link href="<%=URL_FILE__EPRINTER_CSS%>" rel="stylesheet" type="text/css" media="print">

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
.TdCliLbl
{
	width:130px;
	text-align:right;
}
.TdCliCel
{
	width:520px;
	text-align:left;
}
.TdCliBtn
{
	width:30px;
	text-align:center;
}
</style>

<% if alerta <> "" then %>
<!-- ************************************************************ -->
<!-- **********  PÁGINA PARA EXIBIR MENSAGENS DE ERRO  ********** -->
<!-- ************************************************************ -->
<body>
<center>
<br>
<!--  T E L A  -->
<p class="T">A V I S O</p>
<div class="MtAlerta" style="width:600px;font-weight:bold;" align="center"><P style='margin:5px 2px 5px 2px;'><%=alerta%></p></div>
<% if s_tabela_municipios_IBGE <> "" then %>
	<br /><br />
	<%=s_tabela_municipios_IBGE%>
<% end if %>
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
<body onload="if (trim(fPED.<%=s_campo_inicial%>[0].value)=='') fPED.<%=s_campo_inicial%>[0].focus();">
<center>
    
<form id="fPED" name="fPED" method="post" action="PedidoNovo.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="c_loja" id="c_loja" value='<%=loja%>'>
<input type="hidden" name="cliente_selecionado" id="cliente_selecionado" value='<%=cliente_selecionado%>'>
<input type="hidden" name="rb_end_entrega" id="rb_end_entrega" value='<%=rb_end_entrega%>'>
<input type="hidden" name="EndEtg_endereco" id="EndEtg_endereco" value="<%=EndEtg_endereco%>">
<input type="hidden" name="EndEtg_endereco_numero" id="EndEtg_endereco_numero" value="<%=EndEtg_endereco_numero%>">
<input type="hidden" name="EndEtg_endereco_complemento" id="EndEtg_endereco_complemento" value="<%=EndEtg_endereco_complemento%>">
<input type="hidden" name="EndEtg_bairro" id="EndEtg_bairro" value="<%=EndEtg_bairro%>">
<input type="hidden" name="EndEtg_cidade" id="EndEtg_cidade" value="<%=EndEtg_cidade%>">
<input type="hidden" name="EndEtg_uf" id="EndEtg_uf" value="<%=EndEtg_uf%>">
<input type="hidden" name="EndEtg_cep" id="EndEtg_cep" value="<%=EndEtg_cep%>">
<input type="hidden" name="EndEtg_obs" id="EndEtg_obs" value='<%=EndEtg_obs%>'>
<input type="hidden" name="operacao_origem" id="operacao_origem" value="<%=operacao_origem%>" />
<input type="hidden" name="id_magento_api_pedido_xml" id="id_magento_api_pedido_xml" value="<%=id_magento_api_pedido_xml%>" />
<input type="hidden" name="c_numero_magento" id="c_numero_magento" value="<%=c_numero_magento%>" />
<input type="hidden" name="operationControlTicket" id="operationControlTicket" value="<%=operationControlTicket%>" />
<input type="hidden" name="sessionToken" id="sessionToken" value="<%=sessionToken%>" />

<!-- AJAX EM ANDAMENTO -->
<div id="divAjaxRunning" style="display:none;"><img src="../Imagem/ajax_loader_gray_256.gif" class="AjaxImgLoader"/></div>


<% if operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO then %>
<!--  DADOS DO MAGENTO  -->
<table class="Qx" cellspacing="0">
	<tr style="background-color:azure;">
		<td colspan="3" class="MC MB ME MD" align="center"><span class="N">Dados do Magento (pedido nº <%=c_numero_magento%>)</span></td>
	</tr>
	<tr>
		<td class="MB ME MD TdCliLbl"><span class="PLTd">Cliente</span></td>
		<td class="MB MD TdCliCel">
			<span class="C"><%=s_nome_cliente%></span>
			<% if c_mag_cpf_cnpj_identificado <> "" then %>
			<br /><span class="C"><%=cnpj_cpf_formata(c_mag_cpf_cnpj_identificado)%></span>
			<% end if %>
		</td>
	</tr>
</table>
<% end if %>

<!--  I D E N T I F I C A Ç Ã O   D O   P E D I D O -->
<br />
<table width="749" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Pedido Novo</span></td>
</tr>
</table>
<br>

<!--  R E L A Ç Ã O   D E   P R O D U T O S  -->
<table class="Qx" cellspacing="0">
	<tr bgColor="#FFFFFF">
	<td class="MB tdTitFabr" align="left"><span class="PLTe">Fabr</span></td>
	<td class="MB tdTitProd" align="left"><span class="PLTe">Produto</span></td>
	<td class="MB" align="right"><span class="PLTd">Qtde</span></td>
	<td class="MB" align="left"><span class="PLTe">Descrição</span></td>
	<td class="MB" align="right"><span class="PLTd">VL Unit</span></td>
	<td align="left">&nbsp;</td>
	</tr>
<% for i=1 to MAX_ITENS %>
<%
		s_produto = ""
		s_qtde = ""
		if operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO then
			if idx_map_item <= UBound(v_map_item) then
				s_produto = Trim("" & v_map_item(idx_map_item).sku)
				s_qtde = Trim("" & v_map_item(idx_map_item).qty_ordered)
				end if
			idx_map_item = idx_map_item + 1
			end if
%>
	<tr>
	<td class="MDBE tdDadosFabr" align="left"><input name="c_fabricante" id="c_fabricante" class="PLLe" maxlength="4" style="width:30px;" onkeypress="if (digitou_enter(true)) fPED.c_produto[<%=Cstr(i-1)%>].focus(); filtra_fabricante();" onblur="this.value=normaliza_codigo(this.value,TAM_MIN_FABRICANTE);trataLimpaLinha(<%=Cstr(i-1)%>);"></td>
	<td class="MDB tdDadosProd" align="left"><input name="c_produto" id="c_produto" class="PLLe" maxlength="8" style="width:60px;" onkeypress="if (digitou_enter(true)) fPED.c_qtde[<%=Cstr(i-1)%>].focus(); filtra_produto();" onblur="this.value=normaliza_produto(this.value);consultaAjaxJQueryDadosProduto(<%=Cstr(i-1)%>);trataLimpaLinha(<%=Cstr(i-1)%>);" value="<%=s_produto%>" /></td>
	<td class="MDB" align="right"><input name="c_qtde" id="c_qtde" class="PLLd" maxlength="4" style="width:30px;" onkeypress="if (digitou_enter(true)) {if (<%=Cstr(i)%>==fPED.c_qtde.length) bCONFIRMA.focus(); else fPED.<%=s_campo_inicial%>[<%=Cstr(i)%>].focus();} filtra_numerico();" value="<%=s_qtde%>" /></td>
	<td class="MDB" align="left"><input name="c_descricao" id="c_descricao" class="PLLe" style="width:427px;" readonly tabindex=-1></td>
	<td class="MDB" align="right"><input name="c_preco_lista" id="c_preco_lista" class="PLLd" style="width:62px;" readonly tabindex=-1></td>
	<td align="left">
		<a name="bLimparLinha" href="javascript:LimparLinha(fPED,<%=Cstr(i-1)%>)" title="limpa o conteúdo desta linha"><img src="../botao/botao_x_red.gif" style="vertical-align:bottom;margin-bottom:1px;" width="20" height="20" border="0"></a>
	</td>
	</tr>
<% next %>
</table>

<br>

<!-- ************   SEPARADOR   ************ -->
<table width="749" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc" align="left">&nbsp;</td></tr>
</table>
<br>


<table class="notPrint" width="749" cellspacing="0">
<tr>
	<td align="left"><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back();" title="volta para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
	<td align="right"><div name="dCONFIRMA" id="dCONFIRMA">
		<a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fPEDConfirma(fPED)" title="segue para próxima tela">
		<img src="../botao/proximo.gif" width="176" height="55" border="0"></a></div>
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
	if operacao_origem = OP_ORIGEM__PEDIDO_NOVO_EC_SEMI_AUTO then
		if tMAP_ITEM.State <> 0 then tMAP_ITEM.Close
		set tMAP_ITEM = nothing

		if tMAP_XML.State <> 0 then tMAP_XML.Close
		set tMAP_XML = nothing
		end if

	cn.Close
	set cn = nothing
%>