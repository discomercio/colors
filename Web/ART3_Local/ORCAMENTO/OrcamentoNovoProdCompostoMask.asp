<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/Global.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =================================
'	  OrcamentoNovoProdCompostoMask.asp
'     =================================
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

	dim i, usuario, loja, cliente_selecionado, r_cliente, msg_erro
	usuario = Trim(Session("usuario_atual"))
	loja = Trim(Session("loja_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 
	If (loja = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

	cliente_selecionado = Trim(request("cliente_selecionado"))
	if (cliente_selecionado = "") then Response.Redirect("aviso.asp?id=" & ERR_CLIENTE_NAO_ESPECIFICADO)
	
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))
	
	dim cn
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	dim blnLojaHabilitadaProdCompostoECommerce
	blnLojaHabilitadaProdCompostoECommerce = isLojaHabilitadaProdCompostoECommerce(loja)

	set r_cliente = New cl_CLIENTE
	if Not x_cliente_bd(cliente_selecionado, r_cliente) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
	
	dim rb_end_entrega, EndEtg_endereco, EndEtg_endereco_numero, EndEtg_endereco_complemento
	dim EndEtg_bairro, EndEtg_cidade, EndEtg_uf, EndEtg_cep,EndEtg_obs
	dim EndEtg_email, EndEtg_email_xml, EndEtg_nome, EndEtg_ddd_res, EndEtg_tel_res, EndEtg_ddd_com, EndEtg_tel_com, EndEtg_ramal_com
	dim EndEtg_ddd_cel, EndEtg_tel_cel, EndEtg_ddd_com_2, EndEtg_tel_com_2, EndEtg_ramal_com_2
	dim EndEtg_tipo_pessoa, EndEtg_cnpj_cpf, EndEtg_contribuinte_icms_status, EndEtg_produtor_rural_status
	dim EndEtg_ie, EndEtg_rg
	rb_end_entrega = Trim(Request.Form("rb_end_entrega"))
	EndEtg_endereco = Trim(Request.Form("EndEtg_endereco"))
	EndEtg_endereco_numero = Trim(Request.Form("EndEtg_endereco_numero"))
	EndEtg_endereco_complemento = Trim(Request.Form("EndEtg_endereco_complemento"))
	EndEtg_bairro = Trim(Request.Form("EndEtg_bairro"))
	EndEtg_cidade = Trim(Request.Form("EndEtg_cidade"))
	EndEtg_uf = Trim(Request.Form("EndEtg_uf"))
	EndEtg_cep = Trim(Request.Form("EndEtg_cep"))
    EndEtg_obs = Trim(Request.Form("EndEtg_obs"))
	EndEtg_email = Trim(Request.Form("EndEtg_email"))
	EndEtg_email_xml = Trim(Request.Form("EndEtg_email_xml"))
	EndEtg_nome = Trim(Request.Form("EndEtg_nome"))
	EndEtg_tipo_pessoa = Trim(Request.Form("EndEtg_tipo_pessoa"))
	EndEtg_cnpj_cpf = Trim(Request.Form("EndEtg_cnpj_cpf"))
	EndEtg_contribuinte_icms_status = Trim(Request.Form("EndEtg_contribuinte_icms_status"))
	EndEtg_produtor_rural_status = Trim(Request.Form("EndEtg_produtor_rural_status"))
	EndEtg_ie = Trim(Request.Form("EndEtg_ie"))
	EndEtg_rg = Trim(Request.Form("EndEtg_rg"))

	'Tratamento para obter os dados de telefone conforme os campos exibidos no formulário
	'O objetivo é evitar que dados sejam gravados de forma inconsistente na seguinte situação:
	'	O usuário seleciona o tipo PJ e inicia o preenchimento dos dados do telefone, mas não informa o DDD
	'	Em seguida, altera o tipo para PF e realiza o preenchimento corretamente
	'	Os dados de telefone exibidos p/ o tipo PJ estão inconsistentes e não devem ser gravados no BD, até porque, mesmo que corretos, seriam informações que não pertencem ao contexto selecionado
	EndEtg_ddd_res = ""
	EndEtg_tel_res = ""
	EndEtg_ddd_com = ""
	EndEtg_tel_com = ""
	EndEtg_ramal_com = ""
	EndEtg_ddd_cel = ""
	EndEtg_tel_cel = ""
	EndEtg_ddd_com_2 = ""
	EndEtg_tel_com_2 = ""
	EndEtg_ramal_com_2 = ""

	if EndEtg_tipo_pessoa = ID_PF then
		EndEtg_ddd_res = Trim(Request.Form("EndEtg_ddd_res"))
		EndEtg_tel_res = Trim(Request.Form("EndEtg_tel_res"))
		EndEtg_ddd_cel = Trim(Request.Form("EndEtg_ddd_cel"))
		EndEtg_tel_cel = Trim(Request.Form("EndEtg_tel_cel"))
	else
		EndEtg_ddd_com = Trim(Request.Form("EndEtg_ddd_com"))
		EndEtg_tel_com = Trim(Request.Form("EndEtg_tel_com"))
		EndEtg_ramal_com = Trim(Request.Form("EndEtg_ramal_com"))
		EndEtg_ddd_com_2 = Trim(Request.Form("EndEtg_ddd_com_2"))
		EndEtg_tel_com_2 = Trim(Request.Form("EndEtg_tel_com_2"))
		EndEtg_ramal_com_2 = Trim(Request.Form("EndEtg_ramal_com_2"))
		end if

    dim orcamento_endereco_logradouro, orcamento_endereco_bairro, orcamento_endereco_cidade, orcamento_endereco_uf, orcamento_endereco_cep, orcamento_endereco_numero
    dim orcamento_endereco_complemento, orcamento_endereco_email, orcamento_endereco_email_xml, orcamento_endereco_nome, orcamento_endereco_ddd_res
    dim orcamento_endereco_tel_res, orcamento_endereco_ddd_com, orcamento_endereco_tel_com, orcamento_endereco_ramal_com, orcamento_endereco_ddd_cel
    dim orcamento_endereco_tel_cel, orcamento_endereco_ddd_com_2, orcamento_endereco_tel_com_2, orcamento_endereco_ramal_com_2, orcamento_endereco_tipo_pessoa
    dim orcamento_endereco_cnpj_cpf, orcamento_endereco_contribuinte_icms_status, orcamento_endereco_produtor_rural_status, orcamento_endereco_ie
    dim orcamento_endereco_rg, orcamento_endereco_contato
    orcamento_endereco_logradouro = Trim(Request.Form("orcamento_endereco_logradouro"))
    orcamento_endereco_bairro = Trim(Request.Form("orcamento_endereco_bairro"))
    orcamento_endereco_cidade = Trim(Request.Form("orcamento_endereco_cidade"))
    orcamento_endereco_uf = Trim(Request.Form("orcamento_endereco_uf"))
    orcamento_endereco_cep = Trim(Request.Form("orcamento_endereco_cep"))
    orcamento_endereco_numero = Trim(Request.Form("orcamento_endereco_numero"))
    orcamento_endereco_complemento = Trim(Request.Form("orcamento_endereco_complemento"))
    orcamento_endereco_email = Trim(Request.Form("orcamento_endereco_email"))
    orcamento_endereco_email_xml = Trim(Request.Form("orcamento_endereco_email_xml"))
    orcamento_endereco_nome = Trim(Request.Form("orcamento_endereco_nome"))
    orcamento_endereco_ddd_res = Trim(Request.Form("orcamento_endereco_ddd_res"))
    orcamento_endereco_tel_res = Trim(Request.Form("orcamento_endereco_tel_res"))
    orcamento_endereco_ddd_com = Trim(Request.Form("orcamento_endereco_ddd_com"))
    orcamento_endereco_tel_com = Trim(Request.Form("orcamento_endereco_tel_com"))
    orcamento_endereco_ramal_com = Trim(Request.Form("orcamento_endereco_ramal_com"))
    orcamento_endereco_ddd_cel = Trim(Request.Form("orcamento_endereco_ddd_cel"))
    orcamento_endereco_tel_cel = Trim(Request.Form("orcamento_endereco_tel_cel"))
    orcamento_endereco_ddd_com_2 = Trim(Request.Form("orcamento_endereco_ddd_com_2"))
    orcamento_endereco_tel_com_2 = Trim(Request.Form("orcamento_endereco_tel_com_2"))
    orcamento_endereco_ramal_com_2 = Trim(Request.Form("orcamento_endereco_ramal_com_2"))
    orcamento_endereco_tipo_pessoa = Trim(Request.Form("orcamento_endereco_tipo_pessoa"))
    orcamento_endereco_cnpj_cpf = Trim(Request.Form("orcamento_endereco_cnpj_cpf"))
    orcamento_endereco_contribuinte_icms_status = Trim(Request.Form("orcamento_endereco_contribuinte_icms_status"))
    orcamento_endereco_produtor_rural_status = Trim(Request.Form("orcamento_endereco_produtor_rural_status"))
    orcamento_endereco_ie = Trim(Request.Form("orcamento_endereco_ie"))
    orcamento_endereco_rg = Trim(Request.Form("orcamento_endereco_rg"))
    orcamento_endereco_contato = Trim(Request.Form("orcamento_endereco_contato"))

	dim alerta
	alerta = ""

	if Trim(orcamento_endereco_numero) = "" then
		Response.Redirect("aviso.asp?id=" & ERR_CAD_CLIENTE_ENDERECO_NUMERO_NAO_PREENCHIDO)
	elseif Len(Trim(orcamento_endereco_logradouro)) > CLng(MAX_TAMANHO_CAMPO_ENDERECO) then
		Response.Redirect("aviso.asp?id=" & ERR_CAD_CLIENTE_ENDERECO_EXCEDE_TAMANHO_MAXIMO)
		end if
		
'	CONSISTÊNCIAS P/ EMISSÃO DE NFe
	dim s_lista_sugerida_municipios
	dim v_lista_sugerida_municipios
	dim iCounterLista, iNumeracaoLista
	dim s_tabela_municipios_IBGE
	s_tabela_municipios_IBGE = ""
	if alerta = "" then
	'	DDD VÁLIDO?
		if Not ddd_ok(orcamento_endereco_ddd_res) then
			if alerta <> "" then alerta = alerta & "<br><br>" & String(80,"=") & "<br><br>"
			alerta = alerta & "DDD do telefone residencial é inválido!!"
			end if
			
		if Not ddd_ok(orcamento_endereco_ddd_com) then
			if alerta <> "" then alerta = alerta & "<br><br>" & String(80,"=") & "<br><br>"
			alerta = alerta & "DDD do telefone comercial é inválido!!"
			end if
			
	'	I.E. É VÁLIDA?
		if orcamento_endereco_tipo_pessoa = ID_PJ then
            if orcamento_endereco_ie <> "" then
			    if Not isInscricaoEstadualValida(orcamento_endereco_ie, orcamento_endereco_uf) then
				    if alerta <> "" then alerta = alerta & "<br><br>" & String(80,"=") & "<br><br>"
				    alerta=alerta & "Corrija a IE (Inscrição Estadual) com um número válido!!" & _
						    "<br>" & "Certifique-se de que a UF informada corresponde à UF responsável pelo registro da IE."
				    end if
            end if
		end if

	'	MUNICÍPIO DE ACORDO C/ TABELA DO IBGE?
		if Not consiste_municipio_IBGE_ok(orcamento_endereco_cidade, orcamento_endereco_uf, s_lista_sugerida_municipios, msg_erro) then
			if alerta <> "" then alerta = alerta & "<br><br>" & String(80,"=") & "<br><br>"
			if msg_erro <> "" then
				alerta = alerta & msg_erro
			else
				alerta = alerta & "Município '" & orcamento_endereco_cidade & "' não consta na relação de municípios do IBGE para a UF de '" & orcamento_endereco_uf & "'!!"
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
								"			<p class='N'>" & "Relação de municípios de '" & orcamento_endereco_uf & "' que se iniciam com a letra '" & Ucase(left(orcamento_endereco_cidade,1)) & "'" & "</p>" & chr(13) & _
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
			if EndEtg_ie <> "" then
				if Not isInscricaoEstadualValida(EndEtg_ie, EndEtg_uf) then
					alerta="Endereço de entrega: preencha a IE (Inscrição Estadual) com um número válido!!" & _
							"<br>" & "Certifique-se de que a UF do endereço de entrega corresponde à UF responsável pelo registro da IE."
					end if
				end if

		'	MUNICÍPIO DE ACORDO C/ TABELA DO IBGE?
			if Not consiste_municipio_IBGE_ok(EndEtg_cidade, EndEtg_uf, s_lista_sugerida_municipios, msg_erro) then
				if alerta <> "" then alerta = alerta & "<br><br>" & String(80,"=") & "<br><br>"
				if msg_erro <> "" then
					alerta = alerta & msg_erro
				else
					alerta = alerta & "Município '" & EndEtg_cidade & "' não consta na relação de municípios do IBGE para a UF de '" & EndEtg_uf & "'!"
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

	if alerta = "" then
	'	Validação do DDD dos telefones
		if orcamento_endereco_tipo_pessoa = ID_PF then
			if (orcamento_endereco_tel_res <> "") And (Len(orcamento_endereco_ddd_res) < 2) then
				alerta=texto_add_br(alerta)
				alerta=alerta & "DDD inválido para o telefone (endereço de cobrança): " & orcamento_endereco_tel_res
				end if

			if (orcamento_endereco_tel_cel <> "") And (Len(orcamento_endereco_ddd_cel) < 2) then
				alerta=texto_add_br(alerta)
				alerta=alerta & "DDD inválido para o telefone (endereço de cobrança): " & orcamento_endereco_tel_cel
				end if

			if (orcamento_endereco_tel_com <> "") And (Len(orcamento_endereco_ddd_com) < 2) then
				alerta=texto_add_br(alerta)
				alerta=alerta & "DDD inválido para o telefone (endereço de cobrança): " & orcamento_endereco_tel_com
				end if
		else
			if (orcamento_endereco_tel_com <> "") And (Len(orcamento_endereco_ddd_com) < 2) then
				alerta=texto_add_br(alerta)
				alerta=alerta & "DDD inválido para o telefone (endereço de cobrança): " & orcamento_endereco_tel_com
				end if
			
			if (orcamento_endereco_tel_com_2 <> "") And (Len(orcamento_endereco_ddd_com_2) < 2) then
				alerta=texto_add_br(alerta)
				alerta=alerta & "DDD inválido para o telefone (endereço de cobrança): " & orcamento_endereco_tel_com_2
				end if
			end if 'if orcamento_endereco_tipo_pessoa = ID_PF

		if rb_end_entrega = "S" then
			if EndEtg_tipo_pessoa = ID_PF then
				if (EndEtg_tel_res <> "") And (Len(EndEtg_ddd_res) < 2) then
					alerta=texto_add_br(alerta)
					alerta=alerta & "DDD inválido para o telefone (endereço de entrega): " & EndEtg_tel_res
					end if

				if (EndEtg_tel_cel <> "") And (Len(EndEtg_ddd_cel) < 2) then
					alerta=texto_add_br(alerta)
					alerta=alerta & "DDD inválido para o telefone (endereço de entrega): " & EndEtg_tel_cel
					end if
			else
				if (EndEtg_tel_com <> "") And (Len(EndEtg_ddd_com) < 2) then
					alerta=texto_add_br(alerta)
					alerta=alerta & "DDD inválido para o telefone (endereço de entrega): " & EndEtg_tel_com
					end if

				if (EndEtg_tel_com_2 <> "") And (Len(EndEtg_ddd_com_2) < 2) then
					alerta=texto_add_br(alerta)
					alerta=alerta & "DDD inválido para o telefone (endereço de entrega): " & EndEtg_tel_com_2
					end if
				end if 'if EndEtg_tipo_pessoa = ID_PF
			end if 'if rb_end_entrega = "S"
		end if

	dim s_campo_inicial
	if blnLojaHabilitadaProdCompostoECommerce then
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
</script>

<script language="JavaScript" type="text/javascript">
var objAjaxConsultaDadosProduto;

function trataRespostaAjaxConsultaDadosProduto() {
var f, strResp, i, j, xmlDoc, oNodes;
var strFabricante,strProduto, strStatus, strPrecoLista, strDescricao, strTabelaOrigem, strMsgErro;
	f=fPED;
	if (objAjaxConsultaDadosProduto.readyState == AJAX_REQUEST_IS_COMPLETE) {
		strResp = objAjaxConsultaDadosProduto.responseText;
		if (strResp=="") {
			alert("Falha ao consultar a descrição!!");
			window.status="Concluído";
			$("#divAjaxRunning").hide();
			return;
			}
		
		if (strResp!="") {
			try {
				xmlDoc = objAjaxConsultaDadosProduto.responseXML.documentElement;
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
}

function consultaDadosProduto(intIndice) {
var f, i, strProdutoSelecionado, strUrl;
	f=fPED;
	if (trim(f.c_produto[intIndice].value)=="") return;

	objAjaxConsultaDadosProduto = GetXmlHttpObject();
	if (objAjaxConsultaDadosProduto == null) {
		alert("O browser NÃO possui suporte ao AJAX!!");
		return;
		}
		
	f.c_fabricante[intIndice].value = "";
	strProdutoSelecionado=f.c_fabricante[intIndice].value + "|" + f.c_produto[intIndice].value;
	
	window.status="Aguarde, consultando descrição ...";
	$("#divAjaxRunning").show();
	
	strUrl = "../Global/AjaxConsultaDadosProdutoBD.asp";
	strUrl += "?listaProdutos=" + strProdutoSelecionado;
	strUrl += "&loja=" + f.c_loja.value;
	//  Prevents server from using a cached file
	strUrl=strUrl+"&sid="+Math.random()+Math.random();
	objAjaxConsultaDadosProduto.onreadystatechange = trataRespostaAjaxConsultaDadosProduto;
	objAjaxConsultaDadosProduto.open("GET", strUrl, true);
	objAjaxConsultaDadosProduto.send(null);
}

function trataLimpaLinha(intIndice) {
var f;
	f=fPED;
	if ((trim(f.c_fabricante[intIndice].value)=="")&&(trim(f.c_produto[intIndice].value)=="")) {
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
    
<form id="fPED" name="fPED" method="post" action="OrcamentoNovo.asp">
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

<!--  CAMPOS ADICIONAIS DO ENDERECO DE ENTREGA  -->
<input type="hidden" name="EndEtg_email" id="EndEtg_email" value="<%=EndEtg_email%>" />
<input type="hidden" name="EndEtg_email_xml" id="EndEtg_email_xml" value="<%=EndEtg_email_xml%>" />
<input type="hidden" name="EndEtg_nome" id="EndEtg_nome" value="<%=EndEtg_nome%>" />
<input type="hidden" name="EndEtg_ddd_res" id="EndEtg_ddd_res" value="<%=EndEtg_ddd_res%>" />
<input type="hidden" name="EndEtg_tel_res" id="EndEtg_tel_res" value="<%=EndEtg_tel_res%>" />
<input type="hidden" name="EndEtg_ddd_com" id="EndEtg_ddd_com" value="<%=EndEtg_ddd_com%>" />
<input type="hidden" name="EndEtg_tel_com" id="EndEtg_tel_com" value="<%=EndEtg_tel_com%>" />
<input type="hidden" name="EndEtg_ramal_com" id="EndEtg_ramal_com" value="<%=EndEtg_ramal_com%>" />
<input type="hidden" name="EndEtg_ddd_cel" id="EndEtg_ddd_cel" value="<%=EndEtg_ddd_cel%>" />
<input type="hidden" name="EndEtg_tel_cel" id="EndEtg_tel_cel" value="<%=EndEtg_tel_cel%>" />
<input type="hidden" name="EndEtg_ddd_com_2" id="EndEtg_ddd_com_2" value="<%=EndEtg_ddd_com_2%>" />
<input type="hidden" name="EndEtg_tel_com_2" id="EndEtg_tel_com_2" value="<%=EndEtg_tel_com_2%>" />
<input type="hidden" name="EndEtg_ramal_com_2" id="EndEtg_ramal_com_2" value="<%=EndEtg_ramal_com_2%>" />
<input type="hidden" name="EndEtg_tipo_pessoa" id="EndEtg_tipo_pessoa" value="<%=EndEtg_tipo_pessoa%>" />
<input type="hidden" name="EndEtg_cnpj_cpf" id="EndEtg_cnpj_cpf" value="<%=EndEtg_cnpj_cpf%>" />
<input type="hidden" name="EndEtg_contribuinte_icms_status" id="EndEtg_contribuinte_icms_status" value="<%=EndEtg_contribuinte_icms_status%>" />
<input type="hidden" name="EndEtg_produtor_rural_status" id="EndEtg_produtor_rural_status" value="<%=EndEtg_produtor_rural_status%>" />
<input type="hidden" name="EndEtg_ie" id="EndEtg_ie" value="<%=EndEtg_ie%>" />
<input type="hidden" name="EndEtg_rg" id="EndEtg_rg" value="<%=EndEtg_rg%>" />

<input type="hidden" name="orcamento_endereco_logradouro" id="orcamento_endereco_logradouro" value="<%=orcamento_endereco_logradouro%>" />
<input type="hidden" name="orcamento_endereco_bairro" id="orcamento_endereco_bairro" value="<%=orcamento_endereco_bairro%>" />
<input type="hidden" name="orcamento_endereco_cidade" id="orcamento_endereco_cidade" value="<%=orcamento_endereco_cidade%>" />
<input type="hidden" name="orcamento_endereco_uf" id="orcamento_endereco_uf" value="<%=orcamento_endereco_uf%>" />
<input type="hidden" name="orcamento_endereco_cep" id="orcamento_endereco_cep" value="<%=orcamento_endereco_cep%>" />
<input type="hidden" name="orcamento_endereco_numero" id="orcamento_endereco_numero" value="<%=orcamento_endereco_numero%>" />
<input type="hidden" name="orcamento_endereco_complemento" id="orcamento_endereco_complemento" value="<%=orcamento_endereco_complemento%>" />
<input type="hidden" name="orcamento_endereco_email" id="orcamento_endereco_email" value="<%=orcamento_endereco_email%>" />
<input type="hidden" name="orcamento_endereco_email_xml" id="orcamento_endereco_email_xml" value="<%=orcamento_endereco_email_xml%>" />
<input type="hidden" name="orcamento_endereco_nome" id="orcamento_endereco_nome" value="<%=orcamento_endereco_nome%>" />
<input type="hidden" name="orcamento_endereco_ddd_res" id="orcamento_endereco_ddd_res" value="<%=orcamento_endereco_ddd_res%>" />
<input type="hidden" name="orcamento_endereco_tel_res" id="orcamento_endereco_tel_res" value="<%=orcamento_endereco_tel_res%>" />
<input type="hidden" name="orcamento_endereco_ddd_com" id="orcamento_endereco_ddd_com" value="<%=orcamento_endereco_ddd_com%>" />
<input type="hidden" name="orcamento_endereco_tel_com" id="orcamento_endereco_tel_com" value="<%=orcamento_endereco_tel_com%>" />
<input type="hidden" name="orcamento_endereco_ramal_com" id="orcamento_endereco_ramal_com" value="<%=orcamento_endereco_ramal_com%>" />
<input type="hidden" name="orcamento_endereco_ddd_cel" id="orcamento_endereco_ddd_cel" value="<%=orcamento_endereco_ddd_cel%>" />
<input type="hidden" name="orcamento_endereco_tel_cel" id="orcamento_endereco_tel_cel" value="<%=orcamento_endereco_tel_cel%>" />
<input type="hidden" name="orcamento_endereco_ddd_com_2" id="orcamento_endereco_ddd_com_2" value="<%=orcamento_endereco_ddd_com_2%>" />
<input type="hidden" name="orcamento_endereco_tel_com_2" id="orcamento_endereco_tel_com_2" value="<%=orcamento_endereco_tel_com_2%>" />
<input type="hidden" name="orcamento_endereco_ramal_com_2" id="orcamento_endereco_ramal_com_2" value="<%=orcamento_endereco_ramal_com_2%>" />
<input type="hidden" name="orcamento_endereco_tipo_pessoa" id="orcamento_endereco_tipo_pessoa" value="<%=orcamento_endereco_tipo_pessoa%>" />
<input type="hidden" name="orcamento_endereco_cnpj_cpf" id="orcamento_endereco_cnpj_cpf" value="<%=orcamento_endereco_cnpj_cpf%>" />
<input type="hidden" name="orcamento_endereco_contribuinte_icms_status" id="orcamento_endereco_contribuinte_icms_status" value="<%=orcamento_endereco_contribuinte_icms_status%>" />
<input type="hidden" name="orcamento_endereco_produtor_rural_status" id="orcamento_endereco_produtor_rural_status" value="<%=orcamento_endereco_produtor_rural_status%>" />
<input type="hidden" name="orcamento_endereco_ie" id="orcamento_endereco_ie" value="<%=orcamento_endereco_ie%>" />
<input type="hidden" name="orcamento_endereco_rg" id="orcamento_endereco_rg" value="<%=orcamento_endereco_rg%>" />
<input type="hidden" name="orcamento_endereco_contato" id="orcamento_endereco_contato" value="<%=orcamento_endereco_contato%>" />


<!-- AJAX EM ANDAMENTO -->
<div id="divAjaxRunning" style="display:none;"><img src="../Imagem/ajax_loader_gray_256.gif" class="AjaxImgLoader"/></div>


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
	<tr>
	<td class="MDBE tdDadosFabr" align="left"><input name="c_fabricante" id="c_fabricante" class="PLLe" maxlength="4" style="width:30px;" onkeypress="if (digitou_enter(true)) fPED.c_produto[<%=Cstr(i-1)%>].focus(); filtra_fabricante();" onblur="this.value=normaliza_codigo(this.value,TAM_MIN_FABRICANTE);trataLimpaLinha(<%=Cstr(i-1)%>);"></td>
	<td class="MDB tdDadosProd" align="left"><input name="c_produto" id="c_produto" class="PLLe" maxlength="8" style="width:60px;" onkeypress="if (digitou_enter(true)) fPED.c_qtde[<%=Cstr(i-1)%>].focus(); filtra_produto();" onblur="this.value=normaliza_produto(this.value);consultaDadosProduto(<%=Cstr(i-1)%>);trataLimpaLinha(<%=Cstr(i-1)%>);"></td>
	<td class="MDB" align="right"><input name="c_qtde" id="c_qtde" class="PLLd" maxlength="4" style="width:30px;" onkeypress="if (digitou_enter(true)) {if (<%=Cstr(i)%>==fPED.c_qtde.length) bCONFIRMA.focus(); else fPED.<%=s_campo_inicial%>[<%=Cstr(i)%>].focus();} filtra_numerico();"></td>
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
	cn.Close
	set cn = nothing
%>