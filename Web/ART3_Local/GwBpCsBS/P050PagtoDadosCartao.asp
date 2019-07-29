<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/Global.asp"    -->
<!-- #include file = "../global/Braspag.asp"    -->
<!-- #include file = "../global/BraspagCS.asp"    -->

<%
'     ===========================================
'	  P050PagtoDadosCartao.asp
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


	On Error GoTo 0
	Err.Clear

	dim alerta
	alerta = ""
	
	dim s, usuario, loja, pedido_selecionado, id_pedido_base

	usuario = BRASPAG_USUARIO_CLIENTE

	dim cnpj_cpf_selecionado
	cnpj_cpf_selecionado = retorna_so_digitos(Request("cnpj_cpf_selecionado"))

	pedido_selecionado = ucase(Trim(request("pedido_selecionado")))
	if (pedido_selecionado = "") then Response.Redirect("aviso.asp?id=" & ERR_PEDIDO_NAO_ESPECIFICADO)
	s = normaliza_num_pedido(pedido_selecionado)
	if s <> "" then pedido_selecionado = s
	id_pedido_base = retorna_num_pedido_base(pedido_selecionado)
	
	dim i_cartao
	dim c_qtde_cartoes, qtde_cartoes
	c_qtde_cartoes = Trim(Request.Form("c_qtde_cartoes"))
	qtde_cartoes = converte_numero(c_qtde_cartoes)
	if qtde_cartoes = 0 then Response.Redirect("aviso.asp?id=" & ERR_QTDE_CARTOES_INVALIDA)

	dim strScriptWindowName
	strScriptWindowName = _
				"<script language='JavaScript'>" & chr(13) & _
				"	window.name = '" & SITE_CLIENTE_TITULO_JANELA & "';" & chr(13) & _
				"</script>" & chr(13)

'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs, msg_erro
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	
'	INCLUI UM SUFIXO NO Nº DO PEDIDO C/ O NÚMERO DA TENTATIVA DE TRANSAÇÃO DE PAGAMENTO (A 1ª TRANSAÇÃO É ENVIADA SEM O SUFIXO)
'	ISSO TEM DOIS OBJETIVOS:
'		1) NOS CASOS EM QUE O PEDIDO É PAGO C/ MAIS DE UM CARTÃO, ISSO DIFERENCIA O Nº DO PEDIDO ENTRE AS TRANSAÇÕES
'			E PODE AJUDAR A NÃO SER INTERPRETADO COMO FRAUDE.
'		2) QUANDO OCORRE ERRO NO PROCESSAMENTO DA RESPOSTA E O CAMPO 'BraspagTransactionId' NÃO É ARMAZENADO NO BD,
'			O USO DE UM IDENTIFICADOR DE PEDIDO ÚNICO NO CAMPO 'OrderId' VIABILIZA A UTILIZAÇÃO DO MÉTODO 'GetOrderIdData'
'			PARA RECUPERAR O VALOR DE 'BraspagTransactionId'.
'	LEMBRANDO QUE O CAMPO 'BraspagTransactionId' É NECESSÁRIO P/ CONSULTAR O STATUS ATUALIZADO DA TRANSAÇÃO E TAMBÉM
'	REALIZAR O CANCELAMENTO, ESTORNO, ETC.
	dim intSufixoNsu, pedido_com_sufixo_nsu
	pedido_com_sufixo_nsu = id_pedido_base
	intSufixoNsu = BraspagCSGeraSufixoPedidoNsuPag(id_pedido_base, usuario)
	if intSufixoNsu > 1 then pedido_com_sufixo_nsu = id_pedido_base & "_" & Cstr(intSufixoNsu)

	dim owner
	owner = BraspagObtemOwnerPeloPedido(id_pedido_base)

	dim vl_pagto_em_cartao, vl_pagador, vl_saldo_a_pagar
	vl_pagto_em_cartao = calcula_vl_pagto_em_cartao(id_pedido_base, msg_erro)
	if msg_erro <> "" then
		alerta=texto_add_br(alerta)
		alerta=alerta & msg_erro
		end if

	vl_pagador = BraspagCSCalculaValorPagadorAutorizadoCapturadoFamilia(id_pedido_base, msg_erro)
	if msg_erro <> "" then
		alerta=texto_add_br(alerta)
		alerta=alerta & msg_erro
		end if

	vl_saldo_a_pagar = vl_pagto_em_cartao - vl_pagador
	if vl_saldo_a_pagar < 0 then vl_saldo_a_pagar = 0

	dim FingerPrint_SessionID, idPagtoGwAfSessionID
	FingerPrint_SessionID = gera_FingerPrint_SessionID
	if Not fin_gera_nsu(T_PAGTO_GW_AF_SESSIONID, idPagtoGwAfSessionID, msg_erro) then
		alerta=texto_add_br(alerta)
		alerta=alerta & "FALHA AO GERAR NSU PARA O NOVO REGISTRO DE FINGERPRINT.SESSIONID (" & msg_erro & ")"
	elseif idPagtoGwAfSessionID <= 0 then
		alerta=texto_add_br(alerta)
		alerta=alerta & "NSU GERADO É INVÁLIDO (" & idPagtoGwAfSessionID & ")"
		end if

	if alerta = "" then
		s = "SELECT * FROM t_PAGTO_GW_AF_SESSIONID WHERE (id = -1)"
		if rs.State <> 0 then rs.Close
		rs.open s, cn
		rs.AddNew
		rs("id") = idPagtoGwAfSessionID
		rs("pedido") = id_pedido_base
		rs("pedido_com_sufixo_nsu") = pedido_com_sufixo_nsu
		rs("FingerPrint_SessionID") = FingerPrint_SessionID
		rs("usuario") = usuario
		rs("executado_pelo_cliente_status") = 1
		rs("origem_endereco_IP") = Trim(Request.ServerVariables("REMOTE_ADDR"))
		rs.Update
		end if

	dim i, bandeira, vBandeiraLogo, vBandeiraParcMin
	vBandeiraLogo = BraspagArrayBandeiras
	dim strScript
	strScript = "<script language='JavaScript' type='text/javascript'>" & chr(13) & _
					"var vl_saldo_a_pagar=" & js_formata_numero(vl_saldo_a_pagar) & ";" & chr(13) & _
					"var vBandeiraLogo=[];" & chr(13) & _
					"vBandeiraLogo['']='';" & chr(13)

	for i=LBound(vBandeiraLogo) to UBound(vBandeiraLogo)
		bandeira = UCase(Trim("" & vBandeiraLogo(i)))
		strScript = strScript & _
					"vBandeiraLogo['" & bandeira & "']='" & BraspagObtemNomeArquivoLogoOpcao(bandeira) & "';" & chr(13)
		next

	strScript = strScript & _
				"vBandeiraDescricao=[];" & chr(13) & _
				"vBandeiraDescricao['']='';" & chr(13)
	for i=LBound(vBandeiraLogo) to UBound(vBandeiraLogo)
		bandeira = UCase(Trim("" & vBandeiraLogo(i)))
		strScript = strScript & _
					"vBandeiraDescricao['" & bandeira & "']='" & BraspagDescricaoBandeira(bandeira) & "';" & chr(13)
		next

'	VALOR MÍNIMO DO PAGAMENTO
	strScript = strScript & _
					"var vBandeiraParcMin=[];" & chr(13) & _
					"vBandeiraParcMin['']=0;" & chr(13)
	vBandeiraParcMin = BraspagArrayBandeiras
	for i=LBound(vBandeiraParcMin) to UBound(vBandeiraParcMin)
		if Trim("" & vBandeiraParcMin(i)) <> "" then
			s = "SELECT" & _
					" Min(vl_min_parcela) AS valor_minimo" & _
				" FROM t_PRAZO_PAGTO_VISANET" & _
				" WHERE" & _
					" (tipo IN ('" & BraspagObtemIdRegistroBdPrazoPagtoLoja(vBandeiraParcMin(i)) & "', '" & BraspagObtemIdRegistroBdPrazoPagtoEmissor(vBandeiraParcMin(i)) & "'))"
			if rs.State <> 0 then rs.Close
			rs.open s, cn
			if Not rs.Eof then
				strScript = strScript & _
							"vBandeiraParcMin['" & UCase(Trim("" & vBandeiraParcMin(i))) & "']=" & js_formata_numero(rs("valor_minimo")) & ";" & chr(13)
			else
				strScript = strScript & _
							"vBandeiraParcMin['" & UCase(Trim("" & vBandeiraParcMin(i))) & "']=-1;" & chr(13)
				end if
			end if
		next
	
	strScript = strScript & _
				"</script>" & chr(13)

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
	<title><%=SITE_CLIENTE_TITULO_JANELA%></title>
</head>


<% if False then 'APENAS P/ HABILITAR O INTELLISENSE DURANTE O DESENVOLVIMENTO!! %>
<script src="../Global/jquery.js" language="JavaScript" type="text/javascript"></script>
<% end if %>

<script src="<%=URL_FILE__JQUERY%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_MASKMONEY%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_I18N%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__SSL_JS%>" Language="JavaScript" Type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
	var qtde_cartoes = <%=qtde_cartoes%>;
</script>

<% =strScriptWindowName %>

<% Response.Write strScript %>

<script language="JavaScript" type="text/javascript">
function obtemIndiceCampo(c) {
	var v, s_id, s_index;
	s_id = $(c).attr('id');
	v = s_id.split('_');
	s_index = v[v.length-1];
	return s_index;
}

function exibeLogoBandeira(indice){
	var s_logo, s_bandeira_id, s_img_id;
	s_bandeira_id = "#c_cartao_bandeira_" + indice;
	s_img_id = "#c_cartao_bandeira_logo_" + indice;
	s_logo = vBandeiraLogo[$(s_bandeira_id).val().toUpperCase()];
	if (s_logo.length > 0) {
		s_logo = "../Imagem/Braspag/" + s_logo;
		$(s_img_id).attr("src", s_logo);
	}
}
</script>

<script language="JavaScript" type="text/javascript">
	var v_TxtValor_anterior = [];
	var s_bandeira_id, s_opcao_parcelamento_id, s_valor_id, s_id_aux, indice;

	$(document).ready(function() {
		var s_id;

		$("#divAjaxRunning").hide(); // Mantém oculto inicialmente

		$(document).ajaxStart(function() {
			$("#divAjaxRunning").show();
		})
		.ajaxStop(function() {
			$("#divAjaxRunning").hide();
		});

		$("#divAjaxRunning").css('filter', 'alpha(opacity=60)'); // TRANSPARÊNCIA NO IE8

		//Every resize of window
		$(window).resize(function() {
			sizeDivAjaxRunning();
		});

		//Every scroll of window
		$(window).scroll(function() {
			sizeDivAjaxRunning();
		});

		for (var i = 1; i <= qtde_cartoes; i++) {
			s_id = "#c_cartao_valor_" + i;
			$(s_id).maskMoney({allowNegative: false, thousands:'.', decimal:',', affixesStay: false});
		}

		$(".TxtValor").focus(function(){
			var s_index;
			s_index = obtemIndiceCampo($(this));
			v_TxtValor_anterior[s_index] = $(this).val();
		});

		$(".TxtValor").blur(function(){
			var s_index, s_bandeira_id;
			s_index = obtemIndiceCampo($(this));
			s_bandeira_id = "#c_cartao_bandeira_" + s_index;
			if (v_TxtValor_anterior[s_index].toString() != $(this).val()) {
				obtemOpcoesParcelamento(s_index, $(s_bandeira_id).val(),'<%=id_pedido_base%>', $(this).val());
			}
		});

		$(".TxtValor").change(function(){
			atualizaTotalDestaTransacao();
		});

		$(".TxtValor").keyup(function(){
			atualizaTotalDestaTransacao();
		});

		$(".TxtValor").blur(function(){
			atualizaTotalDestaTransacao();
		});

		$(".SelBand").change(function(){
			var s_index, s_valor_id;
			s_index = obtemIndiceCampo($(this));
			exibeLogoBandeira(s_index);
			s_valor_id = "#c_cartao_valor_" + s_index;
			obtemOpcoesParcelamento(s_index, $(this).val(),'<%=id_pedido_base%>', $(s_valor_id).val());
		});

		$(".SelParc").focus(function(){
			var s_index, s_bandeira_id, s_valor_id;
			s_index = obtemIndiceCampo($(this));
			s_bandeira_id = "#c_cartao_bandeira_" + s_index;
			s_valor_id = "#c_cartao_valor_" + s_index;
			if ($(s_bandeira_id).val().length == 0) {
				alert("Selecione primeiro a bandeira do cartão!!");
				$(s_bandeira_id).focus();
				return;
			}
			if (($(s_valor_id).val().length == 0)||(converte_numero($(s_valor_id).val()) == 0)){
				alert("Informe antes o valor a ser pago com este cartão!!");
				$(s_valor_id).focus();
				return;
			}
			if ($(this).length==0) {
				obtemOpcoesParcelamento(s_index, $(s_bandeira_id).val(),'<%=id_pedido_base%>', $(s_valor_id).val());
			}
		});

		$(".CodSeguranca").mouseover(function() {
			var obj = this.offsetParent;
			var esq = 0;
			var topo = 0;

			while (true) {
				if (!obj.offsetParent) {
					break;
				}

				esq += obj.offsetLeft;
				topo += obj.offsetTop;

				obj = obj.offsetParent;
			}

			$("#divCodSegInfo").css({
				left: (esq - 300) + 'px',
				top: (topo - 200) + 'px'
			});

			$("#divCodSegInfo").show();
		}).mouseout(function() {
			$("#divCodSegInfo").hide();
		});

		restauraDadosMemorizados();
		atualizaTotalDestaTransacao();
		if ($("#c_vl_total_desta_transacao").text()==formata_moeda(0)) $("#c_vl_total_desta_transacao").text("");

		// TRATAMENTO PARA HISTORY.BACK
		for (var i = 1; i <= qtde_cartoes; i++) {
			exibeLogoBandeira(i);
			s_bandeira_id = "#c_cartao_bandeira_" + i;
			s_opcao_parcelamento_id = "#c_opcao_parcelamento_" + i;
			if (($(s_bandeira_id).val().length > 0) && ($(s_opcao_parcelamento_id).children('option').length == 0)){
				s_id_aux = "#c_memo_index_opcao_parcelamento_" + i;
				indice = $(s_id_aux).val();
				if (indice != "") indice=converte_numero(indice);
				for (var j = 0; j <= indice; j++) {
					$(s_opcao_parcelamento_id).append(new Option('', ''));
				}
				if ($(s_opcao_parcelamento_id).children('option').length > 0) $(s_opcao_parcelamento_id).prop('selectedIndex', $(s_opcao_parcelamento_id).children('option').length-1);
				s_valor_id = "#c_cartao_valor_" + i;
				obtemOpcoesParcelamento(i, $(s_bandeira_id).val(), '<%=id_pedido_base%>', $(s_valor_id).val());
			}
		}

		$(".DataEntry").keydown(function (e) {
			if (e.which === 13) {
				e.preventDefault();
				var index = $(".DataEntry").index(this) + 1;
				try {
					$(".DataEntry").eq(index).focus();
				} catch (e) {
					// NOP
				}
			}
		});

		$(".CnpjCpf, .CardNumber, .SecurityCode, .CEP, .TelDDD, .TelNum").keydown(function (e) {
			// Allow: backspace, delete, tab, escape, enter and .
			if ($.inArray(e.keyCode, [46, 8, 9, 27, 13, 110, 190]) !== -1 ||
				// Allow: Ctrl+A, Command+A
				(e.keyCode == 65 && ( e.ctrlKey === true || e.metaKey === true ) ) || 
				// Allow: Ctrl+C, Ctrl+V, Ctrl+X
				(((e.keyCode == 67) || (e.keyCode == 86) ||(e.keyCode == 88)) && ( e.ctrlKey === true || e.metaKey === true ) ) || 
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

		$(".AlfaNum").keypress(function (e) {
			var key = String.fromCharCode(!e.charCode ? e.which : e.charCode);
			if ((key=="'") || (key=="\"") || (key=="|"))
			{
				e.preventDefault();
				return false;
			}
		});

	});

	//Dynamically assign height
	function sizeDivAjaxRunning() {
		var newTop = $(window).scrollTop() + "px";
		$("#divAjaxRunning").css("top", newTop);
	}

	function atualizaTotalDestaTransacao(){
		var vl_total = 0;
		var vl_restante;
		$(".TxtValor").each(function(){
			vl_total += converte_numero($(this).val());
		});
		$("#c_vl_total_desta_transacao").text(formata_moeda(vl_total));
		vl_restante=vl_saldo_a_pagar-vl_total;
		$("#c_vl_restante").text(formata_moeda(vl_restante));
		if (vl_restante == 0){
			$("#spnTitVlRestante").removeClass("corRed");
			$("#c_vl_restante").removeClass("corRed");
		}
		else {
			$("#spnTitVlRestante").addClass("corRed");
			$("#c_vl_restante").addClass("corRed");
		}
	}
</script>

<script language="JavaScript" type="text/javascript">
function obtemOpcoesParcelamento(_indice, _bandeira, _pedido, _valor_pagamento){
	var s_select_id, idx_option_selecionado;

	s_select_id = "#c_opcao_parcelamento_" + _indice;
	idx_option_selecionado = $(s_select_id + " option:selected").index();
	if (idx_option_selecionado < 0) idx_option_selecionado = 0;
	$(s_select_id).empty();
	$(s_select_id).append($('<option>', {
		value : "",
		text : "SELECIONE"
	}));

	if (_bandeira==null) return;
	if (_bandeira.toString().length==0) return;
	if (_valor_pagamento==null) return;
	if (_valor_pagamento.toString().length==0) return;
	if (converte_numero(_valor_pagamento.toString())==0) return;

	var jqxhr = $.ajax({
		url: "../Global/AjaxBraspagCSOpcoesParcelamento.asp",
		type: "POST",
		dataType: 'json',
		data: {
			bandeira : _bandeira,
			pedido : _pedido,
			valor_pagamento : _valor_pagamento
		}
	})
	.done(function(response){
		$(s_select_id).empty();
		$(s_select_id).append($('<option>', {
			value : "",
			text : "SELECIONE"
		}));
		for (var i = 0; i < response.length; i++) {
			$(s_select_id).append($('<option>', {
				value : response[i].value,
				text : response[i].description
			}));
		}
		// EVITA QUE A LISTA DO SELECT SEJA EXIBIDA SEM FORMATAÇÃO (BUG) QUANDO ESTA ROTINA É ACIONADA DURANTE A ABERTURA DA LISTA
		var hasFocus = $(s_select_id).is(':focus');
		$(s_select_id).hide();
		$(s_select_id).show();
		$(s_select_id).prop('selectedIndex', idx_option_selecionado);
		if (hasFocus) $(s_select_id).focus();
	})
	.fail(function(jqXHR, textStatus){
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
		
		alert("Falha ao tentar consultar as opções de parcelamento!!\n\n" + msgErro);
	});
}

function Navega(url) {
	window.location.href = url;
}

function fPEDConsulta() {
	fPED.action = "../ClienteCartao/PedidoConsulta.asp";
	window.status = "Aguarde ...";
	fPED.submit();
}

function fPAGTOConclui(f) {
	var s_id, s_id_aux, s_bandeira_id, s_valor_id, s_parcela, vl_pagto, v, n, s_erro, s_confirma;
	var vl_total_pagto=0;

	for (var i = 1; i <= qtde_cartoes; i++) {
		// BANDEIRA
		s_id = "#c_cartao_bandeira_" + i;
		if (trim($(s_id).val()).length == 0) {
			alert("É necessário selecionar a bandeira do cartão!!");
			$(s_id).focus();
			return;
		}

		// VALOR
		vl_pagto = 0;
		s_id = "#c_cartao_valor_" + i;
		if (($(s_id).val().length == 0)||(converte_numero($(s_id).val()) == 0)){
			alert("Informe o valor do pagamento!!");
			$(s_id).focus();
			return;
		}
		vl_pagto = converte_numero($(s_id).val());
		vl_total_pagto += vl_pagto;

		// PARCELAS
		s_id = "#c_opcao_parcelamento_" + i;
		if (trim($(s_id).val()).length == 0) {
			alert("É necessário selecionar uma opção de parcelamento!!");
			$(s_id).focus();
			return;
		}

		// NOME
		s_id = "#c_cartao_nome_" + i;
		if (trim($(s_id).val()).length <= 5) {
			alert("Nome inválido!!");
			$(s_id).focus();
			return;
		}
	
		if (trim($(s_id).val()).indexOf(" ") == -1) {
			alert("Informe o sobrenome!!");
			$(s_id).focus();
			return;
		}
		
		// CPF/CNPJ
		s_id = "#c_cartao_cpf_cnpj_" + i;
		if (retorna_so_digitos($(s_id).val()) == "") {
			alert("Informe o CPF/CNPJ do titular do cartão!!");
			$(s_id).focus();
			return;
		}
	
		if (!cnpj_cpf_ok($(s_id).val())) {
			alert("CPF/CNPJ inválido!!");
			$(s_id).focus();
			return;
		}
		
		// NÚMERO DO CARTÃO
		s_id = "#c_cartao_numero_" + i;
		if (retorna_so_digitos($(s_id).val()).length == 0) {
			alert("Informe o número do cartão!!");
			$(s_id).focus();
			return;
		}

		if (retorna_so_digitos($(s_id).val()).length < 14) {
			alert("Número do cartão com tamanho inválido!!");
			$(s_id).focus();
			return;
		}
	
		// MÊS DA VALIDADE
		s_id = "#c_cartao_validade_mes_" + i;
		if (trim($(s_id).val()) == "") {
			alert("Informe o mês da validade do cartão!!");
			$(s_id).focus();
			return;
		}
	
		// ANO DA VALIDADE
		s_id = "#c_cartao_validade_ano_" + i;
		if (trim($(s_id).val()) == "") {
			alert("Informe o ano da validade do cartão!!");
			$(s_id).focus();
			return;
		}
		
		// CÓDIGO SEGURANÇA
		s_id = "#c_cartao_codigo_seguranca_" + i;
		if (trim($(s_id).val()) == "") {
			alert("Informe o código de segurança do cartão!!");
			$(s_id).focus();
			return;
		}
	
		// CARTÃO PRÓPRIO
		s_id = "#c_cartao_proprio_" + i;
		if ($(s_id).val().length == 0) {
			alert("Informe se o cartão pertence ao comprador do pedido ou se é de terceiro!!");
			$(s_id).focus();
			return;
		}
	
		// ENDEREÇO - LOGRADOURO
		s_id = "#c_fatura_end_logradouro_" + i;
		if (trim($(s_id).val()) == "") {
			alert("Informe o endereço da fatura!!");
			$(s_id).focus();
			return;
		}
	
		// ENDEREÇO - NÚMERO
		s_id = "#c_fatura_end_numero_" + i;
		if (trim($(s_id).val()) == "") {
			alert("Informe o número do endereço da fatura!!");
			$(s_id).focus();
			return;
		}
	
		// BAIRRO
		s_id = "#c_fatura_end_bairro_" + i;
		if (trim($(s_id).val()) == "") {
			alert("Informe o bairro do endereço da fatura!!");
			$(s_id).focus();
			return;
		}
	
		// CIDADE
		s_id = "#c_fatura_end_cidade_" + i;
		if (trim($(s_id).val()) == "") {
			alert("Informe a cidade do endereço da fatura!!");
			$(s_id).focus();
			return;
		}
		
		// UF
		s_id = "#c_fatura_end_uf_" + i;
		if (trim($(s_id).val()) == "") {
			alert("Informe a UF do endereço da fatura!!");
			$(s_id).focus();
			return;
		}
		
		// CEP
		s_id = "#c_fatura_end_cep_" + i;
		if (retorna_so_digitos($(s_id).val()) == "") {
			alert("Informe o CEP do endereço da fatura!!");
			$(s_id).focus();
			return;
		}
	
		if (retorna_so_digitos($(s_id).val()).length != 8) {
			alert("CEP inválido!!");
			$(s_id).focus();
			return;
		}
		
		// DDD
		s_id = "#c_fatura_telefone_ddd_" + i;
		if (retorna_so_digitos($(s_id).val()).length == 0) {
			alert("Informe o DDD!!");
			$(s_id).focus();
			return;
		}
	
		if (retorna_so_digitos($(s_id).val()).length != 2) {
			alert("DDD inválido!!");
			$(s_id).focus();
			return;
		}
		
		// TELEFONE
		s_id = "#c_fatura_telefone_numero_" + i;
		if (retorna_so_digitos($(s_id).val()).length == 0) {
			alert("Informe o telefone!!");
			$(s_id).focus();
			return;
		}
	
		if ((retorna_so_digitos($(s_id).val()).length < 7) || (retorna_so_digitos($(s_id).val()).length > 9)) {
			alert("Telefone inválido!!");
			$(s_id).focus();
			return;
		}

		// CONSISTE VALOR DA PARCELA MÍNIMA
		s_bandeira_id = "#c_cartao_bandeira_" + i;
		s_id = "#c_opcao_parcelamento_" + i;
		s_parcela = trim($(s_id).val());
		if (s_parcela.length != 0) {
			if (s_parcela.indexOf("|") != -1){
				v = s_parcela.split("|");
				n = converte_numero(v[v.length-1]);
			}
			else{
				n = converte_numero(s_parcela);
				// À Vista (no crédito)?
				if (n==0) n=1;
			}
			if (vBandeiraParcMin[$(s_bandeira_id).val().toUpperCase()] == -1) {
				alert("O valor da parcela mínima não está cadastrado para esta bandeira!!");
				return;
			}
			if ((vl_pagto/n) < vBandeiraParcMin[$(s_bandeira_id).val().toUpperCase()]) {
				alert("O valor está abaixo do valor da parcela mínima aceita!!");
				$(s_id).focus();
				return;
			}
		}
		
	}
	
	// MEMORIZAÇÃO (TRATAMENTO P/ HISTORY.BACK)
	for (var i = 1; i <= qtde_cartoes; i++) {
		s_id = "#c_opcao_parcelamento_" + i;
		s_id_aux = "#c_memo_index_opcao_parcelamento_" + i;
		$(s_id_aux).val($(s_id + " option:selected").index());
	}

	// CONSISTÊNCIA (BANDEIRA E VALOR IGUAIS)
	s_erro = "";
	for (var i = 1; i <= qtde_cartoes; i++) {
		for (var j = 1; j <= (i-1); j++) {
			s_id = "#c_cartao_bandeira_" + i;
			s_id_aux = "#c_cartao_bandeira_" + j;
			if (trim($(s_id).val()) != trim($(s_id_aux).val())) break;

			s_id = "#c_cartao_valor_" + i;
			s_id_aux = "#c_cartao_valor_" + j;
			if (converte_numero(trim($(s_id).val())) != converte_numero(trim($(s_id_aux).val()))) break;
			
			if (s_erro != "") s_erro += "\n";
			s_bandeira_id = "#c_cartao_bandeira_" + i;
			s_valor_id = "#c_cartao_valor_" + i;
			s_erro += "Os cartões " + j + " e " + i + " são da mesma bandeira (" + vBandeiraDescricao[$(s_bandeira_id).val().toUpperCase()] + ") e estão sendo usados para pagar um valor idêntico (" + SIMBOLO_MONETARIO + " " + $(s_valor_id).val() + ")";
		}
	}
	
	if (s_erro != "") 
	{
		s_erro += "\n\n";
		s_erro += "Devido a uma limitação do gateway de pagamentos, por favor, escolha um valor de pagamento diferente para cada cartão de mesma bandeira!";
		alert(s_erro);
		return;
	}

	s_confirma = "";
	if (vl_total_pagto < (vl_saldo_a_pagar - <%=js_formata_numero(MAX_VALOR_MARGEM_ERRO_PAGAMENTO)%>)) {
		s_confirma=(qtde_cartoes > 1?"O valor total deste pagamento é insuficiente para quitar o valor a pagar!":"O valor deste pagamento é insuficiente para quitar o valor a pagar!")
	}

	if (vl_total_pagto > (vl_saldo_a_pagar + <%=js_formata_numero(MAX_VALOR_MARGEM_ERRO_PAGAMENTO)%>)) {
		s_confirma=(qtde_cartoes > 1?"O valor total deste pagamento excede o valor a pagar!":"O valor deste pagamento excede o valor a pagar!")
	}

	if (s_confirma!=""){
		s_confirma += "\n\nDeseja prosseguir com o pagamento assim mesmo?";
		$("#msgConfirm").html(s_confirma.replace("\n","<br />"));
		$("#msgConfirm").css('display', 'block');
		$("#dialogConfirm").dialog({
			resizable: false,
			height: 200,
			width: 600,
			scroll: false,
			modal: true,
			buttons: {
				"Sim": function () {
					$(this).dialog("close");
					fPAGTOConcluiExecutaSubmit(f);
				},
				"Não": function () {
					$(this).dialog("close");
					$("#msgConfirm").css('display', 'none');
					return;
				}
			}
		});
		return;
	}

	fPAGTOConcluiExecutaSubmit(f);
}

function fPAGTOConcluiExecutaSubmit(f){
	memorizaDados();
	f.action = "P060PagtoReady.asp";
	dPROXIMO.style.visibility="hidden";
	window.status = "Aguarde ...";
	f.submit();
}

function memorizaDados(){
	$("#c_memo_form").val("");
	var m = JSON.stringify($("#fPAGTO").serializeArray());
	$("#c_memo_form").val(m);
}

function restauraDadosMemorizados(){
	var s_id, s_value;
	var m=$("#c_memo_form").val();
	if (m==null) return;
	if (m=="") return;
	var r=JSON.parse(m);
	for (var i = 0; i < r.length; i++) {
		try{
			s_id="#"+r[i].name;
			s_value=r[i].value;
			$(s_id).val(s_value);
		}
		catch(e){
			// NOP
		}
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

<link href="<%=URL_FILE__E_CSS%>" rel="stylesheet" type="text/css">
<link href="<%=URL_FILE__EPRINTER_CSS%>" rel="stylesheet" type="text/css" media="print">
<link href="<%=URL_FILE__EGWBP_CSS%>" rel="stylesheet" type="text/css">
<link href="<%=URL_FILE__JQUERY_UI_CSS%>" rel="stylesheet" type="text/css">
<link href="<%=URL_FILE__E_LOGO_TOP_BS_CSS%>" Rel="stylesheet" Type="text/css">

<style type="text/css">
body::before
{
	content: '';
	border: none;
	margin-top: 0px;
	margin-bottom: 0px;
	padding: 0px;
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
.trHidden
{
	display:none;
}
.tdColMargin
{
	width:15px;
}
/* Ajuda */
.modal {
	display:none; 
	position:absolute; 
	width:300px; 
	background-color:#FFFFFF;
	top:240px; 
	left:700px; 
	padding:20px; 
	border:#CCC 1px solid;
}
::-webkit-input-placeholder { /* WebKit, Blink, Edge */
    color: #999;
	text-align:left;
}
:-moz-placeholder { /* Mozilla Firefox 4 to 18 */
   color: #999;
   opacity:  1;
   text-align:left;
}
::-moz-placeholder { /* Mozilla Firefox 19+ */
   color: #999;
   opacity: 1;
   text-align:left;
}
:-ms-input-placeholder { /* Internet Explorer 10-11 */
   color: #999;
   text-align:left;
}
.spnAviso {
	color: #f37b20;
	font-size:10pt;
}
.tdSpnAviso{
	border: 1px solid #f37b20;
	padding:5px;
}
.corRed {
	color:red;
}
</style>

<% if alerta <> "" then %>
<!-- ************************************************************ -->
<!-- **********  PÁGINA PARA EXIBIR MENSAGENS DE ERRO  ********** -->
<!-- ************************************************************ -->
<body onload="bVOLTAR.focus()">
<center>
<br>
<!--  T E L A  -->
<p class="T">A V I S O</p>
<div class="MtAlerta" style="width:600px;font-weight:bold;" align="center"><p style='margin:5px 2px 5px 2px;'><%=alerta%></p></div>
<br><br>
<p class="TracoBottom"></p>
<table cellspacing="0">
<tr>
	<td align="center"><a name="bVOLTAR" href="javascript:history.back()"><img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
</tr>
</table>
</center>
</body>




<% else %>
<!-- ****************************************************************** -->
<!-- **********  PÁGINA PARA EXIBIR RESUMO DO PEDIDO  ***************** -->
<!-- ****************************************************************** -->
<body>
<center>

<table class="notPrint" id="tbl_logotipo_bonshop" width="100%" cellpadding="0" cellspacing="0" border="0">
	<tr>
		<td align="center"><img alt="<%=SITE_CLIENTE_HEADER__ALT_IMG_TEXT%>" src="../imagem/<%=SITE_CLIENTE_HEADER__LOGOTIPO%>" /></td>
	</tr>
</table>
<table class="notPrint" id="pagina_tbl_cabecalho" cellspacing="0px" cellpadding="0px">
	<tbody>
		<tr style="height:78px;">
			<td id="topo_verde" colspan="3">
				<div id="moldura_do_letreiro">
					<div id="letreiro_div" style="display:block;"></div>
				</div>
				<div id="telefone"></div>
			</td>
		</tr>
		<tr>
			<td id="topo_azul" colspan="3">&nbsp;</td>
		</tr>
	</tbody>
</table>

<form id="fPED" name="fPED" METHOD="POST">
<input type="hidden" name="pedido_selecionado" id="pedido_selecionado" value='<%=pedido_selecionado%>'>
<input type="hidden" name="cnpj_cpf_selecionado" id="cnpj_cpf_selecionado" value='<%=cnpj_cpf_selecionado%>'>
</form>

<form id="fPAGTO" name="fPAGTO" method="post" >
<input type="hidden" name="c_owner" value="<%=owner%>" />
<input type="hidden" name="pedido_selecionado" value="<%=pedido_selecionado%>">
<input type="hidden" name="cnpj_cpf_selecionado" id="cnpj_cpf_selecionado" value='<%=cnpj_cpf_selecionado%>'>
<input type="hidden" name="c_fatura_telefone_pais" value="55" />
<input type="hidden" name="pedido_com_sufixo_nsu" value="<%=pedido_com_sufixo_nsu%>" />
<input type="hidden" name="FingerPrint_SessionID" value="<%=FingerPrint_SessionID%>" />
<input type="hidden" name="c_vl_a_pagar" value="<%=formata_moeda(vl_saldo_a_pagar)%>" />
<input type="hidden" name="c_qtde_cartoes" value="<%=c_qtde_cartoes%>" />
<!-- MEMORIZAÇÃO (TRATAMENTO P/ HISTORY.BACK) -->
<% for i = 1 to qtde_cartoes %>
<input type="hidden" id="c_memo_index_opcao_parcelamento_<%=i%>" name="c_memo_index_opcao_parcelamento_<%=i%>" value="" />
<input type="hidden" id="c_memo_form" name="c_memo_form" />
<% next %>


<!--  EXIBE RESUMO DO PAGAMENTO  -->
<br />
<br />
<table class="Qx" cellspacing="0" cellpadding="0" border="0">
	<tr>
		<td class="tdSpnAviso" colspan="5" align="left">
			<span class="spnAviso">Importante:</span><br /><span class="spnAviso">O limite disponível no cartão de crédito deve ser superior ao valor total do pagamento e não ao valor de cada parcela.</span>
		</td>
	</tr>
	<tr>
		<td colspan="5" style="height:20px;" align="left"></td>
	</tr>
	<% if qtde_cartoes > 1 then %>
	<tr><td colspan="5" style="height:20px;" align="left"></td></tr>
	<% end if %>

	<% for i_cartao = 1 to qtde_cartoes %>
	<% if qtde_cartoes > 1 then %>
	<tr>
		<td colspan="5" align="center">
			<span class="PLTd" style="font-size:14pt;"><%=i_cartao%>º CARTÃO</span>
		</td>
	</tr>
	<% end if %>
	<tr>
		<td colspan="5" class="MC ME MD" style="height:6px;" align="left"></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td class="tdColMargin ME">&nbsp;</td>
		<!-- BANDEIRA -->
		<td align="right"><span class="PLTd" style="vertical-align:middle;">BANDEIRA DO CARTÃO</span></td>
		<td align="left" nowrap>
			<select id="c_cartao_bandeira_<%=i_cartao%>" name="c_cartao_bandeira_<%=i_cartao%>" class="CARDSel SelBand DataEntry" style="width:150px;">
				<%=BraspagCS_monta_select_bandeiras(owner, Null)%>
			</select>
		</td>
	<!-- LOGO DA BANDEIRA -->
		<td rowspan="5" align="right" valign="middle">
			<img id="c_cartao_bandeira_logo_<%=i_cartao%>" name="c_cartao_bandeira_logo_<%=i_cartao%>" src="" border="0" />
		</td>
		<td class="tdColMargin MD">&nbsp;</td>
	</tr>
	<tr>
		<td colspan="5" class="ME MD" style="height:6px;" align="left"></td>
	</tr>
	<!-- VALOR -->
	<tr bgcolor="#FFFFFF">
		<td class="tdColMargin ME">&nbsp;</td>
		<td align="right"><span class="PLTd" style="vertical-align:middle;">VALOR <%=SIMBOLO_MONETARIO%></span></td>
		<td align="left" nowrap>
			<input name="c_cartao_valor_<%=i_cartao%>" id="c_cartao_valor_<%=i_cartao%>" class="CCARDe TxtValor DataEntry" style="width:150px;" value="" maxlength="18" onblur="this.value=formata_moeda(trim(this.value));" placeholder="Valor neste cartão" />
		</td>
		<td class="tdColMargin MD">&nbsp;</td>
	</tr>
	<tr><td colspan="5" class="ME MD" style="height:6px;" align="left"></td></tr>
	<!-- PARCELAS -->
	<tr bgcolor="#FFFFFF">
		<td class="tdColMargin ME">&nbsp;</td>
		<td align="right"><span class="PLTd" style="vertical-align:middle;">Nº PARCELAS</span></td>
		<td align="left" nowrap>
			<select id="c_opcao_parcelamento_<%=i_cartao%>" name="c_opcao_parcelamento_<%=i_cartao%>" class="CARDSel SelParc DataEntry" style="width:300px;">
			</select>
		</td>
		<td class="tdColMargin MD">&nbsp;</td>
	</tr>
	<!-- SEPARAÇÃO DE BLOCOS -->
	<tr><td colspan="5" class="ME MD" style="height:20px;" align="left"></td></tr>
	<tr><td colspan="5" class="ME MD" style="height:6px;" align="left"></td></tr>
<!-- DADOS DO CARTÃO -->
	<tr bgcolor="#FFFFFF">
		<td class="tdColMargin ME">&nbsp;</td>
		<td align="right"><span class="PLTd">NOME DO TITULAR</span></td>
		<td align="left" nowrap><input name="c_cartao_nome_<%=i_cartao%>" id="c_cartao_nome_<%=i_cartao%>" class="CCARDe AlfaNum DataEntry" style="width:300px;" value="" maxlength="80" onblur="this.value=trim(this.value);" placeholder="Exatamente como impresso no cartão" /></td>
	<!-- CPF/CNPJ -->
		<td align="right" valign="middle">
			<span class="PLTd" style="vertical-align:middle;">CPF/CNPJ</span>
			<input name="c_cartao_cpf_cnpj_<%=i_cartao%>" id="c_cartao_cpf_cnpj_<%=i_cartao%>" class="CCARDc CnpjCpf DataEntry" style="width:170px;" value="" maxlength="18" onblur="if (!cnpj_cpf_ok(this.value)) {alert('CPF/CNPJ inválido!!');} else {this.value=cnpj_cpf_formata(trim(this.value));}" placeholder="CPF/CNPJ do titular" />
		</td>
		<td class="tdColMargin MD">&nbsp;</td>
	</tr>
	<tr><td colspan="5" class="ME MD" style="height:6px;" align="left"></td></tr>
	<tr bgcolor="#FFFFFF">
		<td class="tdColMargin ME">&nbsp;</td>
		<td align="right"><span class="PLTd">Nº DO CARTÃO</span></td>
		<td align="left"><input name="c_cartao_numero_<%=i_cartao%>" id="c_cartao_numero_<%=i_cartao%>" class="CCARDe CardNumber DataEntry" style="width:200px;" value="" maxlength="19" onblur="this.value=trim(this.value);" /></td>
		<td align="right">
			<table border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td align="right">
						<span class="PLTd" style="vertical-align:middle;">VALIDADE</span>
						<select id="c_cartao_validade_mes_<%=i_cartao%>" name="c_cartao_validade_mes_<%=i_cartao%>" class="CARDSel DataEntry" style="margin-top:4pt; margin-bottom:4pt;">
							<%=braspag_cartao_validade_mes_monta_itens_select(Null)%>
						</select>
					</td>
					<td style="width:10px;">&nbsp</td>
					<td align="right">
						<select id="c_cartao_validade_ano_<%=i_cartao%>" name="c_cartao_validade_ano_<%=i_cartao%>" class="CARDSel DataEntry" style="margin-top:4pt; margin-bottom:4pt;margin-right:0px;">
							<%=braspag_cartao_validade_ano_monta_itens_select(Null)%>
						</select>
					</td>
				</tr>
			</table>
		</td>
		<td class="tdColMargin MD">&nbsp;</td>
	</tr>
	<tr><td colspan="5" class="ME MD" style="height:6px;" align="left"></td></tr>
	<tr bgcolor="#FFFFFF">
		<td class="tdColMargin ME">&nbsp;</td>
		<td align="right"><span class="PLTd">CÓDIGO DE SEGURANÇA</span></td>
		<td align="left">
			<table cellpadding="0" cellspacing="0" border="0">
				<tr>
					<td align="left" style="width:70px;">
						<input name="c_cartao_codigo_seguranca_<%=i_cartao%>" id="c_cartao_codigo_seguranca_<%=i_cartao%>" class="CCARDe SecurityCode DataEntry" style="width:60px;" value="" maxlength="4" onblur="this.value=trim(this.value);" />
					</td>
					<td align="left" style="width:30px;">
						<img src="../Imagem/Braspag/info.png" title="" id="imgCodSeguranca_<%=i_cartao%>" name="imgCodSeguranca_<%=i_cartao%>" class="CodSeguranca" />
					</td>
				</tr>
			</table>
		</td>
	<!-- CARTÃO PRÓPRIO? -->
		<td align="right" valign="middle">
			<span class="PLTd" style="vertical-align:middle;">CARTÃO PERTENCE AO COMPRADOR</span>
			<select name="c_cartao_proprio_<%=i_cartao%>" id="c_cartao_proprio_<%=i_cartao%>" class="DataEntry" style="margin-right:0px;">
				<option value="" selected>&nbsp;</option>
				<option value="PROPRIO">SIM</option>
				<option value="TERCEIRO">NÃO</option>
			</select>
		</td>
		<td class="tdColMargin MD">&nbsp;</td>
	</tr>
	<!-- SEPARAÇÃO DE BLOCOS -->
	<tr><td colspan="5" class="ME MD" style="height:20px;" align="left"></td></tr>
	<tr><td colspan="5" class="ME MD" style="height:6px;" align="left"></td></tr>
	<!-- DADOS CADASTRAIS -->
	<tr bgcolor="#FFFFFF">
		<td class="tdColMargin ME">&nbsp;</td>
		<td align="right"><span class="PLTd">ENDEREÇO DA FATURA</span></td>
		<td align="left"><input name="c_fatura_end_logradouro_<%=i_cartao%>" id="c_fatura_end_logradouro_<%=i_cartao%>" class="CCARDe AlfaNum DataEntry" style="width:300px;" value="" maxlength="55" onblur="this.value=trim(this.value);" placeholder="Rua, avenida, etc" /></td>
		<td align="right"><span class="PLTd" style="vertical-align:middle;margin-right:6px;">NÚMERO</span><input name="c_fatura_end_numero_<%=i_cartao%>" id="c_fatura_end_numero_<%=i_cartao%>" class="CCARDe AlfaNum DataEntry" style="width:100px;" value="" maxlength="10" onblur="this.value=trim(this.value);" /></td>
		<td class="tdColMargin MD">&nbsp;</td>
	</tr>
	<tr><td colspan="5" class="ME MD" style="height:6px;" align="left"></td></tr>
	<tr bgcolor="#FFFFFF">
		<td class="tdColMargin ME">&nbsp;</td>
		<td align="right"><span class="PLTd">COMPLEMENTO</span></td>
		<td align="left"><input name="c_fatura_end_complemento_<%=i_cartao%>" id="c_fatura_end_complemento_<%=i_cartao%>" class="CCARDe AlfaNum DataEntry" style="width:150px;margin-right:0px;" value="" maxlength="60" onblur="this.value=trim(this.value);" /></td>
		<td align="right"><span class="PLTd" style="vertical-align:middle;margin-right:6px;">BAIRRO</span><input name="c_fatura_end_bairro_<%=i_cartao%>" id="c_fatura_end_bairro_<%=i_cartao%>" class="CCARDe AlfaNum DataEntry" style="width:200px;margin-right:0px;" value="" maxlength="150" onblur="this.value=trim(this.value);" /></td>
		<td class="tdColMargin MD">&nbsp;</td>
	</tr>
	<tr><td colspan="5" class="ME MD" style="height:6px;" align="left"></td></tr>
	<tr bgcolor="#FFFFFF">
		<td class="tdColMargin ME">&nbsp;</td>
		<td align="right"><span class="PLTd">CIDADE</span></td>
		<td colspan="2" align="left">
			<table border="0" cellspacing="0" cellpadding="0" width="100%">
			<tr>
				<td align="left">
				<input name="c_fatura_end_cidade_<%=i_cartao%>" id="c_fatura_end_cidade_<%=i_cartao%>" class="CCARDe AlfaNum DataEntry" style="width:300px;" value="" maxlength="60" onblur="this.value=trim(this.value);" />
				</td>
				<td align="right">
				<span class="PLTd" style="vertical-align:middle;margin-right:6px;">UF</span>
				<select id="c_fatura_end_uf_<%=i_cartao%>" name="c_fatura_end_uf_<%=i_cartao%>" class="CARDSel DataEntry" style="margin-right:0px;">
				<% =UF_monta_itens_select(s) %>
				</select>
				</td>
			</tr>
			</table>
		</td>
		<td class="tdColMargin MD">&nbsp;</td>
	</tr>
	<tr><td colspan="5" class="ME MD" style="height:6px;" align="left"></td></tr>
	<tr bgcolor="#FFFFFF">
		<td class="tdColMargin ME">&nbsp;</td>
		<td align="right"><span class="PLTd">CEP</span></td>
		<td align="left"><input name="c_fatura_end_cep_<%=i_cartao%>" id="c_fatura_end_cep_<%=i_cartao%>" class="CCARDc CEP DataEntry" style="width:120px;" value="" maxlength="9" onblur="this.value=cep_formata(trim(this.value));" /></td>
		<td align="right">
			<span class="PLTd" style="vertical-align:middle;margin-right:6px;">TELEFONE</span>
			<input name="c_fatura_telefone_ddd_<%=i_cartao%>" id="c_fatura_telefone_ddd_<%=i_cartao%>" class="CCARDc TelDDD DataEntry" style="width:40px;" value="" maxlength="2" onblur="this.value=trim(this.value);" />
			&nbsp;<input name="c_fatura_telefone_numero_<%=i_cartao%>" id="c_fatura_telefone_numero_<%=i_cartao%>" class="CCARDc TelNum DataEntry" style="width:120px;margin-right:0px;" value="" maxlength="10" 
				onblur="this.value=telefone_formata(this.value);" />
		</td>
		<td class="tdColMargin MD">&nbsp;</td>
	</tr>
	<tr><td colspan="5" class="MB ME MD" style="height:6px;" align="left"></td></tr>
	<% if i_cartao < qtde_cartoes then %>
	<tr><td colspan="5" style="height:30px;" align="left"></td></tr>
	<% end if %>
	<% next %>
</table>
<br />

<!-- ************   SEPARADOR   ************ -->
<table width="735" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
<td class="Rc" align="left">&nbsp;</td>
</tr>
</table>

<!--  Resumo dos valores  -->
<br />
<table class="Qx" cellspacing="0" cellpadding="0" border="0">
	<tr>
		<td align="right"><span class="PLTd" style="font-size:12pt;">Valor a pagar:&nbsp;<%=SIMBOLO_MONETARIO%></span></td>
		<td align="right" style="width:150px;"><span class="PLTd" style="font-size:12pt;"><%=formata_moeda(vl_saldo_a_pagar)%></span></td>
	</tr>
	<tr>
		<% if qtde_cartoes > 1 then s = "Valor total destes pagamentos:" else s = "Valor deste pagamento:"%>
		<td align="right"><span class="PLTd" style="font-size:12pt;"><%=s%>&nbsp;<%=SIMBOLO_MONETARIO%></span></td>
		<td align="right" style="width:150px;"><span class="PLTd" style="font-size:12pt;" id="c_vl_total_desta_transacao"></span></td>
	</tr>
	<tr class="<%if qtde_cartoes = 1 then Response.Write "trHidden"%>">
		<td align="right"><span class="PLTd corRed" style="font-size:12pt;" id="spnTitVlRestante">Valor restante:&nbsp;<%=SIMBOLO_MONETARIO%></span></td>
		<td align="right" style="width:150px;"><span class="PLTd corRed" style="font-size:12pt;" id="c_vl_restante"></span></td>
	</tr>
</table>

<div id="divCodSegInfo" class="modal">
	O Código de Segurança possui 3 dígitos e está localizado no verso do cartão. 
	<br /><br />
	<img src="../Imagem/Braspag/cod_seguranca.gif" alt="Código de segurança" width="183" height="100" />
</div>


<!-- ************   SEPARADOR   ************ -->
<table width="735" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
<td class="Rc" align="left">&nbsp;</td>
</tr>
</table>


<!-- ************   BOTÕES   ************ -->
<table class="notPrint" width="735" cellpadding="0" cellspacing="0">
<tr>
	<td align="left"><a name="bVOLTAR" href="javascript:history.back();" title="volta para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a>
	</td>
	<td align="right">
		<div name="dPROXIMO" id="dPROXIMO"><a name="bPROXIMO" href="javascript:fPAGTOConclui(fPAGTO)" title="vai para a página seguinte">
			<img src="../botao/proximo.gif" width="176" height="55" border="0"></a></div>
	</td>
</tr>
</table>

</form>


</center>

<%' QUANDO O ENVIO É FEITO NO AMBIENTE 'CALL CENTER', OS SCRIPTS DO ANTI-FRAUDE NÃO DEVEM SER EXECUTADOS %>


<div id="divAjaxRunning"><img src="../Imagem/ajax_loader_gray_256.gif" class="AjaxImgLoader"/></div>

<div id="dialogConfirm" title="Prosseguir?">
	<span id="msgConfirm" style="display:none"></span>
</div>

<!-- DF -->
<script type = "text/javascript">
	var _csdp = _csdp || [];
	_csdp.push(['Key', '<%=CLEARSALE_DF_KEY%>']);
	_csdp.push(['App', '<%=CLEARSALE_DF_APP%>']);
	_csdp.push(['SessionID', '<%=FingerPrint_SessionID%>']);

	(function () {
		var csd = document.createElement('script');
		csd.type = 'text/javascript';
		csd.async = true;
		csd.src = 'https://device.clearsale.com.br/profiler/fp.js';
		var sc = document.getElementsByTagName('script')[0];
		sc.parentNode.insertBefore(csd, sc);
	})();
</script>


<% if SITE_CLIENTE_EXIBIR_LOGO_SSL then %>
<script language="JavaScript" type="text/javascript">
	logo_ssl_corner("../imagem/ssl/ssl_corner.gif");
</script>
<% end if %>

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