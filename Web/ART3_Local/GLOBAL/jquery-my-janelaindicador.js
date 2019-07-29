// the semi-colon before the function invocation is a safety 
// net against concatenated scripts and/or other plugins 
// that are not closed properly.
; (function($) {
	/* Bloco 1 - Funções copiadas (e adaptadas) da página que fazia busca do CEP */
	/* início bloco 01 */

	var objAjaxPesqCep;
	var objAjaxPesqLocalidades;
	var OPCAO_PESQUISA_POR_CEP = "POR_CEP";
	var OPCAO_PESQUISA_POR_ENDERECO = "POR_END";
	var COL_CHECK = 0;
	var COL_CEP = 1;
	var COL_UF = 2;
	var COL_LOCALIDADE = 3;
	var COL_BAIRRO = 4;
	var COL_LOGRADOURO = 5;
	var COL_LOGRADOURO_COMPLEMENTO = 6;
	var strCep, strUF, strLocalidade, strBairro, strLogradouro, strEnderecoNumero, strEnderecoComplemento;
	var windowScrollTopAnterior;

	function ConfirmarOperacao() {
		var i, oRow, idxSelecionado;
		idxSelecionado = -1;
		for (i = 0; i < rb_check.length; i++) {
			if (rb_check[i].checked) {
				idxSelecionado = i;
				break;
			}
		}
		if (idxSelecionado == -1) {
			alert("Nenhum CEP foi selecionado!!");
			return false;
		}

		if (trim(c_endereco_numero.value) == "") {
			alert("Informe o número do endereço!!");
			c_endereco_numero.focus();
			return false;
		}

		//	LEMBRE-SE: O ARRAY DE CAMPOS 'RB_CHECK' TEM O 1º CAMPO GERADO POR UM 'INPUT HIDDEN'
		//  =========  C/ A FUNÇÃO DE SEMPRE GERAR UM ARRAY, MESMO NO CASO DA TABELA TER APENAS 1 LINHA.
		//             PORTANTO, SEMPRE HAVERÁ 1 RB_CHECK A MAIS QUE O TOTAL DE LINHAS DE RESPOSTA E A 1ª
		//             LINHA CORRESPONDE AO RB_CHECK[1] E NÃO AO RB_CHECK[0]
		idxSelecionado--;
		oRow = oTBodyDados.rows[idxSelecionado];
		strCep = trim(oRow.cells[COL_CEP].innerHTML);
		if (strCep == "&nbsp;") strCep = "";
		strUF = trim(oRow.cells[COL_UF].innerHTML);
		if (strUF == "&nbsp;") strUF = "";
		strLocalidade = trim(oRow.cells[COL_LOCALIDADE].innerHTML);
		if (strLocalidade == "&nbsp;") strLocalidade = "";
		strBairro = trim(oRow.cells[COL_BAIRRO].innerHTML);
		if (strBairro == "&nbsp;") strBairro = "";
		strLogradouro = trim(oRow.cells[COL_LOGRADOURO].innerHTML);
		if (strLogradouro == "&nbsp;") strLogradouro = "";
		strEnderecoNumero = trim(c_endereco_numero.value);
		strEnderecoComplemento = trim(c_endereco_complemento.value);

		return true;
	}

	function IniciaPainel() {
		if (trim(c_cep_pesq.value) != "") {
			ExecutaPesquisaCEP(OPCAO_PESQUISA_POR_CEP);
		}
		c_cep_pesq.select();
		c_cep_pesq.focus();
	}

	function LimpaListaLocalidades() {
		var i, oOption;
		for (i = c_localidade_pesq.length - 1; i >= 0; i--) {
			c_localidade_pesq.remove(i);
		}

		//  Cria um item vazio
		oOption = document.createElement("OPTION");
		c_localidade_pesq.options.add(oOption);
		oOption.innerText = "";
		oOption.value = "";
	}

	function LimpaTabelaResultado() {
		var i;
		for (i = oTBodyDados.rows.length - 1; i >= 0; i--) {
			oTBodyDados.deleteRow(i);
		}
	}

	function TrataRespostaAjaxPesquisaLocalidades() {
		var i, strAux, strResp, xmlDoc, oOption, oNodes;
		if (objAjaxPesqLocalidades.readyState == AJAX_REQUEST_IS_COMPLETE) {
			strResp = objAjaxPesqLocalidades.responseText;
			if (strResp == "") {
				window.status = "Concluído";
				$("#divAjaxRunning").hide();
				alert("Nenhuma localidade encontrada!!");
				return;
			}

			if (strResp != "") {
				try {
					xmlDoc = objAjaxPesqLocalidades.responseXML.documentElement;
					for (i = 0; i < xmlDoc.getElementsByTagName("registro").length; i++) {
						oOption = document.createElement("OPTION");
						c_localidade_pesq.options.add(oOption);

						oNodes = xmlDoc.getElementsByTagName("localidade")[i];
						if (oNodes.childNodes.length > 0) strAux = oNodes.childNodes[0].nodeValue; else strAux = "";
						if (strAux == null) strAux = "";
						oOption.innerText = strAux;
						oOption.value = strAux;
					}
				}
				catch (e) {
					alert("Falha na consulta!!");
				}
			}
			window.status = "Concluído";
			$("#divAjaxRunning").hide();
			c_localidade_pesq.focus();
		}
	}

	$.CepUfCarregaLocalidades = function() {
		var strUrl, strUF;
		objAjaxPesqLocalidades = GetXmlHttpObject();
		if (objAjaxPesqLocalidades == null) {
			alert("O browser NÃO possui suporte ao AJAX!!");
			return;
		}

		//  Limpa lista de localidades
		LimpaListaLocalidades();
		LimpaTabelaResultado();

		strUF = trim(c_uf_pesq.value);
		if (strUF == "") {
			return;
		}

		window.status = "Aguarde, pesquisando as localidades de " + c_uf_pesq.value + " ...";
		$("#divAjaxRunning").fadeIn(100);

		strUrl = "../GLOBAL/AjaxCepLocalidadesPesqBD.asp";
		strUrl = strUrl + "?uf=" + c_uf_pesq.value;
		//  Prevents server from using a cached file
		strUrl = strUrl + "&sid=" + Math.random() + Math.random();
		objAjaxPesqLocalidades.onreadystatechange = TrataRespostaAjaxPesquisaLocalidades;
		objAjaxPesqLocalidades.open("GET", strUrl, true);
		objAjaxPesqLocalidades.send(null);
	}

	function TrataRespostaAjaxPesquisaCEP() {
		var i, intQtdeLinhas, strAux, strResp, xmlDoc, oRow, oCell, oNodes;
		if (objAjaxPesqCep.readyState == AJAX_REQUEST_IS_COMPLETE) {
			strResp = objAjaxPesqCep.responseText;
			if (strResp == "") {
				oRow = document.createElement("TR");
				oRow.style.backgroundColor = "whitesmoke";
				oTBodyDados.appendChild(oRow);
				oCell = document.createElement("TD");
				strAux = "<span class='N' style='font-size:14pt;font-weight:bold;color:red;'>Nenhum CEP encontrado</span>";
				oCell.colSpan = 7;
				oCell.align = "center";
				oCell.innerHTML = strAux;
				oRow.appendChild(oCell);
				window.status = "Concluído";
				$("#divAjaxRunning").hide();
				return;
			}

			intQtdeLinhas = 0;
			if (strResp != "") {
				try {
					xmlDoc = objAjaxPesqCep.responseXML.documentElement;
					for (i = 0; i < xmlDoc.getElementsByTagName("registro").length; i++) {
						intQtdeLinhas++;
						oRow = document.createElement("TR");
						oTBodyDados.appendChild(oRow);

						oCell = document.createElement("TD");
						strAux = "<input type='RADIO' id='rb_check' name='rb_check'>";
						oCell.align = "center";
						oCell.vAlign = "top";
						oCell.innerHTML = strAux;
						oRow.appendChild(oCell);

						oCell = document.createElement("TD");
						oNodes = xmlDoc.getElementsByTagName("cep")[i];
						if (oNodes.childNodes.length > 0) strAux = oNodes.childNodes[0].nodeValue; else strAux = "";
						if (strAux == null) strAux = "";
						if (strAux == "") strAux = "&nbsp;";
						oCell.noWrap = true;
						oCell.vAlign = "top";
						oCell.innerHTML = strAux;
						oRow.appendChild(oCell);

						oCell = document.createElement("TD");
						oNodes = xmlDoc.getElementsByTagName("uf")[i];
						if (oNodes.childNodes.length > 0) strAux = oNodes.childNodes[0].nodeValue; else strAux = "";
						if (strAux == null) strAux = "";
						if (strAux == "") strAux = "&nbsp;";
						oCell.noWrap = true;
						oCell.align = "center";
						oCell.vAlign = "top";
						oCell.innerHTML = strAux;
						oRow.appendChild(oCell);

						oCell = document.createElement("TD");
						oNodes = xmlDoc.getElementsByTagName("localidade")[i];
						if (oNodes.childNodes.length > 0) strAux = oNodes.childNodes[0].nodeValue; else strAux = "";
						if (strAux == null) strAux = "";
						if (strAux == "") strAux = "&nbsp;";
						oCell.vAlign = "top";
						oCell.innerHTML = strAux;
						oRow.appendChild(oCell);

						oCell = document.createElement("TD");
						oNodes = xmlDoc.getElementsByTagName("bairro")[i];
						if (oNodes.childNodes.length > 0) strAux = oNodes.childNodes[0].nodeValue; else strAux = "";
						if (strAux == null) strAux = "";
						if (strAux == "") strAux = "&nbsp;";
						oCell.vAlign = "top";
						oCell.innerHTML = strAux;
						oRow.appendChild(oCell);

						oCell = document.createElement("TD");
						oNodes = xmlDoc.getElementsByTagName("logradouro_nome")[i];
						if (oNodes.childNodes.length > 0) strAux = oNodes.childNodes[0].nodeValue; else strAux = "";
						if (strAux == null) strAux = "";
						if (strAux == "") strAux = "&nbsp;";
						oCell.vAlign = "top";
						oCell.innerHTML = strAux;
						oRow.appendChild(oCell);

						oCell = document.createElement("TD");
						oNodes = xmlDoc.getElementsByTagName("logradouro_complemento")[i];
						if (oNodes.childNodes.length > 0) strAux = oNodes.childNodes[0].nodeValue; else strAux = "";
						if (strAux == null) strAux = "";
						if (strAux == "") strAux = "&nbsp;";
						oCell.vAlign = "top";
						oCell.innerHTML = strAux;
						oRow.appendChild(oCell);
					}
				}
				catch (e) {
					alert("Falha na consulta!!");
				}
			}

			window.status = "Concluído";
			$("#divAjaxRunning").hide();

			//  RETORNOU APENAS 1 REGISTRO?
			if (intQtdeLinhas == 1) {
				rb_check[1].checked = true;
				try {
					c_endereco_numero.focus();
				}
				catch (e) {
					// NOP
				}
			}
		}
	}

	function ExecutaPesquisaCEP(OpcaoPesquisaPor) {
		var strUrl, strCep;

		strUrl = "";
		strCep = "";

		objAjaxPesqCep = GetXmlHttpObject();
		if (objAjaxPesqCep == null) {
			alert("O browser NÃO possui suporte ao AJAX!!");
			return;
		}

		if (OpcaoPesquisaPor == OPCAO_PESQUISA_POR_CEP) {

			strCep = retorna_so_digitos(trim(c_cep_pesq.value));
			if ((strCep.length != 5) && (strCep.length != 8)) {
				alert("CEP com tamanho inválido!!");
				c_cep_pesq.focus();
				return;
			}
			window.status = "Aguarde, pesquisando o CEP " + c_cep_pesq.value + " ...";
		}
		else if (OpcaoPesquisaPor == OPCAO_PESQUISA_POR_ENDERECO) {
			if (trim(c_uf_pesq.value) == "") {
				alert("Informe a UF do endereço a ser pesquisado!!");
				c_uf_pesq.focus();
				return;
			}
			else if (!uf_ok(trim(c_uf_pesq.value))) {
				alert("UF inválida!!");
				c_uf_pesq.focus();
				return;
			}
			else if (trim(c_localidade_pesq.value) == "") {
				alert("Informe a localidade do endereço a ser pesquisado!!");
				c_localidade_pesq.focus();
				return;
			}
			window.status = "Aguarde, executando a pesquisa...";
		}
		else {
			alert("Opção de pesquisa inválida!!");
			return;
		}

		//  Limpa tabela com resultados
		LimpaTabelaResultado();

		$("#divAjaxRunning").fadeIn(100);

		strUrl = "../GLOBAL/AjaxCepPesqBD.asp";
		if (OpcaoPesquisaPor == OPCAO_PESQUISA_POR_ENDERECO) {
			strUrl = strUrl + "?endereco=" + c_endereco_pesq.value + "&uf=" + c_uf_pesq.value + "&localidade=" + c_localidade_pesq.value;
		}
		else {
			strUrl = strUrl + "?cep=" + c_cep_pesq.value;
		}
		strUrl = strUrl + "&opcao=" + OpcaoPesquisaPor;
		//  Prevents server from using a cached file
		strUrl = strUrl + "&sid=" + Math.random() + Math.random();
		objAjaxPesqCep.onreadystatechange = TrataRespostaAjaxPesquisaCEP;
		objAjaxPesqCep.open("GET", strUrl, true);
		objAjaxPesqCep.send(null);
	}

	/* fim bloco 01 */

	/* Bloco 2 - Funções manipulação via JQuery da janela de pesquisa de CEP */
	/* início bloco 02*/

	var campo_cep, campo_uf, campo_localidade, campo_bairro, campo_logradouro, campo_endnumero, campo_endcompl;

	function atribuirCamposEndereco() {
		if (ConfirmarOperacao()) {
			$("input[name='" + campo_cep + "']").val(strCep);
			$("input[name='" + campo_uf + "']").val(strUF);
			$("input[name='" + campo_localidade + "']").val(strLocalidade);
			$("input[name='" + campo_bairro + "']").val(strBairro);
			$("input[name='" + campo_logradouro + "']").val(strLogradouro);
			$("input[name='" + campo_endnumero + "']").val(strEnderecoNumero);
			$("input[name='" + campo_endcompl + "']").val(strEnderecoComplemento);
			fechaJanelaCEP();
		}
	}

	function sizeDivAjaxRunning() {
		var newTop = $(window).scrollTop() + "px";
		$("#divAjaxRunning").css("top", newTop);
	}

	function sizeDivBaseJanelaBuscaCEP() {
		var newHeight = $(document).height() + "px";
		$("#divBaseJanelaBuscaCEP").css("height", newHeight);
	}

	function fechaJanelaCEP() {
		$(window).scrollTop(windowScrollTopAnterior);
		$("#divBaseJanelaBuscaCEP").hide();
	}

	$.mostraJanelaCEP = function(seletor_cep, seletor_uf, seletor_localidade, seletor_bairro, seletor_logradouro, seletor_endnumero, seletor_endcompl) {
		if ((seletor_cep == null) || (seletor_cep == "")) {
			$("#bFechar").show();
			$("#bCancelar").hide();
			$("#bConfirmar").hide();
			$("#rowSeparadorNumeroEComplemento").hide();
			$("#rowNumeroEComplemento").hide();
			$("#rowEspacadorQuandoSomenteConsulta").show();
			campo_cep = "";
			campo_uf = "";
			campo_localidade = "";
			campo_bairro = "";
			campo_logradouro = "";
			campo_endnumero = "";
			campo_endcompl = "";
		}
		else {
			$("#bFechar").hide();
			$("#bCancelar").show();
			$("#bConfirmar").show();
			$("#rowSeparadorNumeroEComplemento").show();
			$("#rowNumeroEComplemento").show();
			$("#rowEspacadorQuandoSomenteConsulta").hide();
			campo_cep = seletor_cep;
			campo_uf = seletor_uf;
			campo_localidade = seletor_localidade;
			campo_bairro = seletor_bairro;
			campo_logradouro = seletor_logradouro;
			campo_endnumero = seletor_endnumero;
			campo_endcompl = seletor_endcompl;
		}

		windowScrollTopAnterior = $(window).scrollTop();

		sizeDivBaseJanelaBuscaCEP();
		sizeDivAjaxRunning();

		$("#c_cep_pesq").val("");
		$("#c_endereco_pesq").val("");
		$("#c_endereco_numero").val("");
		$("#c_endereco_complemento").val("");

		var altDelta, altWindowBase, altWindowReal, altNova;
		var altMin = 135;
		//o teste abaixo visa detectar se o navegador é o IE 11 ou posterior, visto que, a partir desta versão,
		//a Microsoft modificou a string de detecção de versão
		if ((navigator.userAgent.toUpperCase().indexOf(".NET") > -1) || (navigator.userAgent.toUpperCase().indexOf(" RV:") > -1)) {
			altMin = 189;
		}

		altWindowBase = 550;
		altWindowReal = $(window).height();
		altDelta = altWindowReal - altWindowBase;
		if (altDelta < 0) altDelta = 0;
		altNova = altMin + parseInt(0.8 * altDelta, 10);
		$("#divResultado").css({ "height": altNova + "px" });

		var largura;
		largura = retorna_so_digitos($("#divResultado").css("width"));
		largura = largura - 18;
		$("#tabResposta").css({ "width": largura + "px" });

		LimpaListaLocalidades();
		LimpaTabelaResultado();

		$("#divBaseJanelaBuscaCEP").fadeIn(100);

		$("#c_uf_pesq").val("");
		c_cep_pesq.focus();

		var cep_formatado = cep_formata($("input[name='" + seletor_cep + "']").val());
		if (cep_formatado != "") {
			$("#c_cep_pesq").val(cep_formatado);
			if ($("#c_cep_pesq").val() != "") {
				ExecutaPesquisaCEP(OPCAO_PESQUISA_POR_CEP);
			}
		}
	};

	$(document).ready(function() {

		$('#divTabJanelaBuscaCEP').addClass('divFixo');

		$("#divAjaxRunning").hide(); // Mantém oculto inicialmente
		$("#divAjaxRunning").css('filter', 'alpha(opacity=60)'); // TRANSPARÊNCIA NO IE8

		sizeDivBaseJanelaBuscaCEP();
		sizeDivAjaxRunning();

		//Every resize of window
		$(window).resize(function() {
			sizeDivAjaxRunning();
			sizeDivBaseJanelaBuscaCEP();
		});

		//Every scroll of window
		$(window).scroll(function() {
			sizeDivAjaxRunning();
		});

		//habilitando a tecla ESC para fechar a tela (para a div principal da janela)
		$("#divBaseJanelaBuscaCEP").keydown(function(event) {
			if (event.which == 27) {
				event.preventDefault();
				fechaJanelaCEP();
			}
		});

		//habilitando a tecla ESC para fechar a tela (para os filhos da div principal da janela)
		$("#divBaseJanelaBuscaCEP > *").keydown(function(event) {
			if (event.which == 27) {
				event.preventDefault();
				fechaJanelaCEP();
			}
		});

		$("#bConfirmar").click(function() {
			atribuirCamposEndereco();
		});

		$("#c_endereco_complemento").keydown(function(event) {
			if (event.which == 13) {
				event.preventDefault();
				atribuirCamposEndereco();
			}
		});

		$("#bCancelar").click(function() {
			fechaJanelaCEP();
		});

		$("#bFechar").click(function() {
			fechaJanelaCEP();
		});

		$("#imgFechaJanelaPesqCEP").click(function() {
			fechaJanelaCEP();
		});

		$("#bPesquisaCEP").click(function() {
			ExecutaPesquisaCEP(OPCAO_PESQUISA_POR_CEP);
		});

		$("#bPesquisaEndereco").click(function() {
			ExecutaPesquisaCEP(OPCAO_PESQUISA_POR_ENDERECO);
		});

		$("#c_uf_pesq").change(function() {
			$.CepUfCarregaLocalidades();
		});
	});

	/* fim bloco 02*/

})(jQuery);
