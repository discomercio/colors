// the semi-colon before the function invocation is a safety 
// net against concatenated scripts and/or other plugins 
// that are not closed properly.
; (function($) {

	/* Bloco 2 - Funções manipulação via JQuery da janela de pesquisa de EAN */
	/* início bloco 02*/

    var campo_ean;
    var titulo_ean;
    var cod_prod_xml;

	function atribuirCampoEAN() {
	    var strEan = $("#ed_EAN").val();
	    $("input[name='" + campo_ean + "']").val(strEan);
	    $("#" + titulo_ean).attr('title', strEan);
		fechaJanelaEAN();
	}

	function sizeDivAjaxRunning() {
	    var newTop = $(window).scrollTop() + "px";
	    $("#divAjaxRunning").css("top", newTop);
	}

	function sizeDivBaseJanelaEditaEAN() {
	    var newHeight = $(document).height() + "px";
	    var myTop = ($(document).width()) / 2 + "px";
	    $("#divBaseJanelaTrataEAN").css("height", newHeight);
	    //$("#divTabJanelaTrataEAN").css("top", myTop);
	   //$("#divBaseJanelaTrataEAN").css("top", "300px");
	}

	function fechaJanelaEAN() {
		$(window).scrollTop(windowScrollTopAnterior);
		$("#divBaseJanelaTrataEAN").hide();
	}

	$.mostraJanelaEAN = function(seletor_ean, seletor_titulo, seletor_prod) {
	    if (seletor_ean != null) {
	        campo_ean = seletor_ean;
	        titulo_ean = seletor_titulo;
	        cod_prod_xml = seletor_prod;
	        $("#ed_EAN").val($("#" + campo_ean).val());
	        $("#ed_cod_xml").text($("#" + cod_prod_xml).val());
		}

		windowScrollTopAnterior = $(window).scrollTop();

		sizeDivBaseJanelaEditaEAN();
		sizeDivAjaxRunning();

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
		//$("#divResultado").css({ "height": altNova + "px" });

		var largura;
		//largura = retorna_so_digitos($("#divResultado").css("width"));
		largura = largura - 18;
		//$("#tabResposta").css({ "width": largura + "px" });

		$("#divBaseJanelaTrataEAN").fadeIn(100);

	};

	$(document).ready(function() {

	    $("#divAjaxRunning").hide(); // Mantém oculto inicialmente
	    $("#divAjaxRunning").css('filter', 'alpha(opacity=60)'); // TRANSPARÊNCIA NO IE8

	    sizeDivBaseJanelaEditaEAN();
	    sizeDivAjaxRunning();

	    //Every resize of window
		$(window).resize(function() {
			sizeDivBaseJanelaEditaEAN();
		});

		//Every scroll of window
		$(window).scroll(function() {
			sizeDivAjaxRunning();
		});

		//habilitando a tecla ESC para fechar a tela (para a div principal da janela)
		$("#divBaseJanelaTrataEAN").keydown(function (event) {
			if (event.which == 27) {
				event.preventDefault();
				fechaJanelaEAN();
			}
		});

		//habilitando a tecla ESC para fechar a tela (para os filhos da div principal da janela)
		$("#divBaseJanelaTrataEAN > *").keydown(function (event) {
			if (event.which == 27) {
				event.preventDefault();
				fechaJanelaEAN();
			}
		});

		$("#bConfirmar").click(function() {
			atribuirCampoEAN();
		});

		$("#bCancelar").click(function() {
			fechaJanelaEAN();
		});

		$("#bFechar").click(function () {
		    fechaJanelaEAN();
		});

		$("#imgFechaJanelaEAN").click(function () {
			fechaJanelaEAN();
		});

	});

	/* fim bloco 02*/

})(jQuery);
