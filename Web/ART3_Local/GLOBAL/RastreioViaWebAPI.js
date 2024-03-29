//<SCRIPT>

// Realiza a consulta da p�gina via WebAPI e carrega o conte�do no iframe para contornar o problema do
// header x-frame-options que impede a exibi��o da p�gina dentro do iframe.
// Esse problema faz com que o navegador exiba a mensagem de erro "Este conte�do n�o pode ser exibido em um quadro" e
// exibe um link para que a p�gina seja exibida em uma nova guia.
function executaRastreioConsultaViaWebApiView(urlRastreioCompleto, urlRastreioBase, urlWebApi, usuario, sessionToken, id_iframe_rastreio, id_div_rastreio) {
	var newHeight = $(document).height() + "px";
	$(id_div_rastreio).css("height", newHeight);

	var jqxhr = $.ajax({
		url: urlWebApi,
		type: 'GET',
		data: {
			usuario: usuario,
			sessionToken: sessionToken
		},
		headers: {
			"X-Query-Url-Get": urlRastreioCompleto
		}
	})
		.done(function (response) {
			if (response.length == 0) {
				var myFrameView = $(id_iframe_rastreio).contents().find('body');
				myFrameView.html("Falha ao tentar consultar a p�gina de rastreamento!");
				$(id_div_rastreio).fadeIn();
				return;
			}

			// Remove links de bibliotecas javascript
			var idxJsStart, idxLastJsStart, idxJsEnd, strToDelete;
			var nIteracoes = 0;
			idxLastJsStart = 0;
			while (response.indexOf("<script ", idxLastJsStart) > -1) {
				// Contador de seguran�a p/ assegurar que nunca ficar� em loop infinito
				nIteracoes++;
				idxJsStart = response.indexOf("<script ", idxLastJsStart);
				if (idxJsStart > -1) idxLastJsStart = idxJsStart;
				idxJsEnd = response.indexOf("</script>", idxJsStart + 1);
				if ((idxJsStart > -1) && (idxJsEnd > -1) && (idxJsEnd > idxJsStart)) {
					strToDelete = response.substring(idxJsStart, idxJsEnd + "</script>".length);
					if (strToDelete.length > 0) {
						response = response.replace(strToDelete, "");
					}
					else {
						// Prossegue a partir do ponto em que se encontra
						idxLastJsStart = idxJsStart + 1;
					}
				}
				else {
					// Prossegue a partir do ponto em que se encontra
					idxLastJsStart = idxJsStart + 1;
				}

				if (idxLastJsStart >= (response.length - 1)) break;
				if (nIteracoes > 20) break;
			}

			// Acerta as href que n�o tenham a especifica��o completa do endere�o, caso contr�rio ir�o ser direcionadas p/ o site do sistema
			response = replaceAll(response, 'href="/', 'href="' + urlRastreioBase + '/');
			// Remove o bot�o Fechar
			response = response.replace('<tr><td><a href=# onclick="window.open(\'\',\'_self\').close();">Fechar</a></td></tr>', '');
			// Carrega o conte�do da p�gina no iframe
			var myFrameView = $(id_iframe_rastreio).contents().find('body');
			myFrameView.html(response);
			// Esconde as linhas que informam os dados de contato com a transportadora
			var spnFaleConosco = $(id_iframe_rastreio).contents().find("span").filter(function () { return ($(this).text() === "Fale conosco") });
			var trFaleConosco = spnFaleConosco.closest("tr");
			trFaleConosco.nextAll("tr").hide();
			trFaleConosco.hide();
			// Esconde o link "Processado por ssw.inf.br" que abre uma p�gina com as op��es: "Sou COMPRADOR e gostaria de rastrear minha mercadoria" e "Sou TRANSPORTADOR e gostaria de conhecer o Sistema SSW"
			var anchorProcessadoSSW = $(id_iframe_rastreio).contents().find("a").filter(function () { return ($(this).text() === "Processado por ssw.inf.br") });
			var divProcessadoSSW = anchorProcessadoSSW.closest("div");
			divProcessadoSSW.hide();
			// Mant�m vis�vel somente o nome da transportadora
			var spnNomeTransportadora = $(id_iframe_rastreio).contents().find("span").filter(function () { return ($(this).text() === "Estamos transportando a sua mercadoria:") });
			var trNomeTransportadora = spnNomeTransportadora.closest("tr");
			trNomeTransportadora.show();
			$(id_div_rastreio).fadeIn();
		})
		.fail(function (jqXHR, textStatus) {
			var myFrameView = $(id_iframe_rastreio).contents().find('body');
			myFrameView.html("Falha ao tentar consultar a p�gina de rastreamento!");
			$(id_div_rastreio).fadeIn();
		});
}
