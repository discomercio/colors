// the semi-colon before the function invocation is a safety 
// net against concatenated scripts and/or other plugins 
// that are not closed properly.
;var HHO = {};
;(function($) {
	HHO = {
		digitouEnter: function(event) {
			if ((event.which == 13) || (event.which == 10) || (event.keyCode == 13) || (event.keyCode == 10)) {
				return true;
			}
			return false;
		},
		temInfo: function(texto) {
			var s = "" + texto;
			if (s.length > 0) return true;
			return false;
		},
		digitacaoCnpjCpfOk: function(event) {
			var code = (event.keyCode ? event.keyCode : event.which);
			var letra = String.fromCharCode(code);
			// Aceita BACKSPACE(8), TAB(9), ESCAPE(27), PAGE UP(33), PAGE DOWN(34), END(35), HOME(36) LEFT ARROW(37), UP ARROW(38), RIGHT ARROW(39), DOWN ARROW(40)
			if ((code == 8) || (code == 9) || (code == 27) || (code == 33) || (code == 34) || (code == 35) || (code == 36) || (code == 37) || (code == 38) || (code == 39) || (code == 40)) return true;
			if (((letra < "0") || (letra > "9")) && (letra != ".") && (letra != "/") && (letra != "-")) return false;
			return true;
		},
		digitacaoNumPedidoOk: function(event) {
			var code = (event.keyCode ? event.keyCode : event.which);
			var letra = String.fromCharCode(code);
			// Aceita BACKSPACE(8), TAB(9), ESCAPE(27), PAGE UP(33), PAGE DOWN(34), END(35), HOME(36) LEFT ARROW(37), UP ARROW(38), RIGHT ARROW(39), DOWN ARROW(40)
			if ((code == 8) || (code == 9) || (code == 27) || (code == 33) || (code == 34) || (code == 35) || (code == 36) || (code == 37) || (code == 38) || (code == 39) || (code == 40)) return true;
			if ((!isDigit(letra)) && (!isLetra(letra)) && (letra != COD_SEPARADOR_FILHOTE)) return false;
			return true;
		},
		digitacaoMoedaPositivoOk: function(event) {
			var code = (event.keyCode ? event.keyCode : event.which);
			var letra = String.fromCharCode(code);
			// Aceita BACKSPACE(8), TAB(9), ESCAPE(27), PAGE UP(33), PAGE DOWN(34), END(35), HOME(36) LEFT ARROW(37), UP ARROW(38), RIGHT ARROW(39), DOWN ARROW(40)
			if ((code == 8) || (code == 9) || (code == 27) || (code == 33) || (code == 34) || (code == 35) || (code == 36) || (code == 37) || (code == 38) || (code == 39) || (code == 40)) return true;
			if (((letra < "0") || (letra > "9")) && (letra != ".") && (letra != ",")) return false;
			return true;
		}
	};
})(jQuery);
