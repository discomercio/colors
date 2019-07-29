// the semi-colon before the function invocation is a safety 
// net against concatenated scripts and/or other plugins 
// that are not closed properly.
; (function($) {
	var metodosUI = {
		init: function(opcoes) {
			// NOP
		},

		datepicker_filtro_inicial: function() {
			this.datepicker($.datepicker.regional['pt-BR'])
				.datepicker("option", {
					showOn: "button",
					buttonImage: "../imagem/jquery/calendar3.gif",
					changeMonth: true,
					changeYear: true,
					numberOfMonths: 1,
					showCurrentAtPos: 0,
					firstDay: 0,
					dayNamesMin: ["Dom", "Seg", "Ter", "Qua", "Qui", "Sex", "Sáb"],
					showButtonPanel: true
				});
			return this;
		},

		datepicker_filtro_final: function() {
			this.datepicker($.datepicker.regional['pt-BR'])
				.datepicker("option", {
					showOn: "button",
					buttonImage: "../imagem/jquery/calendar3.gif",
					changeMonth: true,
					changeYear: true,
					numberOfMonths: 1,
					showCurrentAtPos: 0,
					firstDay: 0,
					dayNamesMin: ["Dom", "Seg", "Ter", "Qua", "Qui", "Sex", "Sáb"],
					showButtonPanel: true
				});
			return this;
		},

		datepicker_padrao: function() {
			this.datepicker($.datepicker.regional['pt-BR'])
				.datepicker("option", {
					showOn: "button",
					buttonImage: "../imagem/jquery/calendar3.gif",
					changeMonth: true,
					changeYear: true,
					numberOfMonths: 1,
					showCurrentAtPos: 0,
					firstDay: 0,
					dayNamesMin: ["Dom", "Seg", "Ter", "Qua", "Qui", "Sex", "Sáb"],
					showButtonPanel: true
				});
			return this;
		},

		dialog_modal: function() {
			this.dialog({
				autoOpen: false,
				modal: true,
				resizable: true,
				title: "Aviso",
				minHeight: 200,
				minWidth: 600,
				buttons: [{ text: "Ok", click: function() { $(this).dialog("close"); } }]
			});
			return this;
		}
	};



	$.fn.hUtilUI = function(metodo) {
		if (metodosUI[metodo]) {
			return metodosUI[metodo].apply(this, Array.prototype.slice.call(arguments, 1));
		} else if (typeof metodo === 'object' || !metodo) {
			return metodosUI.init.apply(this, arguments);
		} else {
			$.error('Método ' + metodo + ' não existe neste plugin!!');
		}
	};
})(jQuery);
