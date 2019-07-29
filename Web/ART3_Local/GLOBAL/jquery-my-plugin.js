// the semi-colon before the function invocation is a safety 
// net against concatenated scripts and/or other plugins 
// that are not closed properly.
; (function($) {
	var metodos = {
		init: function(opcoes) {
			// NOP
		},
		focusNext: function() {
			var campos, idx;
			campos = $(document).find("input:text:enabled:visible:not([readonly])");
			idx = campos.index(this);
			if ((idx > -1) && ((idx + 1) < campos.length)) {
				campos.eq(idx + 1).focus();
			}
			return this;
		},

		fix_radios: function() {
			function focus() {
				// if this isn't checked then no option is yet selected. bail
				if (!this.checked) return;

				// if this wasn't already checked, manually fire a change event
				if (!this.was_checked) {
					$(this).change();
				}
			}

			function change(e) {
				// shortcut if already checked to stop IE firing again
				if (this.was_checked) {
					e.stopImmediatePropagation();
					return;
				}

				// reset all the was_checked properties
				$("input[name=" + this.name + "]").each(function() {
					this.was_checked = this.checked;
				});
			}

			// attach the handlers and return so chaining works
			return this.focus(focus).change(change);
		}
	};



	$.fn.hUtil = function(metodo) {
		if (metodos[metodo]) {
			return metodos[metodo].apply(this, Array.prototype.slice.call(arguments, 1));
		} else if (typeof metodo === 'object' || !metodo) {
			return metodos.init.apply(this, arguments);
		} else {
			$.error('Método ' + metodo + ' não existe neste plugin!!');
		}
	};
})(jQuery);
