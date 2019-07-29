if (typeof document.compatMode != 'undefined' && document.compatMode != 'BackCompat') {
	var css_logo_ssl_aux = "_top:expression(document.documentElement.scrollTop+document.documentElement.clientHeight-this.clientHeight);_left:expression(document.documentElement.scrollLeft + document.documentElement.clientWidth - offsetWidth);}";
} else {
	var css_logo_ssl_aux = "_top:expression(document.body.scrollTop+document.body.clientHeight-this.clientHeight);_left:expression(document.body.scrollLeft + document.body.clientWidth - offsetWidth);}";
}
var css_div_logo_ssl = '#div_logo_ssl{position:fixed;';
css_div_logo_ssl = css_div_logo_ssl + '_position:absolute;';
css_div_logo_ssl = css_div_logo_ssl + 'bottom:0px;';
css_div_logo_ssl = css_div_logo_ssl + 'right:0px;';
css_div_logo_ssl = css_div_logo_ssl + css_logo_ssl_aux;
document.write('<style type="text/css">' + css_div_logo_ssl + '</style>');

function logo_ssl_corner(url_logo_ssl) {
	document.write('<div id="div_logo_ssl" class="notPrint">');
	document.write('<a href="http://www.instantssl.com" target="_blank"><img src=' + url_logo_ssl + ' alt="logotipo do certificado SSL" border="0"></a>');
	document.write('</div>');
}
