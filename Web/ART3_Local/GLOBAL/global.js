//<SCRIPT>
var COD_SITE_ARTVEN_BONSHOP = "ArtBS";
var COD_SITE_ARTVEN_FABRICANTE = "ArtFab";
var COD_SITE_ASSISTENCIA_TECNICA = "AssTec";

var COR_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__INEXISTENTE = "#EE80EE";
var COR_AJAX_CONSULTA_DADOS_PRODUTO__INEXISTENTE = "#EE80EE";

var MAX_DECIMAIS_COEFICIENTE_CUSTO_FINANCEIRO_FORNECEDOR = 6;

var COMISSAO_INDICADOR_PERC_DESCONTO_SEM_NF = 16;

var OP_CONSULTA = "C";
var OP_INCLUI   = "I";
var OP_EXCLUI   = "E";
var OP_ALTERA   = "A";

var ID_VENDEDOR      = "V";
var ID_SEPARADOR     = "S";
var ID_ADMINISTRADOR = "A";
var ID_GERENCIAL     = "G";

var SIMBOLO_MONETARIO = "R$";

var OP_CEN_BLOCO_NOTAS_PEDIDO_LEITURA = "22700";
var OP_CEN_BLOCO_NOTAS_PEDIDO_CADASTRAMENTO = "22800";
var OP_LJA_BLOCO_NOTAS_PEDIDO_LEITURA = "55600";
var OP_LJA_BLOCO_NOTAS_PEDIDO_CADASTRAMENTO = "55700";
var OP_CEN_PEDIDO_CHAMADO_LEITURA_QUALQUER_CHAMADO = "28000";
var OP_CEN_PEDIDO_CHAMADO_CADASTRAMENTO	= "28100";
var OP_CEN_PEDIDO_CHAMADO_ESCREVER_MSG_QUALQUER_CHAMADO = "28200";
var OP_LJA_PEDIDO_CHAMADO_LEITURA_QUALQUER_CHAMADO = "58300";
var OP_LJA_PEDIDO_CHAMADO_CADASTRAMENTO = "58400";
var OP_LJA_PEDIDO_CHAMADO_ESCREVER_MSG_QUALQUER_CHAMADO = "58500";

// FORMA DE PAGAMENTO
var COD_FORMA_PAGTO_A_VISTA = "1";
var COD_FORMA_PAGTO_PARCELADO_CARTAO = "2";
var COD_FORMA_PAGTO_PARCELADO_COM_ENTRADA = "3";
var COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA = "4";
var COD_FORMA_PAGTO_PARCELA_UNICA = "5";
var COD_FORMA_PAGTO_PARCELADO_CARTAO_MAQUINETA = "6";

var ID_FORMA_PAGTO_DINHEIRO = "1";
var ID_FORMA_PAGTO_DEPOSITO = "2";
var ID_FORMA_PAGTO_CHEQUE = "3";
var ID_FORMA_PAGTO_BOLETO = "4";
var ID_FORMA_PAGTO_CARTAO = "5";
var ID_FORMA_PAGTO_BOLETO_AV = "6";
var ID_FORMA_PAGTO_CARTAO_MAQUINETA = "7";

var COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__COM_ENTRADA = "CE";
var COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__SEM_ENTRADA = "SE";
var COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__A_VISTA = "AV";

var ID_PF = "PF";
var ID_PJ = "PJ";

var CIELO_BANDEIRA__VISA = "visa";
var CIELO_BANDEIRA__MASTERCARD = "mastercard";
var CIELO_BANDEIRA__AMEX = "amex";
var CIELO_BANDEIRA__ELO = "elo";
var CIELO_BANDEIRA__HIPERCARD = "hipercard";
var CIELO_BANDEIRA__DINERS = "diners";
var CIELO_BANDEIRA__DISCOVER = "discover";
var CIELO_BANDEIRA__AURA = "aura";
var CIELO_BANDEIRA__JCB = "jcb";
var CIELO_BANDEIRA__CELULAR = "celular";

var BRASPAG_BANDEIRA__VISA = "visa";
var BRASPAG_BANDEIRA__MASTERCARD = "mastercard";
var BRASPAG_BANDEIRA__AMEX = "amex";
var BRASPAG_BANDEIRA__ELO = "elo";
var BRASPAG_BANDEIRA__HIPERCARD = "hipercard";
var BRASPAG_BANDEIRA__DINERS = "diners";
var BRASPAG_BANDEIRA__DISCOVER = "discover";
var BRASPAG_BANDEIRA__AURA = "aura";
var BRASPAG_BANDEIRA__JCB = "jcb";
var BRASPAG_BANDEIRA__CELULAR = "celular";

var VISANET_TIPO_CARTAO_CREDITO = "1";
var VISANET_TIPO_CARTAO_DEBITO = "2";
var MASTERCARD_TIPO_CARTAO_CREDITO = "3";

var MAX_TAM_OBS1 = 800;
var MAX_TAM_FORMA_PAGTO = 250;
var MAX_TAM_MENSAGEM_BLOCO_NOTAS = 400;
var MAX_TAM_DESCRICAO_OCORRENCIAS_EM_PEDIDOS = 240;
var MAX_TAM_MENSAGEM_OCORRENCIAS_EM_PEDIDOS = 1200;
var MAX_TAM_MENSAGEM_BLOCO_NOTAS_EM_ITEM_DEVOLVIDO = 400;
var MAX_TAM_NF_TEXTO = 800;
var MAX_TAM_DESCRICAO_CHAMADO_EM_PEDIDOS = 4000;
var MAX_TAM_MENSAGEM_CHAMADOS_EM_PEDIDOS = 2000;
var MAX_TAM_DESCRICAO_OBSERVACAO_DEVOLUCAO = 4000;

var MAX_TAM_OS_OBS_PROBLEMA = 80;
var MAX_TAM_OS_OBS_PECAS_NECESSARIAS = 400;

var COD_SEPARADOR_FILHOTE = "-";

var TAM_MIN_FABRICANTE		= 3;
var TAM_MIN_LOJA			= 2;
var TAM_MIN_GRUPO_LOJAS		= 2;
var TAM_MIN_MIDIA			= 3;
var TAM_MIN_PRODUTO			= 6;
var TAM_MIN_NUM_PEDIDO		= 6;	// SOMENTE PARTE NUMÉRICA DO NÚMERO DO PEDIDO
var TAM_MIN_ID_PEDIDO		= 7;	// PARTE NUMÉRICA DO NÚMERO DO PEDIDO + LETRA REFERENTE AO ANO
var TAM_MIN_NUM_ORCAMENTO	= 6;	// SOMENTE PARTE NUMÉRICA DO NÚMERO DO ORÇAMENTO
var TAM_MIN_ID_ORCAMENTO	= 7;	// PARTE NUMÉRICA DO NÚMERO DO ORÇAMENTO + LETRA (SUFIXO) QUE IDENTIFICA COMO ORÇAMENTO
var TAM_PLANO_CONTAS__EMPRESA = 2;
var TAM_PLANO_CONTAS__GRUPO = 2;
var TAM_PLANO_CONTAS__CONTA = 4;

var SUFIXO_ID_ORCAMENTO = "Z";

var ID_ESTOQUE_VENDA			= "VDA";
var ID_ESTOQUE_VENDIDO			= "VDO";
var ID_ESTOQUE_SEM_PRESENCA		= "SPE"; 
var ID_ESTOQUE_KIT				= "KIT";
var ID_ESTOQUE_SHOW_ROOM		= "SHR";
var ID_ESTOQUE_DANIFICADOS		= "DAN";
var ID_ESTOQUE_DEVOLUCAO		= "DEV";
var ID_ESTOQUE_ROUBO			= "ROU";
var ID_ESTOQUE_ENTREGUE		    = "ETG";

var KEY_LINEFEED	= String.fromCharCode(10);
var KEY_RETURN		= String.fromCharCode(13);
var KEY_BACK		= String.fromCharCode(8);
var KEY_CRLF		= KEY_RETURN + KEY_LINEFEED;
var KEY_LFCR		= KEY_LINEFEED + KEY_RETURN;
var KEY_ASPAS		= String.fromCharCode(34);
var KEY_APOSTROFE	= String.fromCharCode(39);
var KEYCODE_DELETE  = 46;


function replaceAll(str, find, replace) {
	return str.replace(new RegExp(find, 'g'), replace);
}

function trim( texto ){
var i,s,s_aux,tam;
	s = "" + texto;
	tam = s.length;

	s_aux = ""
	for( i=0; i<tam; i++ ) 
		if((s.charAt(i)!=" ")||(s_aux!="")) {
			s_aux+=s.charAt(i);
			}
	
	s = s_aux;
	tam = s.length;
	s_aux = ""
	for( i=(tam-1); i>=0; i-- ) 
		if ( ((s.charAt(i)!=" ")&&(s.charAt(i)!=KEY_RETURN)&&(s.charAt(i)!=KEY_LINEFEED)) || (s_aux!="") ) {
			s_aux = s.charAt(i) + s_aux;
			}
			
	return s_aux;
}

function left( texto, n ) {
var s;
	if (n <= 0) return "";
	texto = "" + texto;
	s=texto.substring(0, n);
	return s;
}

function right( texto, n ) {
var s;
	if (n <= 0) return "";
	texto = "" + texto;
	if (n >= texto.length) return texto;
	s=texto.substring(texto.length-n, texto.length);
	return s;
}

function substitui_caracteres(texto, antigo, novo) {
var i, s;
    texto="" + texto;
    s = "";
    for (i=0; i<texto.length; i++) {
        if (texto.charAt(i)==antigo) {
			if ((novo!="") && (novo!=String.fromCharCode(0))) s = s + novo;
			}
        else {
           s = s + texto.charAt(i);
           }
        }
    return s;
}

function isDigit(d){
	return ((d>='0')&&(d<='9'))
}

function isLetra(letra){
var c;
	c=ucase(trim("" + letra));
	return ((c>="A")&&(c<="Z"))
}

function tem_info( texto ){
var s;
	s = "" + texto;
	s = trim(s);
	if (s.length > 0) return true;
	return false;
}

function digitou_enter(limpa_enter){
	if (window.event.keyCode==13){
		if (limpa_enter) window.event.keyCode=0;
		return true;
		}
}

function digitou_char( char ){
	if (String.fromCharCode(window.event.keyCode)==char){
		return true;
		}
	return false;
}


/* IMPORTANTE - Função converte_data()
======================================
1) O parâmetro "strData" da função converte_data() deve ser um texto
   no formato DDMMYY, DDMMYYYY, DD/MM/YY ou DD/MM/YYYY
2) A função irá retornar um objeto do tipo Date ou null.
*/
function converte_data(strData) {
var val, s, dt;

	val = trim(strData);
	if (val == "") return null;

	/* SEPARADOR */
	var sep1 = parseInt(val.indexOf("/"), 10);
	var sep2 = parseInt(val.indexOf("/", sep1 + 1), 10);
	if ((val.length == 6) && (sep1 == -1) && (sep2 == -1)) {
		val = val.substr(0, 2) + "/" + val.substr(2, 2) + "/" + val.substr(4, 2);
		var sep1 = parseInt(val.indexOf("/"), 10);
		var sep2 = parseInt(val.indexOf("/", sep1 + 1), 10);
	}
	if ((val.length == 8) && (sep1 == -1) && (sep2 == -1)) {
		val = val.substr(0, 2) + "/" + val.substr(2, 2) + "/" + val.substr(4, 4);
		var sep1 = parseInt(val.indexOf("/"), 10);
		var sep2 = parseInt(val.indexOf("/", sep1 + 1), 10);
	}

	var len = parseInt(val.length, 10);

	s = val.substr(0, sep1);
	if (s.length == 0) return null;
	var dd = parseInt(s, 10);

	s = val.substr(sep1 + 1, sep2 - sep1 - 1);
	if (s.length == 0) return null;
	var mm = parseInt(s, 10);

	s = val.substr(sep2 + 1, len - sep2 - 1);
	if (s.length == 0) return null;
	var yy = parseInt(s, 10);

	/* ANO */
	if (yy <= 90) yy += 2000;
	if ((yy > 90) && (yy < 100)) yy += 1900;
	if ((yy < 1900) || (yy > 2099)) {
		return null;
	}

	var leap = ((yy == (parseInt(yy / 4, 10) * 4)) && !(yy == (parseInt(yy / 100, 10) * 100)));

	/* MES */
	if (!((mm >= 1) && (mm <= 12))) {
		return null;
	}

	/* DIA */
	if ((mm == 2) && (leap)) dom = 29;
	if ((mm == 2) && !(leap)) dom = 28;
	if ((mm == 1) || (mm == 3) || (mm == 5) || (mm == 7) || (mm == 8) || (mm == 10) || (mm == 12)) dom = 31;
	if ((mm == 4) || (mm == 6) || (mm == 9) || (mm == 11)) dom = 30;
	if (dd > dom) {
		return null;
	}

	dt = new Date(parseInt(yy), parseInt(mm) - 1, parseInt(dd));

	return dt;
}


/* IMPORTANTE - Função isDate()
   ============================	
   1) O parâmetro "d" da função isDate() deve ser um campo do formulário,
      pois a função irá referenciar a propriedade "value" do parâmetro.
      Tipicamente, se a função for chamada a partir de eventos do campo,
      a sintaxe deve ser: onblur="if (!isDate(this)) {...}"
      sendo que "this" está referenciando o próprio campo.
   2) A função irá alterar o conteúdo do campo para exibir a data formatada.
*/
function isDate(d) {
var val,s;

  d.value=trim(d.value);
  val = d.value;
  if (val == "") return true;
  
  /* SEPARADOR */
  var sep1 = parseInt(val.indexOf("/"),10);
  var sep2 = parseInt(val.indexOf("/",sep1+1),10);
  if ((val.length==6)&&(sep1==-1)&&(sep2==-1)) {
		val = val.substr(0,2)+"/"+val.substr(2,2)+"/"+val.substr(4,2);
  		var sep1 = parseInt(val.indexOf("/"),10);
		var sep2 = parseInt(val.indexOf("/",sep1+1),10);
		} 
  if ((val.length==8)&&(sep1==-1)&&(sep2==-1)) {
		val = val.substr(0,2)+"/"+val.substr(2,2)+"/"+val.substr(4,4);
  		var sep1 = parseInt(val.indexOf("/"),10);
		var sep2 = parseInt(val.indexOf("/",sep1+1),10);
		} 
  
  var len = parseInt(val.length,10);
  
  s=val.substr(0,sep1);
  if (s.length==0) return false;
  var dd = parseInt(s,10);
  
  s=val.substr(sep1+1,sep2-sep1-1);
  if (s.length==0) return false;
  var mm = parseInt(s,10);
  
  s=val.substr(sep2+1,len-sep2-1);
  if (s.length==0) return false;
  var yy = parseInt(s,10);
 
  /* ANO */
  if (yy<=90) yy+=2000;
  if ((yy>90)&&(yy<100)) yy+=1900;
  if ((yy<1900)||(yy>2099)) {
		return false;
		}  
  
  var leap = ((yy == (parseInt(yy/4,10) * 4)) && !(yy == (parseInt(yy/100,10) * 100)));
  
  /* MES */
  if (!((mm >= 1) && (mm <= 12))) {
		return false;
		}
		
  /* DIA */ 		
  if ((mm == 2) && (leap)) dom = 29;
  if ((mm == 2) && !(leap)) dom = 28;
  if ((mm == 1) || (mm == 3) || (mm == 5) || (mm == 7) || (mm == 8) || (mm == 10) || (mm == 12)) dom = 31;
  if ((mm == 4) || (mm == 6) || (mm == 9) || (mm == 11)) dom = 30;
  if (dd > dom) {
		return false;
		}
		
  if (dd < 10) dd = '0' + dd;		
  if (mm < 10) mm = '0' + mm;

  d.value = dd+'/'+mm+'/'+yy;		

  return true;
}	

function filtra_data(){
var letra;
	letra=String.fromCharCode(window.event.keyCode);
	if (!isDigit(letra)&&(window.event.keyCode!=47)&&(window.event.keyCode!=8)&&(window.event.keyCode!=13)){
		window.event.keyCode=0;
		}
}

function filtra_email() {
var letra;
	letra=String.fromCharCode(window.event.keyCode);
	if ((letra==" ")||(letra=="|")||(window.event.keyCode==34)||(window.event.keyCode==39)) window.event.keyCode=0;
}

function filtra_nome_identificador() {
var letra;
	letra=String.fromCharCode(window.event.keyCode);
	if ((letra=="|")||(window.event.keyCode==34)||(window.event.keyCode==39)) window.event.keyCode=0;
}

function filtra_agencia_bancaria() {
var letra;
	letra=String.fromCharCode(window.event.keyCode);
	letra=letra.toUpperCase();
	if (((letra<"0")||(letra>"9"))&&((letra<"A")||(letra>"Z"))&&(letra!=".")&&(letra!="-")&&(letra!="/")) window.event.keyCode=0;
//  Converte p/ maiusculas
	if ((window.event.keyCode > 96) && (window.event.keyCode < 123)) window.event.keyCode = window.event.keyCode - 32;
	
}

function filtra_agencia_bancaria_sem_digito() {
var letra;
	letra=String.fromCharCode(window.event.keyCode);
	letra=letra.toUpperCase();
	if (((letra<"0")||(letra>"9"))&&((letra<"A")||(letra>"Z"))&&(letra!=".")) window.event.keyCode=0;
//  Converte p/ maiusculas
	if ((window.event.keyCode > 96) && (window.event.keyCode < 123)) window.event.keyCode = window.event.keyCode - 32;
	
}

function filtra_conta_bancaria() {
var letra;
	letra=String.fromCharCode(window.event.keyCode);
	letra=letra.toUpperCase();
	if (((letra<"0")||(letra>"9"))&&((letra<"A")||(letra>"Z"))&&(letra!=".")&&(letra!="-")&&(letra!="/")) window.event.keyCode=0;
//  Converte p/ maiusculas
	if ((window.event.keyCode > 96) && (window.event.keyCode < 123)) window.event.keyCode = window.event.keyCode - 32;
}

function filtra_conta_bancaria_sem_digito() {
var letra;
	letra=String.fromCharCode(window.event.keyCode);
	letra=letra.toUpperCase();
	if (((letra<"0")||(letra>"9"))&&((letra<"A")||(letra>"Z"))&&(letra!=".")) window.event.keyCode=0;
//  Se for letra minúscula, converte p/ maiuscula
	if ((window.event.keyCode > 96) && (window.event.keyCode < 123)) window.event.keyCode = window.event.keyCode - 32;
}

function filtra_numerico() {
var letra;
	letra=String.fromCharCode(window.event.keyCode);
	if ((letra<"0")||(letra>"9")) window.event.keyCode=0;
}

function filtra_letra() {
var letra;
	letra=String.fromCharCode(window.event.keyCode);
	letra = letra.toUpperCase();
	if ((letra<"A")||(letra>"Z")) window.event.keyCode=0;
}

function filtra_alfanumerico() {
var letra;
	letra=String.fromCharCode(window.event.keyCode);
	letra = letra.toUpperCase();
	if ( ((letra<"A")||(letra>"Z")) && ((letra<"0")||(letra>"9")) ) window.event.keyCode=0;
}

function filtra_pedido() {
var letra;
	letra=String.fromCharCode(window.event.keyCode);
	letra=letra.toUpperCase();
	if ((!isDigit(letra))&&(!isLetra(letra))&&(letra!=COD_SEPARADOR_FILHOTE)) window.event.keyCode=0;
}

function filtra_orcamento() {
var letra;
	letra=String.fromCharCode(window.event.keyCode);
	letra=letra.toUpperCase();
	if ((!isDigit(letra))&&(letra!=SUFIXO_ID_ORCAMENTO)) window.event.keyCode=0;
}

function filtra_produto() {
var letra;
	letra=String.fromCharCode(window.event.keyCode);
	letra=letra.toUpperCase();
	if (((letra<"0")||(letra>"9"))&&((letra<"A")||(letra>"Z"))) window.event.keyCode=0;
}

function filtra_fabricante() {
var letra;
	letra=String.fromCharCode(window.event.keyCode);
	letra=letra.toUpperCase();
	if ((letra<"0")||(letra>"9")) window.event.keyCode=0;
}

function filtra_sexo() {
var letra;
	letra=String.fromCharCode(window.event.keyCode);
	letra=letra.toUpperCase();
	if ((letra!="M")&&(letra!="F")) window.event.keyCode=0;
}

function filtra_percentual() {
var letra;
	letra=String.fromCharCode(window.event.keyCode);
	if (((letra<"0")||(letra>"9"))&&(letra!=".")&&(letra!=",")) window.event.keyCode=0;
}

function filtra_coeficiente_custo_financ_fornecedor() {
var letra;
	letra=String.fromCharCode(window.event.keyCode);
	if (((letra<"0")||(letra>"9"))&&(letra!=".")&&(letra!=",")) window.event.keyCode=0;
}

function filtra_moeda() {
var letra;
	letra=String.fromCharCode(window.event.keyCode);
	if (((letra<"0")||(letra>"9"))&&(letra!=".")&&(letra!=",")&&(letra!="-")) window.event.keyCode=0;
}

function filtra_moeda_positivo() {
var letra;
	letra=String.fromCharCode(window.event.keyCode);
	if (((letra<"0")||(letra>"9"))&&(letra!=".")&&(letra!=",")) window.event.keyCode=0;
}

function filtra_num_real() {
var letra;
	letra=String.fromCharCode(window.event.keyCode);
	if (((letra<"0")||(letra>"9"))&&(letra!=".")&&(letra!=",")) window.event.keyCode=0;
}

function filtra_nextel() {
var c;
	c = String.fromCharCode(window.event.keyCode);
	if (((c < "0") || (c > "9")) && (letra != "*") && (letra != "#") && (letra != "-") && (letra != "(") && (letra != ")")) window.event.keyCode = 0;
}

function retorna_so_digitos(numero){
var i,s_num,s_resp;
	s_resp = "";
	s_num = "" + numero;
	for (i=0; i<s_num.length; i++)
		if (isDigit(s_num.charAt(i))) s_resp = s_resp + s_num.charAt(i);
	return s_resp;
}

function ucase(texto) {
var s;
	s = "" + texto;
	s = s.toUpperCase();
	return s;
}

function lcase(texto) {
var s;
	s = "" + texto;
	s = s.toLowerCase();
	return s;
}

function email_ok(email) {
var filtro_regex = /^([0-9a-zA-Z]([-.\w]*[0-9a-zA-Z][_]*)*@([0-9a-zA-Z][-\w]*\.)+[a-zA-Z]{2,9})$/;
	if (!filtro_regex.test(email)) return false;
	return true;
}

function sexo_ok(sexo) {
var s_sexo;
	s_sexo = "" + sexo;
	s_sexo = s_sexo.toUpperCase();
	if ((s_sexo=="M")||(s_sexo=="F")) return true;
}

function uf_ok(uf) {
var i, sigla;
    uf = ucase(trim(uf));
    if (uf == "") return true;
    if (uf.length != 2) return false;
	sigla = "AC AL AM AP BA CE DF ES GO MA MG MS MT PA PB PE PI PR RJ RN RO RR RS SC SE SP TO    ";
    for (i=0; i<(sigla.length-1); i++) {
		if (uf == sigla.substring(i,i+2)) return true;
		}
}

function cep_ok(cep) {
var s_cep;
	s_cep = "" + cep;
	s_cep = retorna_so_digitos(s_cep);
	if ((s_cep.length==0)||(s_cep.length==5)||(s_cep.length==8)) return true;
}

function cep_formata(cep) {
var s_cep;
	s_cep = "" + cep;
	s_cep = retorna_so_digitos(s_cep);
	if ((s_cep=="")||(s_cep.length==5)||(!cep_ok(s_cep))) return s_cep;
	s_cep=s_cep.substring(0,5)+"-"+s_cep.substring(5,8);
	return s_cep;
}

function filtra_cep() {
var letra;
	letra=String.fromCharCode(window.event.keyCode);
	if (((letra<"0")||(letra>"9")) && (letra!="-")) window.event.keyCode=0;
}

function ddd_ok(ddd) {
var s_ddd;
	s_ddd = "" + ddd;
	s_ddd = retorna_so_digitos(s_ddd);
	if ((s_ddd.length==0)||(s_ddd.length==2)) return true;
}

function telefone_ok(telefone) {
var s_tel;
	s_tel = "" + telefone;
	s_tel = retorna_so_digitos(s_tel);
	if ((s_tel.length==0)||(s_tel.length>=6)) return true;
}

function telefone_formata(telefone) {
var i,s_tel;
	s_tel = "" + telefone;
	s_tel = retorna_so_digitos(s_tel);
	if ((s_tel=="")||(s_tel.length>9)||(!telefone_ok(s_tel))) return s_tel;
	i=s_tel.length-4;
	s_tel=s_tel.substring(0,i)+"-"+s_tel.substring(i,s_tel.length+1);
	return s_tel;
}

function filtra_cnpj_cpf() {
var letra;
	letra=String.fromCharCode(window.event.keyCode);
	if (((letra<"0")||(letra>"9")) && (letra!=".") && (letra!="/") && (letra!="-")) window.event.keyCode=0;
}

function filtra_cnpj() {
var letra;
	letra=String.fromCharCode(window.event.keyCode);
	if (((letra<"0")||(letra>"9")) && (letra!=".") && (letra!="/") && (letra!="-")) window.event.keyCode=0;
}

function filtra_cpf() {
var letra;
	letra=String.fromCharCode(window.event.keyCode);
	if (((letra<"0")||(letra>"9")) && (letra!=".") && (letra!="/") && (letra!="-")) window.event.keyCode=0;
}

function cnpj_formata(cnpj) {
var s_cnpj;
	s_cnpj = "" + cnpj;
	s_cnpj = retorna_so_digitos(s_cnpj);
	if ((s_cnpj=="")||(!cnpj_ok(s_cnpj))) return s_cnpj;
	s_cnpj=s_cnpj.substring(0,2)+"."+s_cnpj.substring(2,5)+"."+s_cnpj.substring(5,8)+"/"+s_cnpj.substring(8,12)+"-"+s_cnpj.substring(12,14);
	return s_cnpj;
}

function cpf_formata(cpf) {
var s_cpf;
	s_cpf = "" + cpf;
	s_cpf = retorna_so_digitos(s_cpf);
	if ((s_cpf=="")||(!cpf_ok(s_cpf))) return s_cpf;
	s_cpf=s_cpf.substring(0,3)+"."+s_cpf.substring(3,6)+"."+s_cpf.substring(6,9)+"/"+s_cpf.substring(9,11);
	return s_cpf;
}

function cnpj_ok(cnpj) {
var d, i, p1, p2, tudo_igual;
var s_cnpj;

	s_cnpj = "" + cnpj;
	s_cnpj = retorna_so_digitos(s_cnpj);

	p1 = "543298765432";
	p2 = "6543298765432";

    if (s_cnpj == "") return true;
	if (s_cnpj.length!=14) return false;

 // DÍGITOS TODOS IGUAIS?
	tudo_igual=true;
	for (i=0; i<(s_cnpj.length-1); i++)
		if (s_cnpj.substring(i,i+1)!=s_cnpj.substring(i+1,i+2)) {
			tudo_igual=false;
			break;
			}

	if (tudo_igual) return false;
	
 // VERIFICA O PRIMEIRO CHECK DIGIT
    d = 0;
    for (i=0; i<12; i++)
        d = d + parseInt(p1.substring(i,i+1),10) * parseInt(s_cnpj.substring(i, i+1),10);

    d = 11 - (d % 11);
    if (d > 9) d = 0;
    if (d != parseInt(s_cnpj.substring(12,13),10)) return false;

 // VERIFICA O SEGUNDO CHECK DIGIT
    d = 0;
    for (i=0; i<13; i++)
        d = d + parseInt(p2.substring(i,i+1),10) * parseInt(s_cnpj.substring(i,i+1),10);

    d = 11 - (d % 11);
    if (d > 9) d = 0;
    if (d != parseInt(s_cnpj.substring(13,14),10)) return false;

	return true;
}

function cpf_ok(cpf) {
var d, i, tudo_igual;
var s_cpf;

	s_cpf = "" + cpf;
	s_cpf = retorna_so_digitos(s_cpf);

 // VERIFICA OS 'CHECK DIGITS' DO CPF
    if (s_cpf=="") return true;
	if (s_cpf.length != 11) return false;

 // DÍGITOS TODOS IGUAIS?
	tudo_igual=true;
	for (i=0; i<(s_cpf.length-1); i++)
		if (s_cpf.substring(i,i+1)!=s_cpf.substring(i+1,i+2)) {
			tudo_igual=false;
			break;
			}

	if (tudo_igual) return false;

//  VERIFICA O PRIMEIRO CHECK DIGIT
    d = 0;
    for (i=1; i<=9; i++)
        d = d + (11 - i) * parseInt(s_cpf.substring(i-1,i),10);

    d = 11 - (d % 11);
    if (d > 9) d = 0;
    if (d != parseInt(s_cpf.substring(9,10),10)) return false;

 // VERIFICA O SEGUNDO CHECK DIGIT
    d = 0;
    for (i=2; i<=10; i++)
        d = d + (12 - i) * parseInt(s_cpf.substring(i-1,i),10);
    
    d = 11 - (d % 11);
    if (d > 9) d = 0;
    if (d != parseInt(s_cpf.substring(10,11),10)) return false;

    return true;
}

function cnpj_cpf_ok( cnpj_cpf ) {
var s;
	s="" + cnpj_cpf;
	s=retorna_so_digitos(s);
	if (s.length==11) {
		if (cpf_ok(s)) return true;
		}
	else if (s.length==14) {
		if (cnpj_ok(s)) return true;
		}
	else if (s.length==0) {
		return true;
		}
	
	return false;
}

function cnpj_cpf_formata( cnpj_cpf ) {
var s;
	s="" + cnpj_cpf;
	s=retorna_so_digitos(s);
	if (s.length==11) {
		s=cpf_formata(s);
		}
	else if (s.length==14) {
		s=cnpj_formata(s);
		}
	
	return s;
}

function IntPart( valor ) {
var s, n_valor, s_valor, s_resp, i;
	if (isNaN(valor)) return 0;
	n_valor = valor * 1;
	s_valor = n_valor.toString();
	s_resp="";
	for (i=0; i < s_valor.length; i++) {
		s = s_valor.charAt(i);
		if (isDigit(s)||(s=="-")) s_resp=s_resp + s; else break;
		}
	if (isNaN(s_resp)) return 0;
	return (s_resp * 1);
}

function DecPart( valor ) {
var sinal, s, n_valor, s_valor, s_resp, i;
	if (isNaN(valor)) return 0;
	n_valor = valor * 1;
	if (n_valor < 0) sinal=-1; else sinal=1;
	s_valor = n_valor.toString();
	if (s_valor.indexOf(".")==-1) return 0;
	s_resp="";
	for (i=s_valor.length-1; i >= 0; i--) {
		s = s_valor.charAt(i);
		if (isDigit(s)) s_resp=s+s_resp; else break;
		}
	if (isNaN(s_resp)) return 0;
	s_resp = "0." + s_resp;
	return (s_resp * sinal);
}

function retorna_separador_decimal( numero ) {
var i, c, s_num, n_ponto, n_virg, s_ult_sep, n_digitos_finais;
	n_digitos_finais=0;
	n_ponto=0;
	n_virg=0;
	s_ult_sep="";
	s_num = "" + numero;
	for (i=s_num.length-1; i>=0; i--) {
		c=s_num.charAt(i);
		if (c==".") {
			n_ponto=n_ponto+1;
			if (s_ult_sep=="") s_ult_sep=c;
			}
		else if (c==",") {
			n_virg=n_virg+1;
			if (s_ult_sep=="") s_ult_sep=c;
			}
		if (isDigit(c)&&(n_ponto==0)&&(n_virg==0)) n_digitos_finais=n_digitos_finais+1;
		}

	if (s_ult_sep==".") {
		if ((n_ponto==1)&&(n_virg==0)&&(n_digitos_finais==3)) {
		/* NOP: CONSIDERA 123.456 COMO CENTO E VINTE E TRÊS MIL E QUATROCENTOS E CINQUENTA E SEIS */
			}
		else {
			if (n_ponto==1) return ".";
			}
		}
	else if (s_ult_sep==",") {
		if ((n_virg > 1)&&(n_ponto==0)) return ".";
		}
	return ",";
}

function converte_numero( valor ) {
var s, c, s_valor, s_sep, i;
	if (typeof valor == "number") return valor;
	s_valor="" + valor;
	s_sep=retorna_separador_decimal(s_valor);
	s_valor=substitui_caracteres(s_valor, s_sep, "V");
	s="";
	for (i=0; i<s_valor.length; i++) {
		c=s_valor.charAt(i);
		if ((!isDigit(c))&&(c!="-")&&(c!="V")) c="";
		s=s+c;
		}
	s_valor=substitui_caracteres(s, "V", ".");
	if (isNaN(s_valor)) return 0;
	return (s_valor * 1);
}

function formata_moeda_xml(in_valor) {
    var sinal, valor, decimais, valor_decimal, fator, i, j, n, c, s, s_valor, s_int, s_dec, s_resp, achou;
    valor = in_valor;
    s_valor = in_valor;
    decimais = 2;
    if (trim("" + valor) == "") return "";
    /* Verifca se o número é positivo ou negativo. */
    valor = converte_numero(valor);
    if (valor < 0) sinal = -1; else sinal = 1;
    valor = Math.abs(valor);
    if (isNaN(valor)) return "";
    /*  Separa a parte inteira e decimal */
    s_int = "";
    s_dec = "";
    achou = false;
    for (i = 0; i < s_valor.length; i++) {
        c = s_valor.charAt(i);
        if (c == ".") {
            achou = true;
        }
        else {
            if (!achou) s_int = s_int + c; else if (s_dec.length < decimais + 1) s_dec = s_dec + c;
        }
    }

    /*  Formata parte decimal com arredondamento */
    while (s_dec.length < decimais) s_dec = s_dec + "0";
    if (s_dec.length > decimais) {
        valor_decimal = s_dec;
        valor_decimal = converte_numero(valor_decimal);
        n = Math.round(valor_decimal/10)*10
        s = "" + n;
        //se o valor convertido for menor que 100, significa que a primeira casa decimal é 0
        if (n < 100) s = "0" + s;
        s_dec = left(s, decimais);
    }    


    /*  Formata parte inteira */
    s = "";
    j = 0;
    for (i = s_int.length; i >= 0; i--) {
        s = s_int.charAt(i) + s;
        if (((j % 3) == 0) && (i != s_int.length) && (i != 0)) s = "." + s;
        j = j + 1;
    }
    s_int = s;

    /*  Monta número formatado final */
    s = s_int;
    if (s_dec != "") {
        s = s + ",";
        s = s + s_dec;
    }
    if (sinal == -1) s = "-" + s;
    return s;
}


function formata_numero(in_valor, in_decimais) {
var sinal, valor, decimais, fator, i, j, n, c, s, s_valor, s_int, s_dec, s_resp, achou;
	valor=in_valor;
	decimais=in_decimais;
	if (trim("" + valor)=="") return "";
	if (isNaN(decimais)) decimais=0; else decimais=parseInt(decimais);
 /* Retira formatação e mantém apenas o separador decimal, se houver, no formato inglês. */
	valor = converte_numero(valor);
	if (valor<0) sinal=-1; else sinal=1;
	valor=Math.abs(valor);
	if (isNaN(valor)) return "";
/*  Define número de casas decimais */
	fator=1;
	for (i=1; i<=decimais; i++) fator=fator*10;
	n = Math.round(valor*fator)/fator;
	s_valor = "" + n;
/*  Separa a parte inteira e decimal */
	s_int="";
	s_dec="";
	achou=false;
	for (i=0; i < s_valor.length; i++) {
		c=s_valor.charAt(i);
		if (c==".") {
			achou=true;
			}
		else {
			if (!achou) s_int=s_int+c; else s_dec=s_dec+c;
			}
		}
	
/*  Formata parte decimal */
	while (s_dec.length < decimais) s_dec=s_dec+"0";
	
/*  Formata parte inteira */
	s = "";
	j = 0;
	for (i=s_int.length; i>=0; i--) {
		s = s_int.charAt(i) + s;
		if (((j%3)==0)&&(i!=s_int.length)&&(i!=0)) s = "." + s;
		j=j+1;
		}
	s_int=s;

/*  Monta número formatado final */	
	s = s_int;
	if (s_dec!="") {
		s = s + ",";
		s = s + s_dec;
		}
	if (sinal==-1) s = "-" + s;
	return s;	
}

function formata_moeda(valor) {
var s;
	s=formata_numero(valor, 2);
	return s;
}

function formata_perc_RT(valor) {
var s;
	s=formata_numero(valor, 1);
	return s;
}

function formata_perc_desc(valor) {
var s;
	s=formata_numero(valor, 1);
	return s;
}

function formata_perc_2dec(valor) {
	var s;
	s = formata_numero(valor, 2);
	return s;
}

function formata_coeficiente_custo_financ_fornecedor(valor) {
var s;
	s=formata_numero(valor, MAX_DECIMAIS_COEFICIENTE_CUSTO_FINANCEIRO_FORNECEDOR);
	return s;
}

function formata_perc_comissao(valor) {
var s;
	s=formata_numero(valor, 1);
	return s;
}

function formata_perc_markup(valor) {
var s;
	s=formata_numero(valor, 1);
	return s;
}

function formata_inteiro(valor) {
var s;
	s=formata_numero(valor, 0);
	return s;
}

function limita_tamanho(campo, tam_max){
var s;
	if (window.event.keyCode==8) return false;
	s = "" + campo.value;
	if (s.length >= tam_max) {
		window.event.keyCode=0;
		return true;
		}
	return false;
}

function formata_ddmmyyyy_yyyymmdd( data ) {
var s, i, c, separador, dia, mes, ano;
	s = "" + data;
	s = trim(data);
	if (s.length==0) return "";
	separador="";
	dia="";
	mes="";
	ano="";
	for (i=0; i<s.length; i++) {
		c=s.charAt(i);
		if (!isDigit(c)) {
			if (separador=="") separador=c;
			if (dia!="") while (dia.length < 2) dia="0"+dia;
			if (mes!="") while (mes.length < 2) mes="0"+mes;
			}
		else {
			if (dia.length < 2) dia=dia+c;
			else if (mes.length < 2) mes=mes+c;
			else if (ano.length < 4) ano=ano+c;
			}
		}
	if (ano.length==2) {
		if (converte_numero(ano)>=80) ano="19"+ano;
		else ano="20"+ano;
		}
	else if (ano.length==3) {
		if (converte_numero(ano)>=900) ano="1"+ano;
		else ano="2"+ano;
		}
	else if (ano.length==1) {
		ano="200"+ano;
		}
	return ano + separador + mes + separador + dia;
}

function normaliza_codigo(codigo, tamanho_default) {
var s;
	s = trim("" + codigo);
	if (s=="") return "";
	while (s.length < tamanho_default) s="0"+s;
	return s;
}

function normaliza_num_pedido( pedido ) {
var i, c, s, s_num, s_ano, s_filhote, id_pedido;
	id_pedido = ucase(trim("" + pedido));
	if (id_pedido=="") return "";
	s_num = "";
	for (i=0; i<id_pedido.length; i++) {
		if (isDigit(id_pedido.charAt(i))) s_num=s_num+id_pedido.charAt(i); else break;
		}
	if (s_num=="") return "";
	s_ano = "";
	s_filhote = "";
	for (i=0; i<id_pedido.length; i++) {
		c = id_pedido.charAt(i);
		if (isLetra(c)) {
			if (s_ano=="") 
				s_ano=c;
			else 
				if (s_filhote=="") s_filhote=c;
			}
		}
	if (s_ano=="") return "";
	s_num = normaliza_codigo(s_num, TAM_MIN_NUM_PEDIDO);
	s = s_num + s_ano;
	if (s_filhote != "") s = s + COD_SEPARADOR_FILHOTE + s_filhote;
	return s;
}

function normaliza_num_orcamento( orcamento ) {
var i, c, s, s_num, s_ano, id_orcamento;
	id_orcamento = ucase(trim("" + orcamento));
	if (id_orcamento=="") return "";
	s_num = "";
	for (i=0; i<id_orcamento.length; i++) {
		if (isDigit(id_orcamento.charAt(i))) s_num=s_num+id_orcamento.charAt(i); else break;
		}
	if (s_num=="") return "";
	s_ano = "";
	for (i=0; i<id_orcamento.length; i++) {
		c = id_orcamento.charAt(i);
		if (isLetra(c)) {
			if (s_ano=="") 
				s_ano=c;
			else 
				return "";
			}
		}
	if (s_ano=="") s_ano=SUFIXO_ID_ORCAMENTO;
	s_num = normaliza_codigo(s_num, TAM_MIN_NUM_ORCAMENTO);
	s = s_num + s_ano;
	return s;
}

function isNumeroOrcamento(orcamento) {
	var s_orcamento;
	s_orcamento = ucase(trim("" + orcamento));
	if (s_orcamento == "") return false;
	if (s_orcamento.charAt(s_orcamento.length - 1) == SUFIXO_ID_ORCAMENTO) return true;
	return false;
}

function normaliza_produto( codigo_produto ) {
var produto;
	produto = ucase(trim("" + codigo_produto));
	if (produto=="") return "";
/* NORMALIZA COM ZEROS À ESQUERDA SOMENTE SE O CÓDIGO COMEÇA COM NUMÉRICOS */
	if (!isDigit(left(produto,1))) return produto;
	produto=normaliza_codigo(produto, TAM_MIN_PRODUTO);
	return produto;
}

function isEAN ( codigo ) {
var cod;
	cod = trim("" + codigo);
	return (cod.length==13);
}

function retorna_dados_formulario( f ) {
var i, j, c, s, s_aux;
	s = "";
	for (i=0; i < f.elements.length; i++) {
		c = f.elements(i);
		s_aux = c.type;
		s_aux = s_aux.toUpperCase();
		if (s_aux=="SELECT-ONE") {
			s = s + "|" + c.options[c.selectedIndex].value;
			}
		else if (s_aux=="SELECT-MULTIPLE") {
			for (j=0; j < c.options.length; j++) {
				s = s + "|";
				if (c.options[j].selected) s = s + c.options[j].value;
				}
			}
		else if ((s_aux=="RADIO")||(s_aux=="CHECKBOX")) {
			s = s + "|";
			if (c.checked) s = s + c.value;
			}
		else if ((s_aux=="PASSWORD")||(s_aux=="TEXT")||(s_aux=="TEXTAREA")) {
			s = s + "|" + c.value;
			}
		}
	return s;
}

function normaliza_lista_pedidos( lista ) {
var v, s, s_lista, i;
var v_aux = new Array();
var s_quebra_linha = KEY_RETURN;
	s_lista = trim("" + lista);
	if (s_lista.indexOf(KEY_CRLF) != -1) {
		s_quebra_linha = KEY_CRLF;
		s_lista = substitui_caracteres(s_lista, KEY_LINEFEED, "");
		v = s_lista.split(KEY_RETURN);
	}
	else if (s_lista.indexOf(KEY_LFCR) != -1) {
		s_quebra_linha = KEY_LFCR;
		s_lista = substitui_caracteres(s_lista, KEY_LINEFEED, "");
		v = s_lista.split(KEY_RETURN);
	}
	else if ((s_lista.indexOf(KEY_RETURN) == -1) && (s_lista.indexOf(KEY_LINEFEED) != -1)) {
		// QUEBRA DE LINHA É FEITA APENAS PELO 'LF'
		s_quebra_linha = KEY_LINEFEED;
		v = s_lista.split(KEY_LINEFEED);
	}
	else if ((s_lista.indexOf(KEY_RETURN) != -1) && (s_lista.indexOf(KEY_LINEFEED) == -1)) {
		// QUEBRA DE LINHA É FEITA APENAS PELO 'CR'
		s_quebra_linha = KEY_RETURN;
		v = s_lista.split(KEY_RETURN);
	}
	else {
		// QUEBRA DE LINHA USA CARACTERES DESCONHECIDOS
		return s_lista;
	}

	v_aux[0]="";
	for (i = 0; i < v.length; i++) {
		if (trim("" + v[i]) != "") {
			s = normaliza_num_pedido(v[i]);
			if (s != "") v[i] = s;
			if (trim("" + v[i]) != "") {
				if (v_aux[v_aux.length - 1] == "") {
					v_aux[v_aux.length - 1] = v[i];
				}
				else {
					v_aux[v_aux.length] = v[i];
				}
			}
		}
	}
	s_lista = v_aux.join(s_quebra_linha);
	return s_lista;
}

function normaliza_lista_lojas( lista ) {
var v, s, s_lista, i, j;
var v_aux = new Array();
var l;
var s_quebra_linha = KEY_RETURN;
	s_lista = trim("" + lista);
	if (s_lista.indexOf(KEY_CRLF) != -1) {
		s_quebra_linha = KEY_CRLF;
		s_lista = substitui_caracteres(s_lista, KEY_LINEFEED, "");
		v = s_lista.split(KEY_RETURN);
	}
	else if (s_lista.indexOf(KEY_LFCR) != -1) {
		s_quebra_linha = KEY_LFCR;
		s_lista = substitui_caracteres(s_lista, KEY_LINEFEED, "");
		v = s_lista.split(KEY_RETURN);
	}
	else if ((s_lista.indexOf(KEY_RETURN) == -1) && (s_lista.indexOf(KEY_LINEFEED) != -1)) {
		// QUEBRA DE LINHA É FEITA APENAS PELO 'LF'
		s_quebra_linha = KEY_LINEFEED;
		v = s_lista.split(KEY_LINEFEED);
	}
	else if ((s_lista.indexOf(KEY_RETURN) != -1) && (s_lista.indexOf(KEY_LINEFEED) == -1)) {
		// QUEBRA DE LINHA É FEITA APENAS PELO 'CR'
		s_quebra_linha = KEY_RETURN;
		v = s_lista.split(KEY_RETURN);
	}
	else {
		// QUEBRA DE LINHA USA CARACTERES DESCONHECIDOS
		return s_lista;
	}
	
	v_aux[0]="";
	for (i = 0; i < v.length; i++) {
		if (trim("" + v[i]) != "") {
			s = v[i];
			l = s.split("-");
			for (j = 0; j < l.length; j++) {
				if (trim("" + l[j]) != "") {
					s = normaliza_codigo(l[j], TAM_MIN_LOJA);
					if (s != "") l[j] = s;
				}
			}
			v[i] = l.join("-");
			if (trim("" + v[i]) != "") {
				if (v_aux[v_aux.length - 1] == "") {
					v_aux[v_aux.length - 1] = v[i];
				}
				else {
					v_aux[v_aux.length] = v[i];
				}
			}
		}
	}
	s_lista = v_aux.join(s_quebra_linha);
	return s_lista;
}

function normaliza_lista_fabricantes( lista ) {
var v, s, s_lista, i, j;
var v_aux = new Array();
var l;
var s_quebra_linha = KEY_RETURN;
	s_lista = trim("" + lista);
	if (s_lista.indexOf(KEY_CRLF) != -1) {
		s_quebra_linha = KEY_CRLF;
		s_lista = substitui_caracteres(s_lista, KEY_LINEFEED, "");
		v = s_lista.split(KEY_RETURN);
	}
	else if (s_lista.indexOf(KEY_LFCR) != -1) {
		s_quebra_linha = KEY_LFCR;
		s_lista = substitui_caracteres(s_lista, KEY_LINEFEED, "");
		v = s_lista.split(KEY_RETURN);
	}
	else if ((s_lista.indexOf(KEY_RETURN) == -1) && (s_lista.indexOf(KEY_LINEFEED) != -1)) {
		// QUEBRA DE LINHA É FEITA APENAS PELO 'LF'
		s_quebra_linha = KEY_LINEFEED;
		v = s_lista.split(KEY_LINEFEED);
	}
	else if ((s_lista.indexOf(KEY_RETURN) != -1) && (s_lista.indexOf(KEY_LINEFEED) == -1)) {
		// QUEBRA DE LINHA É FEITA APENAS PELO 'CR'
		s_quebra_linha = KEY_RETURN;
		v = s_lista.split(KEY_RETURN);
	}
	else {
		// QUEBRA DE LINHA USA CARACTERES DESCONHECIDOS
		return s_lista;
	}
	
	v_aux[0]="";
	for (i = 0; i < v.length; i++) {
		if (trim("" + v[i]) != "") {
			s = v[i];
			l = s.split("-");
			for (j = 0; j < l.length; j++) {
				if (trim("" + l[j]) != "") {
					s = normaliza_codigo(l[j], TAM_MIN_FABRICANTE);
					if (s != "") l[j] = s;
				}
			}
			v[i] = l.join("-");
			if (trim("" + v[i]) != "") {
				if (v_aux[v_aux.length - 1] == "") {
					v_aux[v_aux.length - 1] = v[i];
				}
				else {
					v_aux[v_aux.length] = v[i];
				}
			}
		}
	}
	s_lista = v_aux.join(s_quebra_linha);
	return s_lista;
}

function normaliza_lista_cnpj_cpf(lista) {
	var v, s, s_lista, i, j;
	var v_aux = new Array();
	var l;
	var s_quebra_linha = KEY_RETURN;
	s_lista = trim("" + lista);
	if (s_lista.indexOf(KEY_CRLF) != -1) {
		s_quebra_linha = KEY_CRLF;
		s_lista = substitui_caracteres(s_lista, KEY_LINEFEED, "");
		v = s_lista.split(KEY_RETURN);
	}
	else if (s_lista.indexOf(KEY_LFCR) != -1) {
		s_quebra_linha = KEY_LFCR;
		s_lista = substitui_caracteres(s_lista, KEY_LINEFEED, "");
		v = s_lista.split(KEY_RETURN);
	}
	else if ((s_lista.indexOf(KEY_RETURN) == -1) && (s_lista.indexOf(KEY_LINEFEED) != -1)) {
		// QUEBRA DE LINHA É FEITA APENAS PELO 'LF'
		s_quebra_linha = KEY_LINEFEED;
		v = s_lista.split(KEY_LINEFEED);
	}
	else if ((s_lista.indexOf(KEY_RETURN) != -1) && (s_lista.indexOf(KEY_LINEFEED) == -1)) {
		// QUEBRA DE LINHA É FEITA APENAS PELO 'CR'
		s_quebra_linha = KEY_RETURN;
		v = s_lista.split(KEY_RETURN);
	}
	else {
		// QUEBRA DE LINHA USA CARACTERES DESCONHECIDOS
		return s_lista;
	}

	v_aux[0] = "";
	for (i = 0; i < v.length; i++) {
		if (trim("" + v[i]) != "") {
			s = v[i];
			l = s.split("-");
			for (j = 0; j < l.length; j++) {
				if (trim("" + l[j]) != "") {
					s = cnpj_cpf_formata(l[j]);
					if (s != "") l[j] = s;
				}
			}
			v[i] = l.join("-");
			if (trim("" + v[i]) != "") {
				if (v_aux[v_aux.length - 1] == "") {
					v_aux[v_aux.length - 1] = v[i];
				}
				else {
					v_aux[v_aux.length] = v[i];
				}
			}
		}
	}
	s_lista = v_aux.join(s_quebra_linha);
	return s_lista;
}

/* IMPORTANTE
   ==========
   Os parâmetros da função devem ser campos do formulário,
   pois a função irá referenciar a propriedade "value" do parâmetro.
*/
function consiste_periodo( c_dt_i, c_dt_f ) {
var s_de, s_ate;
	
	if (trim(c_dt_i.value)!="") {
		if (!isDate(c_dt_i)) {
			alert("Data inválida!!");
			c_dt_i.focus();
			return false;
			}
		}

	if (trim(c_dt_f.value)!="") {
		if (!isDate(c_dt_f)) {
			alert("Data inválida!!");
			c_dt_f.focus();
			return false;
			}
		}
				
	s_de = trim(c_dt_i.value);
	s_ate = trim(c_dt_f.value);
	if ((s_de!="")&&(s_ate!="")) {
		s_de=retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(s_de));
		s_ate=retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(s_ate));
		if (s_de > s_ate) {
			alert("Data de término é menor que a data de início!!");
			c_dt_f.focus();
			return false;
			}
		}
	
	return true;
}

function tem_digito(texto) {
var i, achou;
	texto = trim("" + texto);
	achou = false;
	for (i=0; i<texto.length; i++) {
		if (isDigit(texto.charAt(i))) {
			achou = true;
			break;
			}
		}
	if (achou) return true;
	return false;
}

function tem_letra(texto) {
	var i, achou;
	texto = trim("" + texto);
	achou = false;
	for (i = 0; i < texto.length; i++) {
		if (isLetra(texto.charAt(i))) {
			achou = true;
			break;
		}
	}
	if (achou) return true;
	return false;
}

function iniciais_em_maiusculas(texto){
var palavras_minusculas = "|A|AS|E|O|OS|UM|UNS|UMA|UMAS|DA|DAS|DE|DO|DOS|EM|NA|NAS|NO|NOS|COM|SEM|POR|PELO|PARA|PRA|P/|S/|C/|TEM|OU|E/OU|";
var letra, palavra, frase, s, i, i_max;
	frase = "";
	palavra = "";
	i_max = texto.length;
	for (i=0; i<i_max; i++) {
		letra = texto.charAt(i);
		palavra = palavra + letra;
		if ((letra==" ")||(i==(i_max-1))) {
			s = "|" + ucase(trim(palavra)) + "|";
			if ((palavras_minusculas.indexOf(s)!=-1)&&(frase!="")) {
				palavra = lcase(palavra);
				}
			else {
			//	SE POSSUI DÍGITOS, ENTÃO É ALGUM TIPO DE CÓDIGO
				if (!tem_digito(palavra)){
					palavra = ucase(left(palavra,1)) + lcase(palavra.substr(1));
					}
				}
			frase = frase + palavra;
			palavra = "";
			}
		}
	return frase;
}

function configura_painel_logon() {
	// As versões mais novas não informam corretamente o número da versão na string do user agent
	// Além disso, quando se ativa o modo de exibição de compatibilidade, o número da versão não corresponde ao real
	return;
	
	var ver = getInternetExplorerVersion();
	try {
		if (ver >= 9.0) return;
		window.moveTo(Math.floor((screen.availWidth-600)/2), Math.floor((screen.availHeight-540)/2));
		window.resizeTo( 600, 540 );
		}
	catch (e) {
	 // NOP
		}
}

function configura_painel() {
	var ver = getInternetExplorerVersion();
	try {
		if (ver >= 9.0) return;
		window.moveTo(0,0);
		window.resizeTo(Math.floor(screen.availWidth), Math.floor(screen.availHeight));
		}
	catch (e) {
	 // NOP
		}
}

function excel_converte_numeracao_digito_para_letra(numeracao_digito) {
var TotalLetrasAlfabeto=26;
var strResp, intQuoc, intNumero, intResto;
	strResp='';
	intNumero=(numeracao_digito-1);
	intQuoc=IntPart((numeracao_digito-1)/TotalLetrasAlfabeto);
	intResto=numeracao_digito-(intQuoc*TotalLetrasAlfabeto);
	if (intQuoc > TotalLetrasAlfabeto) return '';
	if (intQuoc > 0) strResp=String.fromCharCode(65-1+intQuoc);
	strResp+=String.fromCharCode(65-1+intResto);
	return strResp;
}

function calcula_valor_presente(vl_valor_futuro, perc_taxa, n_periodos) {
var vl_valor_presente;
//	PV = FV / (1+i)^n
	vl_valor_presente = vl_valor_futuro / Math.pow((1+perc_taxa), n_periodos);
	return vl_valor_presente;
}

function calcula_total_RA_liquido(percentual_desagio_RA_liquida, vl_total_RA) {
var vl_total_RA_liquido;
	this.blnAplicouDesagioRA = true;
	if (vl_total_RA<0) {
		vl_total_RA_liquido=0;
		}
	else {
		vl_total_RA_liquido = vl_total_RA - (percentual_desagio_RA_liquida / 100) * vl_total_RA;
		}
	this.vl_total_RA_liquido = converte_numero(formata_moeda(vl_total_RA_liquido));
}

function maxLength(campo,maxChars)
{
	if(campo.value.length >= maxChars) {
		event.returnValue=false;
		return false;
		}
}

function maxLengthPaste(campo,maxChars)
{
	if((campo.value.length + window.clipboardData.getData("Text").length) > maxChars) {
		event.returnValue=false;
		return false;
		}
	event.returnValue=true;
}

function autoTab() {
var s, objShell;
	try {
		objShell = new ActiveXObject("wscript.shell");
		objShell.SendKeys("{tab}");
	}
	catch (e) {
		// NOP
	}
}

function descricaoCustoFinancFornecTipoParcelamento(strCodigoTipoParcelamento) {
var strResp;
	strResp = "";
	if (strCodigoTipoParcelamento == COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__COM_ENTRADA) {
		strResp = "Com Entrada";
		}
	else if (strCodigoTipoParcelamento == COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__SEM_ENTRADA) {
		strResp = "Sem Entrada";
		}
	else if (strCodigoTipoParcelamento == COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__A_VISTA) {
		strResp = "À Vista";
		}
	return strResp;
}

function isSelectAllCheckBoxesKeywordOk(keyword) {
	if (right(keyword.toUpperCase(), "*ALL*".length) == "*ALL*") return true;
	if (right(keyword.toUpperCase(), "*TODOS*".length) == "*TODOS*") return true;
	return false;
}

function getInternetExplorerVersion()
// Returns the version of Internet Explorer or a -1
// (indicating the use of another browser).
{
	var rv = -1; // Return value assumes failure.
	if (navigator.appName == 'Microsoft Internet Explorer') {
		var ua = navigator.userAgent;
		var re = new RegExp("MSIE ([0-9]{1,}[\.0-9]{0,})");
		if (re.exec(ua) != null)
			rv = parseFloat(RegExp.$1);
	}
	return rv;
}

function isVersaoNavegadorOk() {
	if (navigator.userAgent.toUpperCase().indexOf(" EDGE/") > -1) return true;

	if (navigator.userAgent.toUpperCase().indexOf("CHROME") > -1) return false;
	if (navigator.userAgent.toUpperCase().indexOf("SAFARI") > -1) return false;
	if (navigator.userAgent.toUpperCase().indexOf("FIREFOX") > -1) return false;
	
	if (navigator.appName.toUpperCase().indexOf("MICROSOFT") > -1) return true;
	if (navigator.appName.toUpperCase().indexOf("INTERNET") > -1) return true;
	if (navigator.appName.toUpperCase().indexOf("EXPLORER") > -1) return true;
	if (navigator.appName.toUpperCase().indexOf("MSIE") > -1) return true;
	
	if (navigator.userAgent.toUpperCase().indexOf("MSIE") > -1) return true;
	if ((navigator.userAgent.toUpperCase().indexOf(".NET") > -1) && (navigator.userAgent.toUpperCase().indexOf(" RV:") > -1)) return true;
	return false;
}

function isPlacaVeiculoOk(numeroPlaca) {
var i, c, letras, numeros;

	if (numeroPlaca == null) return false;
	if (trim(numeroPlaca) == "") return false;

	letras = "";
	numeros = "";
	for (i = 0; i < numeroPlaca.length; i++) {
		c = numeroPlaca.charAt(i);

		if (c == " ") {
		//  O ESPAÇO EM BRANCO APARECEU EM POSIÇÃO INESPERADA?
			if (letras.length != 3) return false;
			if (numeros.length > 0) return false;
		}
		else if (isLetra(c)) {
		//  APARECEU UMA LETRA DEPOIS DE JÁ TER INICIADO A PARTE DOS DÍGITOS?
			if (numeros.length > 0) return false;
			letras += c;
		}
		else if (isDigit(c)) {
		//  APARECEU UM DÍGITO EM POSIÇÃO INESPERADA?
			if (letras.length != 3) return false;
			numeros += c;
		}
		else {
		//  CARACTER INVÁLIDO!
			return false;
		}
	}

	if (letras.length != 3) return false;
	if (numeros.length != 4) return false;

	return true;
}

function filtra_digitacao_wms_deposito_zona_codigos(listaCodigos) {
var letra;
	listaCodigos = "" + listaCodigos;
	listaCodigos = ucase(listaCodigos);
	letra = String.fromCharCode(window.event.keyCode);
	letra = letra.toUpperCase();
	if (listaCodigos.indexOf(letra) == -1) {
		window.event.keyCode = 0;
	}
	else {
	//  Se for letra minúscula, converte p/ maiuscula
		if ((window.event.keyCode > 96) && (window.event.keyCode < 123)) window.event.keyCode = window.event.keyCode - 32;
	}
}

function wms_deposito_codigo_ok(listaCodigos, codigo) {
	listaCodigos = "" + listaCodigos;
	listaCodigos = ucase(listaCodigos);
	codigo = "" + codigo;
	codigo = ucase(codigo);
	if ((codigo != "") && (codigo.indexOf("|") == -1)) codigo = "|" + codigo + "|";
	if (listaCodigos.indexOf(codigo) > -1) return true;
	return false;
}

function synchronous_ajax(url, postData) {
	if (window.XMLHttpRequest) {
		AJAX = new XMLHttpRequest();
	}
	else {
		AJAX = new ActiveXObject("Microsoft.XMLHTTP");
	}
	if (AJAX) {
	//	Prevents server from using a cached file
		if (url.toString().indexOf("?") > -1) url += "&"; else url += "?";
		url += "anticache=" + Math.random() + Math.random();
		AJAX.open("POST", url, false);
		AJAX.setRequestHeader("Content-type", "application/x-www-form-urlencoded");
		AJAX.send(postData);
		return AJAX.responseText;
	}
	else {
		return false;
	}
}

function formToString(filledForm) {
	// Without hidden fields
	if ((typeof (filledForm) === "undefined") || (filledForm === null)) return;
	formObject = new Object
	filledForm.find("input, select, textarea").each(function () {
		if (this.id) {
			elem = $(this);
			if (elem.attr("type") == 'hidden') {
				// NOP
			}
			else if (elem.attr("type") == 'checkbox' || elem.attr("type") == 'radio') {
				formObject[this.id] = elem.prop("checked");
			} else {
				formObject[this.id] = elem.val();
			}
		}
	});
	formString = JSON.stringify(formObject);
	return formString;
}

function formToStringAll(filledForm) {
	// Including hidden fields
	if ((typeof (filledForm) === "undefined") || (filledForm === null)) return;
	formObject = new Object
	filledForm.find("input, select, textarea").each(function () {
		if (this.id) {
			elem = $(this);
			if (elem.attr("type") == 'checkbox' || elem.attr("type") == 'radio') {
				formObject[this.id] = elem.prop("checked");
			} else {
				formObject[this.id] = elem.val();
			}
		}
	});
	formString = JSON.stringify(formObject);
	return formString;
}

function stringToForm(formString, unfilledForm) {
	if ((typeof (unfilledForm) === "undefined") || (unfilledForm === null)) return;
	if ((typeof (formString) === "undefined") || (formString === null)) return;
	if (formString == "") return;

	formObject = JSON.parse(formString);
	unfilledForm.find("input, select, textarea").each(function () {
		if (this.id) {
			id = this.id;
			if ((typeof (formObject[id]) !== "undefined") && (formObject[id] !== null)) {
				elem = $(this);
				if (elem.attr("type") == "checkbox" || elem.attr("type") == "radio") {
					elem.prop("checked", formObject[id]);
				} else {
					elem.val(formObject[id]);
				}
			}
		}
	});
}

function converte_cst_nfe_fabricante_para_entrada_estoque(cst_nfe_fabricante) {
	var cst, s_resp;
	cst = "" + cst_nfe_fabricante;

	if (cst == "000") {
		s_resp = "000";
	}
	else if (cst == "010") {
		s_resp = "060";
	}
	else if (cst == "100") {
		s_resp = "200";
	}
	else if (cst == "110") {
		s_resp = "260";
	}
	else if (cst == "200") {
		s_resp = "200";
	}
	else if (cst == "300") {
		s_resp = "200";
	}
	else if (cst == "400") {
		s_resp = "000";
	}
	else if (cst == "441") {
		s_resp = "000";
	}
	else if (cst == "500") {
		s_resp = "000";
	}
	else if (cst == "600") {
		s_resp = "200";
	}
	else if (cst == "700") {
		s_resp = "200";
	}
	else if (cst == "800") {
		s_resp = "000";
	}
	else if (cst == "160") {
		s_resp = "260";
	}
	else if (cst == "141") {
		s_resp = "241";
	}
	else {
		s_resp = "";
	}

	return s_resp;
}

function limpaMultiplosCampos() {
	var c;
	for (var i = 0; i < arguments.length; i++) {
		c = arguments[i];
		if (c.type && c.type.toLowerCase() === 'checkbox') {
			c.checked = false;
		}
		else {
			c.value = "";
		}
	}
}
