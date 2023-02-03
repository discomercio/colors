<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp"        -->
<!-- #include file = "../global/Global.asp"        -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =============================================================
'	  P E S Q U I S A D E I N D I C A D O R E S F I L T R O . A S P
'     =============================================================
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

	Const COD_PESQUISAR_POR_UF_LOCALIDADE = "POR_UF_LOCALIDADE"
	Const COD_PESQUISAR_POR_BAIRRO = "POR_BAIRRO"
	Const COD_PESQUISAR_POR_CEP = "POR_CEP"
	Const COD_PESQUISAR_POR_NOME = "POR_NOME"
	Const COD_PESQUISAR_POR_CPF_CNPJ = "POR_CPF_CNPJ"
	Const COD_PESQUISAR_ASSOCIADOS_AO_VENDEDOR = "POR_ASSOCIADOS_AO_VENDEDOR"
	
	dim usuario
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

	dim s

'	CONECTA COM O BANCO DE DADOS
	dim cn
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	if Not operacao_permitida(OP_CEN_PESQUISA_INDICADORES, s_lista_operacoes_permitidas) then 
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
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
   
 
<%=DOCTYPE_LEGADO %>


<head>
	<title>CENTRAL</title>

	
<script src="<%=URL_FILE__JQUERY%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_I18N%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__AJAX_JS%>" language="JavaScript" type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
var objAjaxPesqCep;
var objAjaxListaIndicadores;
var objAjaxListaVendedores;
var strListaIndicadoresUltimaLojaPesquisada="";
var strListaVendedoresUltimaLojaPesquisada="";

function LimpaListasDependentesDaLoja() {
	LimpaListaIndicadores();
	LimpaListaVendedores();
}

function LimpaListaIndicadores() {
var f, oOption;
	f=fFILTRO;
	$("#c_indicador").empty();

//  Cria um item vazio
	oOption=document.createElement("OPTION");
	f.c_indicador.options.add(oOption);
	oOption.innerText="";
	oOption.value="";
}

function LimpaListaVendedores() {
var f, oOption;
	f=fFILTRO;
	$("#c_vendedor").empty();

//  Cria um item vazio
	oOption=document.createElement("OPTION");
	f.c_vendedor.options.add(oOption);
	oOption.innerText="";
	oOption.value="";
}
function LimpaListaCidadeBairro() {
    var f, oOption;
    f = fFILTRO;
    $("#cidade_bairro").empty();

}

function LimpaListaLocalidades() {
var f, oOption;
	f=fFILTRO;
	$("#c_escolher_loc").empty();

}
function LimpaListaLocalidadesSelecionadas() {
    var f, oOption;
    f = fFILTRO;
    $("#c_localidade_pesq").empty();
}

function LimpaListaBairros() {
    var f, oOption;
    f = fFILTRO;
    $("#c_bairro").empty();

}

function LimpaListaBairrosSelecionados() {
    var f, oOption;
    f = fFILTRO;
    $("#bairro_pesq").empty();

}

function TrataRespostaAjaxPesquisaLocalidades() {
var f, i, strAux, strResp, xmlDoc, oOption, oNodes;
	f=fFILTRO;
	if (objAjaxPesqLocalidades.readyState==AJAX_REQUEST_IS_COMPLETE) {
		strResp=objAjaxPesqLocalidades.responseText;
		if (strResp=="") {
			window.status="Concluído";
			divMsgAguardeObtendoDados.style.visibility = "hidden";
			alert("Nenhuma localidade encontrada!!");
			return;
			}
		
		if (strResp!="") {
			try 
				{
				xmlDoc=objAjaxPesqLocalidades.responseXML.documentElement;
				for (i=0; i < xmlDoc.getElementsByTagName("registro").length; i++) {
					oOption=document.createElement("OPTION");
					f.c_escolher_loc.options.add(oOption);
					
					oNodes=xmlDoc.getElementsByTagName("localidade")[i];
					if (oNodes.childNodes.length > 0) strAux=oNodes.childNodes[0].nodeValue; else strAux="";
					if (strAux==null) strAux="";
					oOption.innerText=strAux;
					oOption.value=strAux;
					}
				}
			catch (e)
				{
				alert("Falha na consulta!!");
				}
			}
		window.status="Concluído";
		divMsgAguardeObtendoDados.style.visibility = "hidden";
		f.c_escolher_loc.focus();
		}
}
function TrataRespostaAjaxPesquisaCidadeBairro() {
    var f, i, strAux, strResp, xmlDoc, oOption, oNodes;
    f = fFILTRO;
    if (objAjaxPesqLocalidades.readyState == AJAX_REQUEST_IS_COMPLETE) {
        strResp = objAjaxPesqLocalidades.responseText;
        if (strResp == "") {
            window.status = "Concluído";
            divMsgAguardeObtendoDados.style.visibility = "hidden";
            alert("Nenhuma localidade encontrada!!");
            return;
        }

        if (strResp != "") {
            try {
                xmlDoc = objAjaxPesqLocalidades.responseXML.documentElement;
                for (i = 0; i < xmlDoc.getElementsByTagName("registro").length; i++) {
                    oOption = document.createElement("OPTION");
                    f.cidade_bairro.options.add(oOption);

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
        divMsgAguardeObtendoDados.style.visibility = "hidden";
        f.cidade_bairro.focus();
    }
}
function TrataRespostaAjaxPesquisaBairros() {
    var f, i, strAux, strResp, xmlDoc, oOption, oNodes;
    f = fFILTRO;
    if (objAjaxPesqLocalidades.readyState == AJAX_REQUEST_IS_COMPLETE) {
        strResp = objAjaxPesqLocalidades.responseText;
        if (strResp == "") {
            window.status = "Concluído";
            divMsgAguardeObtendoDados.style.visibility = "hidden";
            alert("Nenhum bairro encontrado!!");
            return;
        }

        if (strResp != "") {
            try {
                xmlDoc = objAjaxPesqLocalidades.responseXML.documentElement;
                for (i = 0; i < xmlDoc.getElementsByTagName("registro").length; i++) {
                    oOption = document.createElement("OPTION");
                    f.c_bairro.options.add(oOption);

                    oNodes = xmlDoc.getElementsByTagName("bairro")[i];
                    if (oNodes.childNodes.length > 0) strAux = oNodes.childNodes[0].nodeValue; else strAux = "";
                    if (strAux == null) strAux = "";
                    oOption.innerText = strAux;
                    oOption.value = strAux;
                }
            }
            catch (e) {
                alert("Falha na consulta!!!");
            }
        }
        window.status = "Concluído";
        divMsgAguardeObtendoDados.style.visibility = "hidden";
        f.c_bairro.focus();
    }
}
function CarregaBairro() {
    var f, strUrl, strCidade;
    f = fFILTRO;
    objAjaxPesqLocalidades = GetXmlHttpObject();
    if (objAjaxPesqLocalidades == null) {
        alert("O browser NÃO possui suporte ao AJAX!!");
        return;
    }

    //  Limpa lista de bairros
    LimpaListaBairros();
    LimpaListaBairrosSelecionados();

    strCidade = trim(f.cidade_bairro.value);
    if (strCidade == "") {
        return;
    }

    window.status = "Aguarde, pesquisando os bairros de " + f.cidade_bairro.value + " ...";
    divMsgAguardeObtendoDados.style.visibility = "";

    strUrl = "../Global/AjaxIndicadoresLocalidadeBairro.asp";
    strUrl = strUrl + "?loja=" + f.c_loja.value;
    strUrl = strUrl + "&cidade=" + f.cidade_bairro.value + "&retira_acentuacao=S";
    //  Prevents server from using a cached file
    strUrl = strUrl + "&sid=" + Math.random() + Math.random();
    objAjaxPesqLocalidades.onreadystatechange = TrataRespostaAjaxPesquisaBairros;
    objAjaxPesqLocalidades.open("GET", strUrl, true);
    objAjaxPesqLocalidades.send(null);
}
function CarregaCidadeBairro() {
    var f, strUrl, strUF;
    f = fFILTRO;
    objAjaxPesqLocalidades = GetXmlHttpObject();
    if (objAjaxPesqLocalidades == null) {
        alert("O browser NÃO possui suporte ao AJAX!!");
        return;
    }

    //  Limpa lista de localidades
    LimpaListaCidadeBairro();

    strUF = trim(f.uf_bairro.value);
    if (strUF == "") {
        return;
    }

    window.status = "Aguarde, pesquisando as localidades de " + f.uf_bairro.value + " ...";
    divMsgAguardeObtendoDados.style.visibility = "";

    strUrl = "../Global/AjaxCepLocalidadesPesqBD.asp";
    strUrl = strUrl + "?uf=" + f.uf_bairro.value;
    //  Prevents server from using a cached file
    strUrl = strUrl + "&sid=" + Math.random() + Math.random();
    strUrl = strUrl + "&retira_acentuacao=S";
    objAjaxPesqLocalidades.onreadystatechange = TrataRespostaAjaxPesquisaCidadeBairro;
    objAjaxPesqLocalidades.open("GET", strUrl, true);
    objAjaxPesqLocalidades.send(null);
}
function CarregaLocalidades() {
var f, strUrl, strUF;
	f=fFILTRO;
	objAjaxPesqLocalidades=GetXmlHttpObject();
	if (objAjaxPesqLocalidades==null) {
		alert("O browser NÃO possui suporte ao AJAX!!");
		return;
		}

//  Limpa lista de localidades
		LimpaListaLocalidades();
		LimpaListaLocalidadesSelecionadas();
		
	strUF=trim(f.c_uf_pesq.value);
	if (strUF=="") {
		return;
		}
		
	window.status="Aguarde, pesquisando as localidades de " + f.c_uf_pesq.value + " ...";
	divMsgAguardeObtendoDados.style.visibility = "";
	
	strUrl="../Global/AjaxCepLocalidadesPesqBD.asp";
	strUrl=strUrl+"?uf="+f.c_uf_pesq.value;
//  Prevents server from using a cached file
	strUrl = strUrl + "&sid=" + Math.random() + Math.random();
	strUrl = strUrl + "&retira_acentuacao=S";
	objAjaxPesqLocalidades.onreadystatechange=TrataRespostaAjaxPesquisaLocalidades;
	objAjaxPesqLocalidades.open("GET",strUrl,true);
	objAjaxPesqLocalidades.send(null);
}

function TrataRespostaAjaxListaIndicadores() {
var f, i, strApelido, strNome, strResp, xmlDoc, oOption, oNodes;
	f=fFILTRO;
	if (objAjaxListaIndicadores.readyState==AJAX_REQUEST_IS_COMPLETE) {
		strResp=objAjaxListaIndicadores.responseText;
		if (strResp=="") {
			window.status="Concluído";
			divMsgAguardeObtendoDados.style.visibility="hidden";
			return;
			}

		if (strResp!="") {
			try 
				{
				xmlDoc=objAjaxListaIndicadores.responseXML.documentElement;
				for (i=0; i < xmlDoc.getElementsByTagName("registro").length; i++) {
					oOption=document.createElement("OPTION");
					f.c_indicador.options.add(oOption);
					
					oNodes=xmlDoc.getElementsByTagName("apelido")[i];
					if (oNodes.childNodes.length > 0) strApelido=oNodes.childNodes[0].nodeValue; else strApelido="";
					if (strApelido==null) strApelido="";
					oOption.value=strApelido;

					oNodes=xmlDoc.getElementsByTagName("razao_social_nome")[i];
					if (oNodes.childNodes.length > 0) strNome=oNodes.childNodes[0].nodeValue; else strNome="";
					if (strNome==null) strNome="";

					oOption.value=strApelido;
					oOption.innerText=strApelido + " - " + strNome;
					}
				}
			catch (e)
				{
				alert("Falha na consulta de indicadores!!" + "\n" + e.description);
				}
			}
		window.status="Concluído";
		divMsgAguardeObtendoDados.style.visibility="hidden";
		}
}

function CarregaListaIndicadores(strLoja) {
var f, strUrl;
	f=fFILTRO;
	objAjaxListaIndicadores=GetXmlHttpObject();
	if (objAjaxListaIndicadores==null) {
		alert("O browser NÃO possui suporte ao AJAX!!");
		return;
		}

//  Limpa lista de Indicadores
	LimpaListaIndicadores();
		
	if (strLoja=="") return;
		
	window.status="Aguarde, pesquisando os indicadores da loja " + strLoja + " ...";
	divMsgAguardeObtendoDados.style.visibility="";
		
	strListaIndicadoresUltimaLojaPesquisada=strLoja;
	strUrl="../Global/AjaxListaIndicadoresLojaPesqBD.asp";
	strUrl=strUrl+"?loja="+strLoja;
//  Prevents server from using a cached file
	strUrl=strUrl+"&sid="+Math.random()+Math.random();
	objAjaxListaIndicadores.onreadystatechange=TrataRespostaAjaxListaIndicadores;
	objAjaxListaIndicadores.open("GET",strUrl,true);
	objAjaxListaIndicadores.send(null);
}

function TrataRespostaAjaxListaVendedores() {
var f, i, strApelido, strNome, strResp, xmlDoc, oOption, oNodes;
	f=fFILTRO;
	if (objAjaxListaVendedores.readyState==AJAX_REQUEST_IS_COMPLETE) {
		strResp=objAjaxListaVendedores.responseText;
		if (strResp=="") {
			window.status="Concluído";
			divMsgAguardeObtendoDados.style.visibility="hidden";
			return;
			}

		if (strResp!="") {
			try 
				{
				xmlDoc=objAjaxListaVendedores.responseXML.documentElement;
				for (i=0; i < xmlDoc.getElementsByTagName("registro").length; i++) {
					oOption=document.createElement("OPTION");
					f.c_vendedor.options.add(oOption);
					
					oNodes=xmlDoc.getElementsByTagName("usuario")[i];
					if (oNodes.childNodes.length > 0) strApelido=oNodes.childNodes[0].nodeValue; else strApelido="";
					if (strApelido==null) strApelido="";
					oOption.value=strApelido;

					oNodes=xmlDoc.getElementsByTagName("nome")[i];
					if (oNodes.childNodes.length > 0) strNome=oNodes.childNodes[0].nodeValue; else strNome="";
					if (strNome==null) strNome="";

					oOption.value=strApelido;
					oOption.innerText=strApelido + " - " + strNome;
					}
				}
			catch (e)
				{
				alert("Falha na consulta de vendedores!!" + "\n" + e.description);
				}
			}
		window.status="Concluído";
		divMsgAguardeObtendoDados.style.visibility="hidden";
		}
}

function CarregaListaVendedores(strLoja) {
var f, strUrl;
	f=fFILTRO;
	objAjaxListaVendedores=GetXmlHttpObject();
	if (objAjaxListaVendedores==null) {
		alert("O browser NÃO possui suporte ao AJAX!!");
		return;
		}

//  Limpa lista de Vendedores
	LimpaListaVendedores();
		
	if (strLoja=="") return;
		
	window.status="Aguarde, pesquisando os vendedores da loja " + strLoja + " ...";
	divMsgAguardeObtendoDados.style.visibility="";
		
	strListaVendedoresUltimaLojaPesquisada=strLoja;
	strUrl="../Global/AjaxListaVendedoresLojaPesqBD.asp";
	strUrl=strUrl+"?loja="+strLoja;
//  Prevents server from using a cached file
	strUrl=strUrl+"&sid="+Math.random()+Math.random();
	objAjaxListaVendedores.onreadystatechange=TrataRespostaAjaxListaVendedores;
	objAjaxListaVendedores.open("GET",strUrl,true);
	objAjaxListaVendedores.send(null);
}

function fFILTROConfirma( f ) {
var strCep;

// PREENCHER LOJA ?

    //if (f.c_loja.value == "") {
    //    alert("Preencha o campo loja!");
    //    f.c_loja.focus();
    //    return;
    //}

// seleciona as cidades e bairros antes de enviar o formulario
    $("#c_localidade_pesq").children().prop('selected', true);
    $("#bairro_pesq").children().prop('selected', true);
    
//  Nenhum parâmetro de pesquisa fornecido
		if ((!f.rb_pesquisar_por[0].checked) && (!f.rb_pesquisar_por[1].checked) && (!f.rb_pesquisar_por[2].checked) && (!f.rb_pesquisar_por[3].checked) && (!f.rb_pesquisar_por[4].checked) && (!f.rb_pesquisar_por[5].checked)) {
		alert("Forneça algum dos parâmetros de pesquisa!!");
		return;
		}

	if (f.rb_pesquisar_por[0].checked) {
		if (trim(f.c_uf_pesq.value)=="") {
			alert("Selecione uma UF!!");
			return;
			}
		}

    
// Selecionou UF e cidade ?
		if (f.rb_pesquisar_por[1].checked) {
		    if (trim(f.uf_bairro.value) == "") {
		        if (f.cidade_bairro.length == 0 && f.c_bairro.length == 0 && f.bairro_pesq.length == 0) {
		            alert("Selecione uma UF!!");
		            f.uf_bairro.focus();
		            return;
		        }
		    }
		    if ((trim(f.cidade_bairro.value) == "") && (trim(f.uf_bairro.value) != "")) {
		        if (f.c_bairro.length == 0 && f.bairro_pesq.length == 0) {
		            alert("Selecione a cidade de onde quer pesquisar o bairro!!");
		            f.cidade_bairro.focus();
		            return;
		        }
		    }
		    if ((trim(f.bairro_pesq.value) == "") && (trim(f.cidade_bairro.value) != "") && (trim(f.uf_bairro.value) != "")) {
		        alert("Escolha ao menos 1 (um) bairro para a pesquisa por bairro!!");
		        f.bairro_pesq.focus();
		        return;
		    }
		}
    
		
//  CEP tem tamanho válido?
	if (f.rb_pesquisar_por[2].checked) {
		strCep=retorna_so_digitos(trim(f.c_cep_pesq.value));
		if ((strCep.length!=5)&&(strCep.length!=8)) {
			alert("CEP com tamanho inválido!!");
			f.c_cep_pesq.focus();
			return;
			}
		}

//  Selecionou algum indicador?
	if (f.rb_pesquisar_por[3].checked) {
		if (trim(f.c_indicador.value)=="") {
			alert("Selecione um indicador da lista!!");
			f.c_indicador.focus();
			return;
			}

}

//  Digitou o CPF / CNPJ ?
if (f.rb_pesquisar_por[4].checked) {
    if (trim(f.c_cpfcnpj_pesq.value) == "") {
        alert("Informe o CPF ou CNPJ a ser pesquisado");
        f.c_cpfcnpj_pesq.focus();
        return;
    }
}

//  Selecionou algum vendedor?
	if (f.rb_pesquisar_por[5].checked) {
		if (trim(f.c_vendedor.value)=="") {
			alert("Selecione um vendedor da lista!!");
			f.c_vendedor.focus();
			return;
			}
		}

	dCONFIRMA.style.visibility="hidden";
	window.status = "Aguarde ...";
    
   
	try {
		// Save the search result table when leaving the page.
		var d;

		if (f.rb_pesquisar_por[0].checked) {
			if (('localStorage' in window) && window['localStorage'] !== null) {
				var d = $("#c_escolher_loc").html();
				localStorage.setItem('c_escolher_loc', d);
			}
			if (('localStorage' in window) && window['localStorage'] !== null) {
				var d = $("#c_localidade_pesq").html();
				localStorage.setItem('c_localidade_pesq', d);
			}
		}
		if (f.rb_pesquisar_por[1].checked) {
			if (('localStorage' in window) && window['localStorage'] !== null) {
				var d = $("#c_bairro").html();
				localStorage.setItem('c_bairro', d);
			}
			if (('localStorage' in window) && window['localStorage'] !== null) {
				var d = $("#cidade_bairro").html();
				localStorage.setItem('cidade_bairro', d);
			}
			if (('localStorage' in window) && window['localStorage'] !== null) {
				var d = $("#bairro_pesq").html();
				localStorage.setItem('bairro_pesq', d);
			}
		}
		if (f.rb_pesquisar_por[3].checked) {
			if (('localStorage' in window) && window['localStorage'] !== null) {
				var d = $("#c_indicador").html();
				localStorage.setItem('c_indicador', d);
			}
		}
		if (f.rb_pesquisar_por[5].checked) {
			if (('localStorage' in window) && window['localStorage'] !== null) {
				var d = $("#c_vendedor").html();
				localStorage.setItem('c_vendedor', d);
			}
		}
	}
	catch (e) {
		// NOP
	}


	    fFILTRO.c_hidden_indice_cidade_bairro.value = $("#cidade_bairro option:selected").index();
	    fFILTRO.c_hidden_indice_indicador.value = $("#c_indicador option:selected").index();
	    fFILTRO.c_hidden_indice_vendedor.value = $("#c_vendedor option:selected").index();
	    fFILTRO.c_hidden_reload.value = "1";

	f.submit();
}

</script>

<script type="text/javascript">

    $(function() {
        var f = fFILTRO;
        $("#divMsgAguardeObtendoDados").css('filter', 'alpha(opacity=50)');
        // Replace the search result table on load.
		if ($("#c_hidden_reload").val() == 1) {
			try {
				if (f.rb_pesquisar_por[0].checked) {
					if (('localStorage' in window) && window['localStorage'] !== null) {
						if ('c_escolher_loc' in localStorage) {
							$("#c_escolher_loc").html(localStorage.getItem('c_escolher_loc'));
							$("#c_escolher_loc").prop('selectedIndex', 1);
						}
					}
					if (('localStorage' in window) && window['localStorage'] !== null) {
						if ('c_localidade_pesq' in localStorage) {
							$("#c_localidade_pesq").html(localStorage.getItem('c_localidade_pesq'));
							$("#c_localidade_pesq").prop('selectedIndex', 1);
						}
					}
				}
				if (f.rb_pesquisar_por[1].checked) {
					if (('localStorage' in window) && window['localStorage'] !== null) {
						if ('c_bairro' in localStorage) {
							$("#c_bairro").html(localStorage.getItem('c_bairro'));
							$("#c_bairro").prop('selectedIndex', 1);
						}
					}

					if (('localStorage' in window) && window['localStorage'] !== null) {
						if ('cidade_bairro' in localStorage) {
							$("#cidade_bairro").html(localStorage.getItem('cidade_bairro'));
							$("#cidade_bairro").prop('selectedIndex', fFILTRO.c_hidden_indice_cidade_bairro.value);
						}
					}
					if (('localStorage' in window) && window['localStorage'] !== null) {
						if ('bairro_pesq' in localStorage) {
							$("#bairro_pesq").html(localStorage.getItem('bairro_pesq'));
							$("#bairro_pesq").prop('selectedIndex', 1);
						}
					}
				}
				if (f.rb_pesquisar_por[3].checked) {
					if (('localStorage' in window) && window['localStorage'] !== null) {
						if ('c_indicador' in localStorage) {
							$("#c_indicador").html(localStorage.getItem('c_indicador'));
							$("#c_indicador").prop('selectedIndex', fFILTRO.c_hidden_indice_indicador.value);
						}
					}
				}
				if (f.rb_pesquisar_por[5].checked) {
					if (('localStorage' in window) && window['localStorage'] !== null) {
						if ('c_vendedor' in localStorage) {
							$("#c_vendedor").html(localStorage.getItem('c_vendedor'));
							$("#c_vendedor").prop('selectedIndex', fFILTRO.c_hidden_indice_vendedor.value);
						}
					}
				}
			}
			catch (e) {
				// NOP
			}
		}

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
            $("#divMsgAguardeObtendoDados").css("top", newTop);
        }

        $("#btnAdiciona").click(function() {
            var x = $("#c_escolher_loc option:selected");
            $("#c_localidade_pesq").append(x);
            reOrdenarEscolhidos();
        });

        $("#btnRemove").click(function() {
            var x = $("#c_localidade_pesq option:selected");
            $("#c_escolher_loc").append(x);
            reOrdenarAEscolher();
        });

        $("#c_escolher_loc").dblclick(function() {
            var x = $("#c_escolher_loc option:selected");
            $("#c_localidade_pesq").append(x);
            reOrdenarEscolhidos();
        });

        $("#c_localidade_pesq").dblclick(function() {
            var x = $("#c_localidade_pesq option:selected");
            $("#c_escolher_loc").append(x);
            reOrdenarAEscolher();
        });


        $("#btnAddBairro").click(function() {
            var x = $("#c_bairro option:selected");
            $("#bairro_pesq").append(x);
            reOrdenarEscolhidosBairro();
        });

        $("#btnRmvBairro").click(function() {
            var x = $("#bairro_pesq option:selected");
            $("#c_bairro").append(x);
            reOrdenarAEscolherBairro();
        });

        $("#c_bairro").dblclick(function() {
            var x = $("#c_bairro option:selected");
            $("#bairro_pesq").append(x);
            reOrdenarEscolhidosBairro();
        });

        $("#bairro_pesq").dblclick(function() {
            var x = $("#bairro_pesq option:selected");
            $("#c_bairro").append(x);
            reOrdenarAEscolherBairro();
        });

        // ESCONDER BLOCOS
        var $tabelas = $("#POR_UF_LOCALIDADE, #POR_BAIRRO, #POR_CEP, #POR_NOME, #POR_CPF_CNPJ, #POR_ASSOCIADOS_AO_VENDEDOR");

        $tabelas.hide();
        $('#' + $("input[name='rb_pesquisar_por']:checked").val()).show();

        $("input[name='rb_pesquisar_por']").on('click', function() {
            $tabelas.hide();
            $('#' + $(this).val()).show();
        });

    });

    function reOrdenarAEscolher() {
        $("#c_escolher_loc").html($("#c_escolher_loc option").sort(function(a, b) {
            return a.text.toUpperCase() == b.text.toUpperCase() ? 0 : a.text.toUpperCase() < b.text.toUpperCase() ? -1 : 1
        }))
    }

    function reOrdenarEscolhidos() {
        $("#c_localidade_pesq").html($("#c_localidade_pesq option").sort(function(a, b) {
            return a.text.toUpperCase() == b.text.toUpperCase() ? 0 : a.text.toUpperCase() < b.text.toUpperCase() ? -1 : 1
        }))
    }

    function reOrdenarAEscolherBairro() {
        $("#c_bairro").html($("#c_bairro option").sort(function(a, b) {
            return a.text.toUpperCase() == b.text.toUpperCase() ? 0 : a.text.toUpperCase() < b.text.toUpperCase() ? -1 : 1
        }))
    }

    function reOrdenarEscolhidosBairro() {
        $("#bairro_pesq").html($("#bairro_pesq option").sort(function(a, b) {
            return a.text.toUpperCase() == b.text.toUpperCase() ? 0 : a.text.toUpperCase() < b.text.toUpperCase() ? -1 : 1
        }))
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

<link href="<%=URL_FILE__E_CSS%>" Rel="stylesheet" Type="text/css">

</head>

<body onload="fFILTRO.c_loja.focus()">
<center>

<form id="fFILTRO" name="fFILTRO" method="post" action="PesquisaDeIndicadoresExec.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="c_hidden_indice_vendedor" id="c_hidden_indice_vendedor" value="" />
<input type="hidden" name="c_hidden_indice_indicador" id="c_hidden_indice_indicador" value="" />
<input type="hidden" name="c_hidden_indice_cidade_bairro" id="c_hidden_indice_cidade_bairro" value="" />
<input type="hidden" name="c_hidden_reload" id="c_hidden_reload" value="0" />

<!-- MENSAGEM: "Aguarde, obtendo dados" -->

	<div id="divMsgAguardeObtendoDados" name="divMsgAguardeObtendoDados" style="background-image: url('../Imagem/ajax_loader_gray_256.gif');background-repeat:no-repeat;background-position: center center;position:absolute;bottom:0px;left:0px;width:100%;height:100%;z-index:9;border: 1pt solid #C0C0C0;background-color: black;opacity:0.6;visibility:hidden;vertical-align: middle">

	</div>

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
  <tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Pesquisa de Indicadores</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
  </tr>
</table>
<br>

<table style="width: 640px; text-align: left">
  <tr>
	<td>
	  <!--  LOJA  -->
	  <span class="PLTc" style="vertical-align:middle;cursor:default;">Loja</span>
	  <table class="Qx" cellSpacing="0">
		<tr bgColor="#FFFFFF">
		  <td class="MT" align="center" nowrap>
			<table cellspacing="0" cellpadding="0" style='margin: 12px 8px 12px 8px;'>
			  <tr bgcolor="#FFFFFF">
				<td align="right">
				  <span class="Cd">LOJA</span>
				</td>
				<td>
				  <input id="c_loja" name="c_loja" maxlength="3" size="6" style="text-align:center;" 
						onkeypress="if (digitou_enter(true)) {bCONFIRMA.focus();} filtra_numerico();"
						onchange="LimpaListasDependentesDaLoja();">
				</td>
			  </tr>
			</table>
		  </td>
		</tr>
	  </table>
	</td>
  </tr>

  <tr>
	<td>
	  &nbsp;
	</td>
  </tr>

  <tr>
	<td>
	  <!--  UF / CIDADE  -->
	  <input type="radio" tabindex="-1" id="rb_pesquisar_por" name="rb_pesquisar_por" class="input1" value="<%=COD_PESQUISAR_POR_UF_LOCALIDADE%>">
	  <span class="PLTc" style="vertical-align:middle;cursor:default;" onclick="fFILTRO.rb_pesquisar_por[0].click();">Pesquisar por UF / Localidade</span>
	  <table class="Qx" cellSpacing="0" id="<%=COD_PESQUISAR_POR_UF_LOCALIDADE%>">
		<tr bgColor="#FFFFFF">
		  <td class="MT" align="center" nowrap>
			<table cellspacing="0" cellpadding="0" style='margin: 12px 8px 12px 8px;'>
			  <tr bgcolor="#FFFFFF">
				<td align="left">
				<table><tr>
				  <td>
				  <span class="Cd">UF</span>&nbsp;
				  <select id="c_uf_pesq" name="c_uf_pesq" style="margin-right:10px;"
						onkeyup="if (window.event.keyCode==KEYCODE_DELETE) {this.options[0].selected=true; CarregaLocalidades();}" 
						onchange="CarregaLocalidades();fFILTRO.rb_pesquisar_por[0].click();" 			
						onkeypress="if (digitou_enter(true)) fFILTRO.c_escolher_loc.focus();">
				  <% =UF_monta_itens_select(Null) %>
				  </select>
				  </td><td valign="bottom">
				  <a href="javascript: CarregaLocalidades()" title="Pesquisar cidades da UF selecionada"><img style="border: 0" src="../IMAGEM/lupa_20x20.jpg" /></a>
				</td></tr>
				</table></td>
				<td align="left">
				  &nbsp;
				</td>
			  </tr>
			  <tr>
				<td>
				  &nbsp;
				</td>
			  </tr>
			  <tr>
				<td style="text-align: left">
				  <span class="Cd">Selecione a cidade:</span><br />
				  <select id="c_escolher_loc" name="c_escolher_loc" size="10" style="width:200px;margin-right:10px;" 
						onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;" 
						onkeypress="if (digitou_enter(true)) bCONFIRMA.focus();" 
						onclick="fFILTRO.rb_pesquisar_por[0].click();" multiple>
				  </select>
				</td>
				<td style="text-align: left">
				    <input type="button" id="btnAdiciona" name="btnAdiciona" style="width:40px; margin-bottom:2px;" value="&raquo;" />
					<br />
				    <input type="button" id="btnRemove" name="btnRemove" style="width:40px; margin-bottom:2px;" value="&laquo;" />
				</td>
				<td style="text-align: left"><br><span id="txtEscolhidos" name="txtEscolhidos" class="C" style="margin:1px 10px 6px 10px;">Cidades Selecionadas</span>
					<br>
					<select id="c_localidade_pesq" name="c_localidade_pesq" size="10" style="width: 200px;margin:1px 10px 6px 10px;" multiple>
				    </select>
				</td>
			  </tr>
			</table>
		  </td>
		</tr>
	  </table>
	</td>
  </tr>
  <tr>
	<td>
	  &nbsp;
	</td>
  </tr>
   <tr>
	<td>
	  <!--  BAIRRO  -->
	  	  <input type="radio" tabindex="-1" id="rb_pesquisar_por" name="rb_pesquisar_por" class="input2" value="<%=COD_PESQUISAR_POR_BAIRRO%>">
	  <span class="PLTc" style="vertical-align:middle;cursor:default;" onclick="fFILTRO.rb_pesquisar_por[1].click();">Pesquisar por Bairro</span>
	  <table class="Qx" cellSpacing="0" id="<%=COD_PESQUISAR_POR_BAIRRO%>">
		<tr bgColor="#FFFFFF">
		  <td class="MT" align="center" nowrap>
			<table cellspacing="0" cellpadding="0" style='margin: 12px 8px 12px 8px;'>
			  <tr bgcolor="#FFFFFF">
				<td align="left" valign="top">
				<table cellspacing="0" cellpadding="0" style='margin: 12px 8px 12px 8px;'><tr><td>
				  <span class="Cd">UF</span><br />
				  <select id="uf_bairro" name="uf_bairro" style="margin-right:10px;"
						onkeyup="if (window.event.keyCode==KEYCODE_DELETE) {this.options[0].selected=true; CarregaCidadeBairro();}" 
						onchange="CarregaCidadeBairro();fFILTRO.rb_pesquisar_por[1].click();" 

						onkeypress="if (digitou_enter(true)) fFILTRO.cidade_bairro.focus();">
				  <% =UF_monta_itens_select(Null) %>
				  </select></td><td valign="bottom">
				  <a href="javascript: CarregaCidadeBairro()" title="Pesquisar cidades da UF selecionada"><img style="border:0" src="../IMAGEM/lupa_20x20.jpg" /></a>
				  </td></tr></table>
				</td>
				<td>
				  &nbsp;
				</td>
			  </tr>
			  <tr>
				<td>
				  &nbsp;
				</td>
			  </tr>
			  <tr>
				<td align="left">
				  <span class="Cd">Selecione a cidade:</span><br />
				  <select id="cidade_bairro" name="cidade_bairro" style="width:200px;margin-right:10px;" 
						onkeyup="if (window.event.keyCode==KEYCODE_DELETE) {this.options[0].selected=true; CarregaBairro();}" 
						onkeypress="if (digitou_enter(true)) bCONFIRMA.focus();" 
						onchange="CarregaBairro();fFILTRO.rb_pesquisar_por[1].click();">
				  </select>
				  </td>
				  <td valign="bottom">
				  <a href="javascript: CarregaBairro()"><img style="border: 0" src="../IMAGEM/lupa_20x20.jpg" title="Pesquisar bairros da cidade selecionada" /></a>
				</td>
			  </tr>
			  <tr><td>
			      &nbsp;
			  </td>
			  </tr>
			  <tr>
			    <td style="text-align: left">
			        <span class="Cd">Selecione o(s) bairro(s):</span><br />
				  <select id="c_bairro" name="c_bairro" size="10" style="width:200px;margin-right:10px;" 
						onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;" 
						onkeypress="if (digitou_enter(true)) bCONFIRMA.focus();" 
						onclick="fFILTRO.rb_pesquisar_por[1].click();" multiple>
				  </select>
				</td>
				<td>
				  <input type="button" id="btnAddBairro" name="btnAddBairro" style="width:40px; margin-bottom:2px;" value="&raquo;" />
					<br />
				    <input type="button" id="btnRmvBairro" name="btnRmvBairro" style="width:40px; margin-bottom:2px;" value="&laquo;" />
				</td>
				<td style="text-align: left">
				   <br><span id="Span1" name="BairrosEscolhidosTxt" class="C" style="margin:1px 10px 6px 10px;">Bairro(s) selecionado(s):</span>
					<br>
					<select id="bairro_pesq" name="bairro_pesq" size="10" style="width: 200px;margin:1px 10px 6px 10px;" multiple>
				    </select>
				</td>
			    </tr>
			</table>
		  </td>
		</tr>
	  </table>
	</td>
  </tr>
  <tr>
	<td>
	  &nbsp;
	</td>
  </tr>
  <tr>
	<td>
	  <!--  CEP  -->
	  <input type="radio" tabindex="-1" id="rb_pesquisar_por" name="rb_pesquisar_por" class="input3" value="<%=COD_PESQUISAR_POR_CEP%>">
	  <span class="PLTc" style="vertical-align:middle;cursor:default;" onclick="fFILTRO.rb_pesquisar_por[2].click();fFILTRO.c_cep_pesq.focus()">Pesquisar por CEP</span>
	  <table class="Qx" cellSpacing="0" id="<%=COD_PESQUISAR_POR_CEP%>">
		<tr bgColor="#FFFFFF">
		  <td class="MT" align="center" nowrap>
			<table cellspacing="0" cellpadding="0" style='margin: 12px 8px 12px 8px;'>
			  <tr bgcolor="#FFFFFF">
				<td align="right">
				  <span class="Cd">CEP</span>
				</td>
				<td>
				  <input id="c_cep_pesq" name="c_cep_pesq" maxlength="9" size="11" 
					onkeypress="if (digitou_enter(true)) {bCONFIRMA.focus();} filtra_cep();" 
					onblur="if (cep_ok(this.value)) this.value=cep_formata(this.value);" 
					onchange="fFILTRO.rb_pesquisar_por[2].click();">
				</td>
			  </tr>
			</table>
		  </td>
		</tr>
	  </table>
	</td>
  </tr>

  <tr>
	<td>
	  &nbsp;
	</td>
  </tr>

  <tr>
	<td>
	  <!--  POR NOME DO INDICADOR  -->
	  <input type="radio" tabindex="-1" id="rb_pesquisar_por" name="rb_pesquisar_por" class="input4" value="<%=COD_PESQUISAR_POR_NOME%>">
	  <span class="PLTc" style="vertical-align:middle;cursor:default;" onclick="fFILTRO.rb_pesquisar_por[3].click();">Pesquisar por Nome</span>
	  <table class="Qx" cellspacing="0" id="<%=COD_PESQUISAR_POR_NOME%>">
		<tr bgcolor="#FFFFFF">
		  <td class="MT" align="center" nowrap>
			<table cellspacing="0" cellpadding="0" style='margin: 12px 8px 12px 8px;'>
			  <tr bgcolor="#FFFFFF">
				<td align="right">
				  <span class="Cd">Indicador</span>
				</td>
				<td>
					<select id="c_indicador" name="c_indicador" style="margin-right:10px;width:480px;" 
						onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;"
						onfocus="if (trim(fFILTRO.c_loja.value)=='') {alert('É necessário informar a loja!!'); fFILTRO.c_loja.focus();} else {if (strListaIndicadoresUltimaLojaPesquisada!=trim(fFILTRO.c_loja.value) && this.value =='') CarregaListaIndicadores(fFILTRO.c_loja.value);}"
						onchange="fFILTRO.rb_pesquisar_por[3].click();">
					</select>
				</td>
			  </tr>
			</table>
		  </td>
		</tr>
	  </table>
	</td>
  </tr>
  <tr>
	<td>
	  &nbsp;
	</td>
  </tr>
  
  <tr>
	<td>
	  <!--  POR CPF/CNPJ DO INDICADOR  -->
	  <input type="radio" tabindex="-1" id="rb_pesquisar_por" name="rb_pesquisar_por" class="input5" value="<%=COD_PESQUISAR_POR_CPF_CNPJ%>">
	  <span class="PLTc" style="vertical-align:middle;cursor:default;" onclick="fFILTRO.rb_pesquisar_por[4].click();fFILTRO.c_cpfcnpj_pesq.focus();">Pesquisar por CPF/CNPJ</span>
	  <table class="Qx" cellSpacing="0" id="<%=COD_PESQUISAR_POR_CPF_CNPJ%>">
		<tr bgColor="#FFFFFF">
		  <td class="MT" align="center" nowrap>
			<table cellspacing="0" cellpadding="0" style='margin: 12px 8px 12px 8px;'>
			  <tr bgcolor="#FFFFFF">
				<td align="right">
				  <span class="Cd">CPF/CNPJ:</span>
				</td>
				<td>
				  <input id="c_cpfcnpj_pesq" name="c_cpfcnpj_pesq" maxlength="18" size="18">
				</td>
			  </tr>
			</table>
		  </td>
		</tr>
	  </table>
	</td>
  </tr>

  <tr>
	<td>
	  &nbsp;
	</td>
  </tr>

  <tr>
	<td>
	  <!--  ASSOCIADOS AO VENDEDOR  -->
	  <input type="radio" tabindex="-1" id="rb_pesquisar_por" name="rb_pesquisar_por" class="input6" value="<%=COD_PESQUISAR_ASSOCIADOS_AO_VENDEDOR%>">
	  <span class="PLTc" style="vertical-align:middle;cursor:default;" onclick="fFILTRO.rb_pesquisar_por[5].click();">Pesquisar por Associados ao Vendedor</span>
	  <table class="Qx" cellspacing="0" id="<%=COD_PESQUISAR_ASSOCIADOS_AO_VENDEDOR%>">
		<tr bgcolor="#FFFFFF">
		  <td class="MT" align="center" nowrap>
			<table cellspacing="0" cellpadding="0" style='margin: 12px 8px 12px 8px;'>
			  <tr bgcolor="#FFFFFF">
				<td align="right">
				  <span class="Cd">Vendedor</span>
				</td>
				<td>
					<select id="c_vendedor" name="c_vendedor" style="margin-right:10px;width:480px;" 
						onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;"
						onfocus="if (trim(fFILTRO.c_loja.value)=='') {alert('É necessário informar a loja!!'); fFILTRO.c_loja.focus();} else {if (strListaVendedoresUltimaLojaPesquisada!=trim(fFILTRO.c_loja.value) && this.value == '') CarregaListaVendedores(fFILTRO.c_loja.value);}"
						onchange="fFILTRO.rb_pesquisar_por[5].click();">
					</select>   
				</td>
			  </tr>
			</table>
		  </td>
		</tr>
	  </table>
	</td>
  </tr>

</table>

<!-- ************   SEPARADOR   ************ -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>
<br>


<table width="649" cellspacing="0">
<tr>
	<td><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()" title="volta para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
	<td align="right"><div name="dCONFIRMA" id="dCONFIRMA"><a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fFILTROConfirma(fFILTRO)" title="executa a consulta">
		<img src="../botao/confirmar.gif" width="176" height="55" border="0"></a></div>
	</td>
</tr>
</table>
</form>

</center>
</body>
</html>

<%
'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>
