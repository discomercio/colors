<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =============================================
'	  T R A N S P O R T A D O R A E D I T A . A S P
'     =============================================
'
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


' _____________________________________________________________________________________________
'
'			I N I C I A L I Z A     P Á G I N A     A S P    N O    S E R V I D O R
' _____________________________________________________________________________________________


	On Error GoTo 0
	Err.Clear
	
'	EXIBIÇÃO DE BOTÕES DE PESQUISA DE CEP
	dim blnPesquisaCEPAntiga, blnPesquisaCEPNova
	
	blnPesquisaCEPAntiga = False
	blnPesquisaCEPNova = True
	
	
'	OBTEM O ID
	dim s, usuario, transportadora_selecionada, operacao_selecionada
	dim i
	usuario = trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 
	
'	TRANSPORTADORA A EDITAR
	transportadora_selecionada = ucase(trim(request("transportadora_selecionada")))
	operacao_selecionada = trim(request("operacao_selecionada"))
	
	if (transportadora_selecionada="") then Response.Redirect("aviso.asp?id=" & ERR_TRANSPORTADORA_NAO_ESPECIFICADA) 
	if (operacao_selecionada<>OP_INCLUI) And (operacao_selecionada<>OP_CONSULTA) then Response.Redirect("aviso.asp?id=" & ERR_OPERACAO_NAO_ESPECIFICADA)
	
'	CONECTA COM O BANCO DE DADOS
	dim cn, rs, rscep
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	set rs = cn.Execute("SELECT * FROM t_TRANSPORTADORA WHERE (id='" & transportadora_selecionada & "')")
	if Err <> 0 then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
	
	set rscep = cn.Execute("SELECT * FROM t_TRANSPORTADORA_CEP WHERE (transportadora_id='" & transportadora_selecionada & "') ORDER BY tipo_range, cep_unico, cep_faixa_inicial")
	if Err <> 0 then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
	
	if operacao_selecionada=OP_INCLUI then
		if Not rs.EOF then Response.Redirect("aviso.asp?id=" & ERR_TRANSPORTADORA_JA_CADASTRADA)
	elseif operacao_selecionada=OP_CONSULTA then
		if rs.EOF then Response.Redirect("aviso.asp?id=" & ERR_TRANSPORTADORA_NAO_CADASTRADA)
		end if
	
%>


<%=DOCTYPE_LEGADO%>

<html>


<head>
	<title>CENTRAL ADMINISTRATIVA</title>
	</head>


<%
'		C L I E N T   S I D E   S C R I P T I N G
'
'      CCCCCCC   LLL        IIIII  EEEEEEEEE  NNN   NNN  TTTTTTTTT EEEEEEEEE
'     CCC   CCC  LLL         III   EEE        NNNN  NNN     TTT    EEE
'     CCC        LLL         III   EEE        NNNNN NNN     TTT    EEE
'     CCC        LLL         III   EEEEEE     NNN NNNNN     TTT    EEEEEE
'     CCC        LLL         III   EEE        NNN  NNNN     TTT    EEE
'     CCC   CCC  LLL   LLL   III   EEE        NNN   NNN     TTT    EEE
'      CCCCCCC   LLLLLLLLL  IIIII  EEEEEEEEE  NNN   NNN     TTT    EEEEEEEEE
'
%>

<% if False then 'APENAS P/ HABILITAR O INTELLISENSE DURANTE O DESENVOLVIMENTO!! %>
<script src="../Global/jquery.js" language="JavaScript" type="text/javascript"></script>
<% end if %>

<script src="<%=URL_FILE__JQUERY%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_I18N%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__AJAX_JS%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JANELACEP_JS%>" language="JavaScript" type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
var fCepPopup;

function ProcessaSelecaoCEP(){};

function AbrePesquisaCep(){
var f, strUrl;
	try
		{
	//  SE JÁ HOUVER UMA JANELA DE PESQUISA DE CEP ABERTA, GARANTE QUE ELA SERÁ FECHADA 
	// E UMA NOVA SERÁ CRIADA (EVITA PROBLEMAS C/ O 'WINDOW.OPENER')	
		fCepPopup=window.open("", "AjaxCepPesqPopup","status=1,toolbar=0,location=0,menubar=0,directories=0,resizable=1,scrollbars=1,width=5,height=5,left=0,top=0");
		fCepPopup.close();
		}
	catch (e) {
	 // NOP
		}
	f=fCAD;
	ProcessaSelecaoCEP=TrataCepEnderecoCadastro;
	strUrl="../Global/AjaxCepPesqPopup.asp";
	if (trim(f.cep.value)!="") strUrl=strUrl+"?CepDefault="+trim(f.cep.value);
	fCepPopup=window.open(strUrl, "AjaxCepPesqPopup", "status=1,toolbar=0,location=0,menubar=0,directories=0,resizable=1,scrollbars=1,width=980,height=650,left=0,top=0");
	fCepPopup.focus();
}

function TrataCepEnderecoCadastro(strCep, strUF, strLocalidade, strBairro, strLogradouro, strEnderecoNumero, strEnderecoComplemento) {
var f;
	f=fCAD;
	f.cep.value=cep_formata(strCep);
	f.uf.value=strUF;
	f.cidade.value=strLocalidade;
	f.bairro.value=strBairro;
	f.endereco.value=strLogradouro;
	f.endereco_numero.value=strEnderecoNumero;
	f.endereco_complemento.value=strEnderecoComplemento;
	f.endereco.focus();
	window.status="Concluído";
}

function RemoveTransportadora( f ) {
var b;
	b=window.confirm('Confirma a exclusão da transportadora?');
	if (b){
		f.operacao_selecionada.value=OP_EXCLUI;
		dREMOVE.style.visibility="hidden";
		window.status = "Aguarde ...";
		f.submit();
		}
}

function AtualizaTransportadora( f ) {
	if (trim(f.nome.value)=="") {
		alert('Preencha o nome!!');
		f.nome.focus();
		return;
		}
	if (!cnpj_ok(f.cnpj.value)) {
		alert('CNPJ inválido!!');
		f.cnpj.focus();
		return;
		}
	if (!cep_ok(f.cep.value)) {
		alert('CEP inválido!!');
		f.cep.focus();
		return;
		}
	if (!uf_ok(f.uf.value)) {
		alert('UF inválida!!');
		f.uf.focus();
		return;
		}
	if (!ddd_ok(f.ddd.value)) {
		alert('DDD inválido!!');
		f.ddd.focus();
		return;
		}
	if (!telefone_ok(f.telefone.value)) {
		alert('Telefone inválido!!');
		f.telefone.focus();
		return;
		}
	if (!telefone_ok(f.fax.value)) {
		alert('Fax inválido!!');
		f.fax.focus();
		return;
		}
	if ((trim(f.ddd.value)!="")||(trim(f.telefone.value)!="")||(trim(f.fax.value)!="")) {
		if (trim(f.ddd.value)=="") {
			alert('Preencha o DDD!!');
			f.ddd.focus();
			return;
			}
		if ((trim(f.telefone.value)=="") && (trim(f.fax.value)=="")) {
			alert('Preencha o telefone ou o nº do fax!!');
			f.telefone.focus();
			return;
			}
		}

	if (trim(f.c_email.value) != "") {
		if (!email_ok(f.c_email.value)) {
			alert("Email inválido!!");
			f.c_email.focus();
			return;
		}
	}

	if (trim(f.c_email2.value) != "") {
		if (!email_ok(f.c_email2.value)) {
			alert("Email inválido!!");
			f.c_email2.focus();
			return;
		}
	}

	if ((trim(f.endereco.value) != "") || (trim(f.bairro.value) != "") || (trim(f.cidade.value) != "") || (trim(f.uf.value) != "") || (trim(f.cep.value) != "")) {
		if (trim(f.endereco.value) == "") {
			alert('Preencha o endereço!!');
			f.endereco.focus();
			return;
		}
		if (f.endereco.length >= parseInt(f.MAX_TAMANHO_CAMPO_ENDERECO)) {
			alert('Endereço excede o tamanho máximo permitido (' + f.MAX_TAMANHO_CAMPO_ENDERECO + ' caracteres)!!');
			f.endereco.focus();
			return;
		}
		if (trim(f.endereco_numero.value) == "") {
			alert('Preencha o número do endereço!!');
			f.endereco_numero.focus();
			return;
		}
		if (trim(f.cidade.value) == "") {
			alert('Preencha a cidade do endereço!!');
			f.cidade.focus();
			return;
		}
		if (trim(f.uf.value) == "") {
			alert('Preencha o a UF do endereço!!');
			f.uf.focus();
			return;
		}
		if (trim(f.cep.value) == "") {
			alert('Preencha o CEP do endereço!!');
			f.cep.focus();
			return;
		}
	}

	dATUALIZA.style.visibility="hidden";
	window.status = "Aguarde ...";
	f.submit();
}

$(function() {
	//Código das funções Adicionar, Salvar, Editar e Excluir
	$("#tableAdiciona").hide();
	$("#msgAguarde").hide();
	$("#btnSalvar").bind("click", SalvarCEP);
	$(".btnExcluir").bind("click", ExcluirCEP);
	$("#btnCancelar").bind("click", CancelarCEP);
	$("#btnAdicionar").bind("click", AdicionarCEP);
});


function AdicionarCEP() {
	//Quando o botão Adicionar for clicado: 1) exibe os campos vazios para preencher as informações necessárias; 2) oculta o botão Adicionar
	$("#tableAdiciona").show();
	$("#cepini").val("");
	$("#cepfim").val("");
	$("#btnAdicionar").css("display", "none");
	$("#cepini").focus();
}

function SalvarCEP() {
	var s_row;
	var stipo;
	var sini;
	var sfim;
	var stransp;
	var i;
	var j;
	var tabela;
	var linha;
	var scompini;
	var scompfim;
	var existeinterseccao;

	//obtém os campos do novo CEP de entrega e verifica se os mesmos estão preenchidos;
	//se estiver tudo OK, já deixa a string s_row preenchida com o HTML da nova linha a ser incluída
	stipo = $("#rbcepunico:checked").val();
	sini = retorna_so_digitos($("#cepini").val());
	sfim = retorna_so_digitos($("#cepfim").val());
	if (stipo == 1) {
		if (sini == "") {
			alert("Preencher o CEP!");
			$("#cepini").focus();
			return;
		}
		if (sini.length != 8) {
			alert("O CEP deve ter 8 dígitos!");
			$("#cepini").focus();
			return;
		}
		s_row = "<tr>" +
				"<td align='center' class='MDBE tdColCepTipo'><span class='C'>ÚNICO</span></td>" +
				"<td align='center' class='MDB tdColCepValor'><span class='C'>" + cep_formata(sini) + "</span></td>" +
				"<td class='notPrint'><img src='../botao/botao_x_red.gif' class='btnExcluir' /></td>" +
				"<td class='notPrint'><input type='hidden' name='ocultotipofaixa' id='ocultotipofaixa' value='" + "1" + "'></td>" + //tipo
				"<td class='notPrint'><input type='hidden' name='ocultocepini' id='ocultocepini' value='" + cep_formata(sini) + "'></td>" + //CEP Inicial
				"<td class='notPrint'><input type='hidden' name='ocultocepfim' id='ocultocepfim' value ='" + "" + "'></td>" + //CEP Final
				"<td class='notPrint'><input type='hidden' name='ocultoid' id='ocultoid' value='" + "0" + "'></td>" + //id
				"<td class='notPrint'><input type='hidden' name='ocultoexcluir' id='ocultoexcluir' value='" + "0" + "'></td>" + //excluir
			"</tr>";
	}
	else {
		if (sini == "") {
			alert("Preencher o CEP inicial!");
			$("#cepini").focus();
			return;
		}
		if (sini.length != 8) {
			alert("O CEP inicial deve ter 8 dígitos!");
			$("#cepini").focus();
			return;
		}

		if (sfim == "") {
			alert("Preencher o CEP final!");
			$("#cepfim").focus();
			return;
		}
		if (sfim.length != 8) {
			alert("O CEP final deve ter 8 dígitos!");
			$("#cepfim").focus();
			return;
		}
		if (sini >= sfim) {
			alert("O CEP inicial deve ser menor que o CEP final!");
			return;
		}
		s_row = "<tr>" +
				"<td align='center' class='MDBE tdColCepTipo'><span class='C'>FAIXA</span></td>" +
				"<td align='center' class='MDB tdColCepValor'><span class='C'>" + cep_formata(sini) + " até " + cep_formata(sfim) + "</span></td>" +
				"<td class='notPrint'><img src='../botao/botao_x_red.gif' class='btnExcluir' /></td>" +
				"<td class='notPrint'><input type='hidden' name='ocultotipofaixa' id='ocultotipofaixa' value='" + "2" + "'></td>" + //tipo
				"<td class='notPrint'><input type='hidden' name='ocultocepini' id='ocultocepini' value='" + cep_formata(sini) + "'></td>" + //CEP Inicial
				"<td class='notPrint'><input type='hidden' name='ocultocepfim' id='ocultocepfim' value ='" + cep_formata(sfim) + "'></td>" + //CEP Final
				"<td class='notPrint'><input type='hidden' name='ocultoid' id='ocultoid' value='" + "0" + "'></td>" + //id
				"<td class='notPrint'><input type='hidden' name='ocultoexcluir' id='ocultoexcluir' value='" + "0" + "'></td>" + //excluir
			"</tr>";
	}

	//antes de inserir, verifica nos registros existentes em tela se o CEP/Faixa já está cadastrado para a transportadora, 
	//ou se existe intersecção com algum cadastro prévio
	existeinterseccao = false;
	tabela = $("#tblCEPEntrega tbody");
	tabela.find("tr").each(function(i) {
		if ($(this).children("td:nth-child(8)").find("input[name=ocultoexcluir]").val() == "0") {
			scompini = $(this).children("td:nth-child(5)").find("input[name=ocultocepini]").val();
			scompini = retorna_so_digitos(scompini);
			scompfim = $(this).children("td:nth-child(6)").find("input[name=ocultocepfim]").val();
			scompfim = retorna_so_digitos(scompfim);
			if (sini == scompini) existeinterseccao = true;
			if (scompfim != "") {
				if ((sini >= scompini) && (sini <= scompfim)) existeinterseccao = true;
			}
			if (sfim != "") {
				if (sfim == scompini) existeinterseccao = true;
				if (scompfim != "") {
					if ((sfim >= scompini) && (sfim <= scompfim)) existeinterseccao = true;
				}
				//outra situação: cadastrar uma faixa que contenha um CEP já cadastrado
				//(exemplo: cadastrar a faixa 11111-111 a 22222-222, mas já foi cadastrado o CEP 12345-678, portanto, contido nesta faixa)
				if ((sini <= scompini) && (sfim >= scompini)) existeinterseccao = true;
			}
		}
	});
	if (existeinterseccao) {
		alert("A transportadora já atende a este CEP ou faixa de entrega (ou há intersecção de faixas)");
		return;
	}

	//antes de inserir, verifica nos registros existentes no banco se o CEP/Faixa já está cadastrado para outra transportadora, 
	//ou se existe intersecção com algum cadastro prévio
	//(neste ponto, foi utilizado AJAX para que a verificação seja efetuada através da página AjaxTransportadoraCepPesqBD.asp)
	$("#msgAguarde").show();
	$.ajax({
		type: "GET",
		url: "../GLOBAL/AjaxTransportadoraCepPesqBD.asp",
		data: "cepini='" + sini + "'&cepfim='" + sfim + "'",
		cache: false,
		async: false,
		success: function(response) {
			if (response != "") {
				$("#msgAguarde").hide();
				alert("CEP ou faixa de entrega já é atendido pela transportadora " + response + " (ou há intersecção de faixas)");
				stransp = response;
			}
		},
		error: function(response) {
			$("#msgAguarde").hide();
			alert("Erro ao pesquisar CEP de entrega existentes!");
		}
	});
	if ((stransp != "") && (stransp != null)) {
		$("#msgAguarde").hide();
		return;
	}

	$("#msgAguarde").hide();

	//se o cadastramento estiver liberado, adicionar a linha na tabela de CEPs
	//ATENÇÃO: A GRAVAÇÃO NO BD SÓ OCORRE APÓS O CLIQUE NO BOTÃO CONFIRMAR, OU SEJA, QUANDO A TRANSPORTADORA FOR CADASTRADA/ALTERADA
	$("#tblCEPEntrega tbody").append(s_row);
	$(".btnExcluir").bind("click", ExcluirCEP);
	$("#tableAdiciona").hide();
	$("#btnAdicionar").css("display", "block");
	VerificarExibicaoHeaderCEPs();

}

function ExcluirCEP() {
	//marcar a linha para exclusão (ocultoexcluir=1) e ocultar a mesma
	var par = $(this).parent().parent();
	var tdExcluir = par.children("td:nth-child(8)");
	tdExcluir.html("<input type='hidden' name='ocultoexcluir' id='ocultoexcluir' value='1'>");
	par.hide();
	VerificarExibicaoHeaderCEPs();
}

function CancelarCEP() {
	//interromper o cadastramento de um novo CEP, ocultando os campos e reexibindo o botão Adicionar
	$("#tableAdiciona").hide();
	$("#btnAdicionar").css("display", "block");
}

function MudarTipoCEP() {
	var f = fCAD;
	
	//limpar o campo de CEP final
	$("#cepfim").val("");

	//ao clicar no radio TIPO DE CEP, exibir os campos apropriados
	if ($("#rbcepunico:checked").val() == 1) {
		$("#labelcepini").html("CEP");
		$("#tdcepfim").css("display", "none");
	}
	else {
		$("#labelcepini").html("CEP Inicial");
		$("#tdcepfim").css("display", "block");
	}
	
	f.cepini.focus();
}

function VerificarExibicaoHeaderCEPs() {
	//verificar se o cabeçalho da tabela de CEPs deve ser exibido ou não
	var i;
	var cont;
	var tabela;

	//procurar os inputs "ocultoexibir"; se existirem linhas exibidas (ocultoexibir=0), exibir cabeçalho, senão, ocultar
	cont = 0;
	tabela = $("#tblCEPEntrega tbody");
	tabela.find("tr").each(function(i) {
		if ($(this).children("td:nth-child(8)").find("input[name=ocultoexcluir]").val() == 0) cont = cont + 1;
	});
	if (cont > 0) {
		// É necessário usar "table-header-group", pois "block" causa exibição de colunas c/ largura diferente das colunas de tbody, como se thead e tbody fossem tabelas independentes uma da outra
		$("#headerceps").css("display", "table-header-group");
	}
	else {
		$("#headerceps").css("display", "none");
	}
}
</script>

<script type="text/javascript">
	function exibeJanelaCEP() {
		$.mostraJanelaCEP("cep", "uf", "cidade", "bairro", "endereco", "endereco_numero", "endereco_complemento");
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
<link href="<%=URL_FILE__E_JANELABUSCACEP_CSS%>" rel="stylesheet" type="text/css">

<style type="text/css">
.tdColCepTipo
{
	width:150px;
}
.tdColCepValor
{
	width:200px;
}
</style>


<%	if operacao_selecionada=OP_INCLUI then
		s = "fCAD.nome.focus();"
	else
		s = "focus();"
		end if
%>
<body id="corpoPagina" onload="<%=s%>">

<center>

<!-- #include file = "../global/JanelaBuscaCEP.htm"    -->


<!--  CADASTRO DA TRANSPORTADORA -->
<br />
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
<%	if operacao_selecionada=OP_INCLUI then
		s = "Cadastro de Nova Transportadora"
	else
		s = "Consulta/Edição de Transportadora Cadastrada"
		end if
%>
	<td align="center" valign="bottom"><span class="PEDIDO"><%=s%></span></td>
</tr>
</table>
<br>


<!--  CAMPOS DO CADASTRO  -->
<form id="fCAD" name="fCAD" method="post" action="TransportadoraAtualiza.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="operacao_selecionada" id="operacao_selecionada" value='<%=operacao_selecionada%>'>

<!-- ************   NOME   ************ -->
<table width="649" class="Q" cellspacing="0">
	<tr>
		<td class="MD" width="15%" align="left"><p class="R">TRANSPORTADORA</p><span class="C"><input id="transportadora_selecionada" name="transportadora_selecionada" class="TA" value="<%=transportadora_selecionada%>" readonly size="18" style="text-align:center; color:#0000ff"></span></td>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("nome")) else s=""%>
		<td width="85%" align="left"><p class="R">NOME</p><span class="C"><input id="nome" name="nome" class="TA" type="text" maxlength="30" size="60" value="<%=s%>" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCAD.razao_social.focus(); filtra_nome_identificador();"></span></td>
	</tr>
</table>

<!-- ************   RAZÃO SOCIAL   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("razao_social")) else s=""%>
		<td width="100%" align="left"><p class="R">RAZÃO SOCIAL</p><span class="C"><input id="razao_social" name="razao_social" class="TA" value="<%=s%>" maxlength="60" size="85" onkeypress="if (digitou_enter(true)) fCAD.cnpj.focus(); filtra_nome_identificador();"></span></td>
	</tr>
</table>

<!-- ************   CNPJ/IE   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
<%if operacao_selecionada=OP_CONSULTA then s=cnpj_cpf_formata(Trim("" & rs("cnpj"))) else s=""%>
	<td class="MD" width="50%" align="left"><p class="R">CNPJ</p><span class="C">
		<input id="cnpj" name="cnpj" class="TA" value="<%=s%>" maxlength="18" size="24" onblur="if (!cnpj_ok(this.value)) {alert('CNPJ inválido!'); this.focus();} else this.value=cnpj_formata(this.value);" onkeypress="if (digitou_enter(true) && cnpj_ok(this.value)) fCAD.ie.focus(); filtra_cnpj();"></span></td>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("ie")) else s=""%>
		<td width="50%" align="left"><p class="R">IE</p><span class="C"><input id="ie" name="ie" class="TA" type="text" maxlength="20" size="25" value="<%=s%>" onkeypress="if (digitou_enter(true)) fCAD.endereco.focus(); filtra_nome_identificador();"></span></td>
	</tr>
</table>

<!-- ************   ENDEREÇO   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("endereco")) else s=""%>
		<td width="100%" align="left"><p class="R">ENDEREÇO</p><span class="C"><input id="endereco" name="endereco" class="TA" value="<%=s%>" maxlength="60" style="width:635px;" onkeypress="if (digitou_enter(true)) fCAD.endereco_numero.focus(); filtra_nome_identificador();"></span></td>
	</tr>
</table>

<!-- ************   Nº/COMPLEMENTO   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
	<td class="MD" width="50%" align="left"><p class="R">Nº</p><span class="C">
		<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("endereco_numero")) else s=""%>
		<input id="endereco_numero" name="endereco_numero" class="TA" value="<%=s%>" maxlength="20" style="width:310px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCAD.endereco_complemento.focus(); filtra_nome_identificador();"></span></td>
	<td align="left"><p class="R">COMPLEMENTO</p><span class="C">
		<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("endereco_complemento")) else s=""%>
		<input id="endereco_complemento" name="endereco_complemento" class="TA" value="<%=s%>" maxlength="60" style="width:310px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCAD.bairro.focus(); filtra_nome_identificador();"></span></td>
	</tr>
</table>

<!-- ************   BAIRRO/CIDADE   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("bairro")) else s=""%>
		<td width="50%" class="MD" align="left"><p class="R">BAIRRO</p><span class="C"><input id="bairro" name="bairro" class="TA" value="<%=s%>" maxlength="72" style="width:310px;" onkeypress="if (digitou_enter(true)) fCAD.cidade.focus(); filtra_nome_identificador();"></span></td>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("cidade")) else s=""%>
		<td width="50%" align="left"><p class="R">CIDADE</p><span class="C"><input id="cidade" name="cidade" class="TA" value="<%=s%>" maxlength="60" style="width:310px;" onkeypress="if (digitou_enter(true)) fCAD.uf.focus(); filtra_nome_identificador();"></span></td>
	</tr>
</table>

<!-- ************   UF/CEP   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("uf")) else s=""%>
		<td width="50%" class="MD" align="left"><p class="R">UF</p><span class="C"><input id="uf" name="uf" class="TA" value="<%=s%>" maxlength="2" size="3" onkeypress="if (digitou_enter(true) && uf_ok(this.value)) fCAD.ddd.focus();" onblur="this.value=trim(this.value); if (!uf_ok(this.value)) {alert('UF inválida!!');this.focus();} else this.value=ucase(this.value);"></span></td>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("cep")) else s=""%>
		<td width="25%" align="left"><p class="R">CEP</p><span class="C"><input id="cep" name="cep" readonly tabindex=-1 class="TA" value="<%=cep_formata(s)%>" maxlength="9" size="11" onkeypress="if (digitou_enter(true) && cep_ok(this.value)) fCAD.ddd.focus(); filtra_cep();" onblur="if (!cep_ok(this.value)) {alert('CEP inválido!!');this.focus();} else this.value=cep_formata(this.value);"></span></td>
		<td align="center">
			<% if blnPesquisaCEPAntiga then %>
			<button type="button" name="bPesqCep" id="bPesqCep" style='width:130px;font-size:10pt;' class="Botao" onclick="AbrePesquisaCep();">Pesquisar CEP</button>
			<% end if %>
			<% if blnPesquisaCEPAntiga and blnPesquisaCEPNova then Response.Write "&nbsp;" %>
			<% if blnPesquisaCEPNova then %>
			<button type="button" name="bPesqCep" id="bPesqCep" style='width:130px;font-size:10pt;' class="Botao" onclick="exibeJanelaCEP();">Pesquisar CEP</button>
			<% end if %>
		</td>
	</tr>
</table>

<!-- ************   DDD/TELEFONE/FAX   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("ddd")) else s=""%>
		<td width="15%" class="MD" align="left"><p class="R">DDD</p><span class="C"><input id="ddd" name="ddd" class="TA" value="<%=s%>" maxlength="4" size="5" onkeypress="if (digitou_enter(true) && ddd_ok(this.value)) fCAD.telefone.focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inválido!!');this.focus();}"></span></td>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("telefone")) else s=""%>
		<td width="35%" class="MD" align="left"><p class="R">TELEFONE</p><span class="C"><input id="telefone" name="telefone" class="TA" value="<%=telefone_formata(s)%>" maxlength="11" size="12" onkeypress="if (digitou_enter(true) && telefone_ok(this.value)) fCAD.fax.focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('Telefone inválido!!');this.focus();} else this.value=telefone_formata(this.value);"></span></td>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("fax")) else s=""%>
		<td align="left"><p class="R">FAX</p><span class="C"><input id="fax" name="fax" class="TA" value="<%=telefone_formata(s)%>" maxlength="11" size="12" onkeypress="if (digitou_enter(true) && telefone_ok(this.value)) fCAD.contato.focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('Fax inválido!!');this.focus();} else this.value=telefone_formata(this.value);"></span></td>
	</tr>
</table>

<!-- ************   CONTATO   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("contato")) else s=""%>
		<td width="100%" align="left"><p class="R">CONTATO</p><span class="C"><input id="contato" name="contato" class="TA" value="<%=s%>" maxlength="40" size="85" onkeypress="if (digitou_enter(true)) fCAD.c_email.focus(); filtra_nome_identificador();"></span></td>
	</tr>
</table>

<!-- ************   EMAIL   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("email")) else s=""%>
		<td width="100%" align="left"><p class="R">E-MAIL</p><span class="C"><input id="c_email" name="c_email" class="TA" value="<%=s%>" maxlength="60" size="85" onkeypress="if (digitou_enter(true)) fCAD.c_email2.focus(); filtra_email();"></span></td>
	</tr>
</table>

<!-- ************   EMAIL TRANSPORTADORA 1° TRECHO  ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("email2")) else s=""%>
		<td width="100%" align="left"><p class="R">E-MAIL TRANSPORTADORA 1° TRECHO</p><span class="C"><input id="c_email2" name="c_email2" class="TA" value="<%=s%>" maxlength="60" size="85" onkeypress="if (digitou_enter(true)) bATUALIZA.focus(); filtra_email();"></span></td>
	</tr>
</table>

<!-- ************   CEP DE ENTREGA   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
		<td width="100%" align="left"><span class="R" style="margin-bottom:10pt;">CEP DE ENTREGA</span></td>
	</tr>
	<tr align="center">
		<td width="100%" align="left">
			<table width="100%" cellspacing='0' cellpadding='0' style="margin:10pt;" id="tblCEPEntrega">
				<%if rscep.Eof then %>
					<thead id="headerceps" class="TA" style="display:none">
				<%else%>
					<thead id="headerceps" class="TA">
				<%end if%>
					<tr style='background:#F0FFFF;' nowrap>
						<td align="center" class="MT tdColCepTipo" id="headertipo"><span class="C">TIPO</span></td>
						<td align="center" class="MTBD tdColCepValor" id="headerCEPs"><span class="C">CEP</span></td>
						<td style="background:#FFFFFF;" align="left">&nbsp;</td>
						<td style="background:#FFFFFF;" align="left">&nbsp;</td>
						<td style="background:#FFFFFF;" align="left">&nbsp;</td>
						<td style="background:#FFFFFF;" align="left">&nbsp;</td>
						<td style="background:#FFFFFF;" align="left">&nbsp;</td>
						<td style="background:#FFFFFF;" align="left">&nbsp;</td>
					</tr>
				</thead>
				<tbody>
					<%do while not rscep.Eof%>
					<tr>
						<%if rscep("tipo_range") = 1 then%>
							<%s=cep_formata(rscep("cep_unico"))%>
							<td align="center" class="MDBE tdColCepTipo"><span class="C">ÚNICO</span></td>
							<td align="center" class="MDB tdColCepValor"><span class="C"><%=s%></span></td>
						<%else%>
							<%s=cep_formata(rscep("cep_faixa_inicial"))%>
							<td align="center" class="MDBE tdColCepTipo"><span class="C">FAIXA</span></td>
							<td align="center" class="MDB tdColCepValor"><span class="C"><%=s%> até <%=cep_formata(rscep("cep_faixa_final"))%></span></td>
						<%end if%>
						<td class="notPrint" align="left"><img src='../botao/botao_x_red.gif' class="btnExcluir"/></td>
						<td class="notPrint" align="left"><input type="hidden" name="ocultotipofaixa" id="ocultotipofaixa" value="<%=rscep("tipo_range")%>"></td>
						<td class="notPrint" align="left"><input type="hidden" name="ocultocepini" id="ocultocepini" value="<%=s%>"></td>
						<td class="notPrint" align="left"><input type="hidden" name="ocultocepfim" id="ocultocepfim" value ="<%=cep_formata(rscep("cep_faixa_final"))%>"></td>
						<td class="notPrint" align="left"><input type="hidden" name="ocultoid" id="ocultoid" value="<%=rscep("id")%>"></td>
						<td class="notPrint" align="left"><input type="hidden" name="ocultoexcluir" id="ocultoexcluir" value="0"></td>
					</tr>

					<%rscep.MoveNext%>
					<%loop%>

				</tbody>
			</table>
		</td>
	</tr>
	<tr align="center">
		<td align="left" style="padding-left:10px;padding-bottom:6px;"><span name="btnAdicionar" id="btnAdicionar" style='width:130px;font-size:10pt;' class="Botao" >&nbsp;Adicionar&nbsp;</span></td>
	</tr>
</table>

<!-- ************   ADICIONAR CEP DE ENTREGA   ************ -->
<table id="tableAdiciona" width="649" class="QS" cellspacing="0">
	<tr id="trAdiciona" align="left">
		<td width="120" id="tdtipocep" style="padding-bottom:8px;" align="left" valign="top">
			<p class="R" id="labeltipocep" >TIPO DE CEP</p>
			<span class="C" style="cursor:default;" onclick="fCAD.rbtipocep[0].click();"><input type="radio" name="rbtipocep" id="rbcepunico" value="1" onclick="MudarTipoCEP()" checked> CEP ÚNICO</span>
			<br />
			<span class="C" style="cursor:default;" onclick="fCAD.rbtipocep[1].click();"><input type="radio" name="rbtipocep" id="rbcepfaixa" value="2" onclick="MudarTipoCEP()">FAIXA DE CEP</span>
		</td>
		<td align="left" style="width:20px">&nbsp;</td>
		<td width="340" style="padding-bottom:8px;" align="left" valign="bottom">
			<table cellspacing="0" cellpadding="0">
				<tr>
					<td id="tdcepini" align="left">
						<p class="R" style='margin-right:10pt;' id="labelcepini">CEP</p><span class="TA" style='margin-right:10pt;'><input type="text" name="cepini" id="cepini" value="" maxlength="9" style="width:100px;" onkeypress="if (digitou_enter(true)){if (fCAD.rbtipocep[0].checked){SalvarCEP();}else{fCAD.cepfim.focus();}}" onblur="if (!cep_ok(this.value)) {alert('CEP inválido!!');this.focus();} else this.value=cep_formata(this.value);"></span>
					</td>
					<td id="tdcepfim" style="display:none;" align="left">
						<p class="R" style='margin-left:10pt;' id="labelcepfim">CEP Final</p><span class="TA" style='margin-left:10pt;'><input type="text" name="cepfim" id="cepfim" value="" maxlength="9" style="width:100px;" onkeypress="if (digitou_enter(true)){SalvarCEP();}" onblur="if (!cep_ok(this.value)) {alert('CEP inválido!!');this.focus();} else this.value=cep_formata(this.value);"></span>
					</td>
				</tr>
			</table>
		</td>
		<td width="120" id="tdbotoes" style="padding-bottom:8px;" align="left" valign="bottom">
			<table cellpadding="0" cellspacing="0">
				<tr>
				<td style="padding-left:10px;padding-bottom:10px;" align="left">
					<span name="btnSalvar" id="btnSalvar" style='width:120px;font-size:10pt;' class="Botao" >&nbsp;Salvar&nbsp;</span>
				</td>
				</tr>
				<tr>
				<td style="padding-left:10px;" align="left">
					<span name="btnCancelar" id="btnCancelar" style='width:120px;font-size:10pt;' class="Botao" >&nbsp;Cancelar&nbsp;</span>
				</td>
				</tr>
			</table>
		</td>
	</tr>
</table>

<!-- ************   MENSAGEM: AGUARDE  ************ -->
<table width="649" cellspacing="0">
		<tr id="msgAguarde"><td align="center">
			<table cellpadding="0" cellspacing="0">
				<tr>
				<td valign="middle" align="center"><span style="color:orangered;font-weight:bold;font-style:italic;font-size:10pt;">Aguarde, pesquisando existência de CEP no banco de dados</span></td>
				<td style="width:10px;" align="left">&nbsp;</td>
				<td align="left"><img src="../imagem/aguarde.gif"border="0"></td>
				</tr>
			</table>
		</td></tr>
</table>

<!-- ************   SEPARADOR   ************ -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc" align="left">&nbsp;</td></tr>
</table>
<br>


<table class="notPrint" width="649" cellspacing="0">
<tr>
	<td align="left"><a href="javascript:history.back()" title="cancela as alterações no cadastro">
		<img src="../botao/cancelar.gif" width="176" height="55" border="0"></a></td>
	<%
	s = ""
	if operacao_selecionada=OP_CONSULTA then
		s = "<td align='center'><div name='dREMOVE' id='dREMOVE'><a href='javascript:RemoveTransportadora(fCAD)' "
		s =s + "title='remove a transportadora cadastrada'><img src='../botao/remover.gif' width=176 height=55 border=0></a></div>"
		end if
	%><%=s%>
	<td align="right"><div name="dATUALIZA" id="dATUALIZA">
		<a name="bATUALIZA" id="bATUALIZA" href="javascript:AtualizaTransportadora(fCAD)" title="atualiza o cadastro da transportadora">
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
	rs.Close
	set rs = nothing
	
	cn.Close
	set cn = nothing
%>