<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ===========================================
'	  O R C A M E N T O . A S P
'     ===========================================
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


	On Error GoTo 0
	Err.Clear

	dim s, usuario, loja, orcamento_selecionado
	usuario = Trim(Session("usuario_atual"))
	loja = Trim(Session("loja_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 
	If (loja = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

	orcamento_selecionado = ucase(Trim(request("orcamento_selecionado")))
	if (orcamento_selecionado = "") then Response.Redirect("aviso.asp?id=" & ERR_ORCAMENTO_NAO_ESPECIFICADO)
	s = normaliza_num_orcamento(orcamento_selecionado)
	if s <> "" then orcamento_selecionado = s
	if Len(orcamento_selecionado) > TAM_MAX_ID_ORCAMENTO then Response.Redirect("aviso.asp?id=" & ERR_ORCAMENTO_INVALIDO)
	
	dim url_back
	url_back = Trim(request("url_back"))
	
	dim i, n, s_fabricante, s_produto, s_descricao, s_descricao_html, s_obs, s_qtde, s_preco_lista, s_desc_dado, s_vl_unitario
	dim s_vl_TotalItem, m_TotalItem, m_TotalItemComRA, m_TotalDestePedido, m_TotalDestePedidoComRA
	dim s_preco_NF, m_total_NF, m_total_RA
	dim s_aux, s2, s3, s4, r_loja, r_cliente
	
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	if (Not operacao_permitida(OP_LJA_CONSULTA_ORCAMENTO, s_lista_operacoes_permitidas)) And _
	   (Not operacao_permitida(OP_LJA_CONSULTA_UNIVERSAL_PEDIDO_ORCAMENTO, s_lista_operacoes_permitidas)) then 
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if
	
'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, msg_erro
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	dim r_orcamento, v_item, alerta
	alerta=""
	if Not le_orcamento(orcamento_selecionado, r_orcamento, msg_erro) then 
		alerta = msg_erro
	else
		if Trim(r_orcamento.loja) <> loja then Response.Redirect("aviso.asp?id=" & ERR_ORCAMENTO_INVALIDO)
	'	TEM ACESSO A ESTE PRÉ-PEDIDO?
		if Not operacao_permitida(OP_LJA_CONSULTA_UNIVERSAL_PEDIDO_ORCAMENTO, s_lista_operacoes_permitidas) then 
			if r_orcamento.vendedor <> usuario then Response.Redirect("aviso.asp?id=" & ERR_ORCAMENTO_INVALIDO)
			end if
	'	PRÉ-PEDIDO JÁ VIROU PEDIDO?
		if r_orcamento.st_orc_virou_pedido = 1 then Response.Redirect("Pedido.asp?pedido_selecionado=" & retorna_num_pedido_base(r_orcamento.pedido) & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")))
		if Not le_orcamento_item(orcamento_selecionado, v_item, msg_erro) then alerta = msg_erro
		end if

	if alerta = "" then
		if Not orcamento_calcula_total_NF_e_RA(orcamento_selecionado, m_total_NF, m_total_RA, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
		end if

	dim r_pedido
	if alerta = "" then
		if r_orcamento.st_orc_virou_pedido = 1 then
			if Not le_pedido(r_orcamento.pedido, r_pedido, msg_erro) then alerta = msg_erro
			end if
		end if

	dim blnTemRA
	blnTemRA = False
	if alerta = "" then
		for i=Lbound(v_item) to Ubound(v_item)
			if Trim("" & v_item(i).produto) <> "" then
				if v_item(i).preco_NF <> v_item(i).preco_venda then
					blnTemRA = True
					exit for
					end if
				end if
			next
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


<%=DOCTYPE_LEGADO%>

<html>


<head>
	<title>LOJA<%=MontaNumPrePedidoExibicaoTitleBrowser(orcamento_selecionado)%></title>
	</head>


<% if False then 'APENAS P/ HABILITAR O INTELLISENSE DURANTE O DESENVOLVIMENTO!! %>
<script src="../Global/jquery.js" language="JavaScript" type="text/javascript"></script>
<% end if %>

<script src="<%=URL_FILE__JQUERY%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
	$(document).ready(function() {
		var topo = $('#divConsultaOrcamento').offset().top - parseFloat($('#divConsultaOrcamento').css('margin-top').replace(/auto/, 0)) - parseFloat($('#divConsultaOrcamento').css('padding-top').replace(/auto/, 0));
		$('#divConsultaOrcamento').addClass('divFixo');
		
		$("#divClienteConsultaView").hide();
		$('#divInternoClienteConsultaView').addClass('divFixo');
		sizeDivClienteConsultaView();

		$(document).keyup(function(e) {
		    if (e.keyCode == 27) fechaDivClienteConsultaView();
		});

		$("#divClienteConsultaView").click(function() {
		    fechaDivClienteConsultaView();
		});

		$("#imgFechaDivClienteConsultaView").click(function() {
		    fechaDivClienteConsultaView();
		});

		$(".tdGarInd").hide();
		// Para a nova versão da forma de pagamento
		if ($(".tdGarInd").prev("td").hasClass("MD")) {$(".tdGarInd").prev("td").removeClass("MD")};
		// Para a versão antiga da forma de pagamento
		if ($(".tdGarInd").prev("td").hasClass("MDB")) {$(".tdGarInd").prev("td").removeClass("MDB").addClass("MB")}
});

//Every resize of window
$(window).resize(function() {
    sizeDivClienteConsultaView();
});

function sizeDivClienteConsultaView() {
    var newHeight = $(document).height() + "px";
    $("#divClienteConsultaView").css("height", newHeight);
}

function fechaDivClienteConsultaView() {
    $("#divClienteConsultaView").fadeOut();
    $("#iframeClienteConsultaView").attr("src", "");
}

function fCLIConsultaView(id_cliente, usuario) {
    sizeDivClienteConsultaView();
    $("#iframeClienteConsultaView").attr("src", "ClienteConsultaView.asp?cliente_selecionado=" + id_cliente + "&usuario=" + usuario + "&ocultar_botoes=S");
    $("#divClienteConsultaView").fadeIn();
}

</script>

<script language="JavaScript" type="text/javascript">
function restauraVisibility(nome_controle) {
	var c;
	c = document.getElementById(nome_controle);
	if (c) c.style.visibility = "";
}

function trataCliqueBotao(id_botao) {
	var c;
	c = document.getElementById(id_botao);
	c.style.visibility = "hidden";
	setTimeout("restauraVisibility('" + id_botao + "')", 20000);
}

function fPEDConcluir(s_pedido){
	window.status = "Aguarde ...";
	fPED.pedido_selecionado.value=s_pedido;
	fPED.submit(); 
}

function fCLIEdita( ){
	window.status = "Aguarde ...";
	fCLI.edicao_bloqueada.value = 'N';
	fCLI.submit(); 
}

function fCLIConsulta() {
	window.status = "Aguarde ...";
	fCLI.edicao_bloqueada.value = 'S';
	fCLI.submit();
}

function fORCPESQConclui() {
var c;
	if (trim(fORCPESQ.orcamento_selecionado.value) == "") return;
	if (normaliza_num_orcamento(fORCPESQ.orcamento_selecionado.value) != '') {
		fORCPESQ.orcamento_selecionado.value = normaliza_num_orcamento(fORCPESQ.orcamento_selecionado.value);
	}

	if (isNumeroOrcamento(fORCPESQ.orcamento_selecionado.value)) {
		fORCPESQ.action = "orcamento.asp";
	}
	else {
		fORCPESQ.pedido_selecionado.value = fORCPESQ.orcamento_selecionado.value;
		fORCPESQ.action = "pedido.asp";
	}

	trataCliqueBotao("imgOrcPesq");

	fORCPESQ.submit();
}

function fORCModifica( f ) {
	f.action="OrcamentoEdita.asp";
	dMODIFICA.style.visibility="hidden";
	window.status = "Aguarde ...";
	f.submit();
}

function fORCRemove( f ) {
var b;
	b=window.confirm('Confirma o cancelamento do Pré-Pedido?');
	if (b){
		f.action="OrcamentoCancela.asp";
		dREMOVE.style.visibility="hidden";
		window.status = "Aguarde ...";
		f.submit();
		}
}

function fORCImprime( f ) {
	f.action="OrcamentoImprime.asp";
	dIMPRIME.style.visibility="hidden";
	window.status = "Aguarde ...";
	f.submit();
}

function fORCVirarPedido( f ) {
	if (f.c_ExibirCamposModoSelecaoCD.value == "S") {
		f.action = "OrcamentoVirarPedidoSelManualCD.asp";
	}
	else {
		f.action = "OrcamentoVirarPedido.asp";
	}

	dVIRAPEDIDO.style.visibility = "hidden";
	window.status = "Aguarde ...";
	f.submit();
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
<link href="<%=URL_FILE__EPRINTER_CSS%>" Rel="stylesheet" Type="text/css" media="print">
<link href="<%=URL_FILE__ESCREEN_CSS%>" Rel="stylesheet" Type="text/css" media="screen">

<style TYPE="text/css">
#rb_etg_imediata {
	margin: 0pt 2pt 1pt 15pt;
	vertical-align: top;
	}
#rb_status {
	margin: 0pt 2pt 1pt 15pt;
	vertical-align: top;
	}
#divConsultaOrcamentoWrapper
{
	left:1px;
	position:absolute;
	margin-left:1px;
	width:110px;
	z-index:0;
}
#divConsultaOrcamento
{
	margin-top:60px;
	border: 1px solid #A9A9A9;
	padding-top: 4px;
	padding-bottom: 4px;
	padding-left: 6px;
	padding-right: 6px;
	position: absolute;
	background-color: #F5F5F5;
	top:0;
	z-index:0;
}
#divConsultaOrcamento.divFixo
{
	position:fixed;
	top:0;
}
#divClienteConsultaView
{
	position:absolute;
	top:0;
	left:0;
	width:100%;
	z-index:1000;
	background-color:#808080;
	opacity: 1;
}
#divInternoClienteConsultaView
{
	position:absolute;
	top:6%;
	left:5%;
	width:90%;
	height:90%;
	z-index:1000;
	background-color:#808080;
	opacity: 1;
}
#divInternoClienteConsultaView.divFixo
{
	position:fixed;
	top:6%;
}
#imgFechaDivClienteConsultaView
{
	position:fixed;
	top:6%;
	left: 50%;
	margin-left: -16px; /* -1 * image width / 2 */
	margin-top: -32px;
	z-index:1001;
}
#iframeClienteConsultaView
{
	position:absolute;
	top:0;
	left:0;
	width:100%;
	height:100%;
	border: solid 4px black;
}
</style>


<% if alerta <> "" then %>
<!-- ************************************************************ -->
<!-- **********  PÁGINA PARA EXIBIR MENSAGENS DE ERRO  ********** -->
<!-- ************************************************************ -->
<body onload="bVOLTAR.focus();">
<center>
<br>
<!--  T E L A  -->
<p class="T">A V I S O</p>
<div class="MtAlerta" style="width:600px;font-weight:bold;" align="center"><p style='margin:5px 2px 5px 2px;'><%=alerta%></p></div>
<br><br>
<p class="TracoBottom"></p>
<table cellspacing="0">
<tr>
	<td align="center"><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()"><img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
</tr>
</table>
</center>
</body>




<% else %>
<!-- ************************************************************* -->
<!-- **********  PÁGINA PARA EXIBIR O PRÉ-PEDIDO  ***************** -->
<!-- ************************************************************* -->
<body onload="fORCPESQ.orcamento_selecionado.focus();" link="#ffffff" alink="#ffffff" vlink="#ffffff">

<div id="divConsultaOrcamentoWrapper" class="notPrint">
	<div id="divConsultaOrcamento" class="notPrint">
	<form action="orcamento.asp" id="fORCPESQ" name="fORCPESQ" method="post" onsubmit="if (trim(fORCPESQ.orcamento_selecionado.value)=='')return false;">
	<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
	<span class="Rf">Nº Pré-Pedido</span><br />
	<span class="Rf">ou Pedido</span><br />
	<input maxlength="10" name="orcamento_selecionado" class="C" style="width:75px;margin-left:0px;margin-right:0px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) {fORCPESQConclui();}" onblur="if (normaliza_num_orcamento(this.value)!='') {this.value=normaliza_num_orcamento(this.value);}">
	<input type="hidden" name="pedido_selecionado" value="" />
	<br />
	<center>
	<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="página inicial"><img src="../imagem/home_22x22.png" id="imgPagInicial" alt="página inicial" title="página inicial" style="border:0;margin-top:3px;" onclick="trataCliqueBotao('imgPagInicial');" /></a>
	<input type="image" id="imgOrcPesq" src="../imagem/ok_24x24.png" alt="Submit" style="vertical-align:bottom;margin-left:15px;margin-right:0px;" onclick="fORCPESQConclui();">
	</center>
	</form>
	</div>
</div>

<center>

<form method="post" action="Pedido.asp" id="fPED" name="fPED">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="pedido_selecionado" id="pedido_selecionado" value=''>
</form>

<form id="fORC" name="fORC" method="post">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="orcamento_selecionado" id="orcamento_selecionado" value='<%=orcamento_selecionado%>'>

<% if operacao_permitida(OP_LJA_CADASTRA_NOVO_PEDIDO_SELECAO_MANUAL_CD, s_lista_operacoes_permitidas) then %>
<input type="hidden" name="c_ExibirCamposModoSelecaoCD" value="S" />
<input type="hidden" name="rb_selecao_cd" value="" />
<% else %>
<input type="hidden" name="c_ExibirCamposModoSelecaoCD" value="N" />
<input type="hidden" name="rb_selecao_cd" value="<%=MODO_SELECAO_CD__AUTOMATICO%>" />
<% end if %>

<!--  I D E N T I F I C A Ç Ã O   D O   O R Ç A M E N T O -->
<%=MontaHeaderIdentificacaoPrePedido(orcamento_selecionado, r_orcamento, 649)%>
<br>


<!--  L O J A   -->
<table width="649" class="Q" cellspacing="0">
	<tr>
<%	s = ""
	set r_loja = New cl_LOJA
	if x_loja_bd(r_orcamento.loja, r_loja) then
		with r_loja
			if Trim(.razao_social) <> "" then
				s = Trim(.razao_social)
			else
				s = Trim(.nome)
				end if
			end with
		end if
%>
	<td class="MD" align="left"><p class="Rf">LOJA</p><p class="C"><%=s%>&nbsp;</p></td>
	<td width="145" class="MD" align="left"><p class="Rf">ORÇAMENTISTA</p><p class="C"><%=r_orcamento.orcamentista%>&nbsp;</p></td>
	<td width="145" align="left"><p class="Rf">VENDEDOR</p><p class="C"><%=r_orcamento.vendedor%>&nbsp;</p></td>
	</tr>
	</table>

<br>

<!--  CLIENTE   -->
<table width="649" class="Q" cellspacing="0">
	<tr>
<%	
    s = ""
	set r_cliente = New cl_CLIENTE
	if x_cliente_bd(r_orcamento.id_cliente, r_cliente) then
	
    'le as variáveis da origem certa: ou do pedido ou do cliente, todas comecam com cliente__
    dim cliente__tipo, cliente__cnpj_cpf, cliente__rg, cliente__ie, cliente__nome
    dim cliente__endereco, cliente__endereco_numero, cliente__endereco_complemento, cliente__bairro, cliente__cidade, cliente__uf, cliente__cep
    dim cliente__tel_res, cliente__ddd_res, cliente__tel_com, cliente__ddd_com, cliente__ramal_com, cliente__tel_cel, cliente__ddd_cel
    dim cliente__tel_com_2, cliente__ddd_com_2, cliente__ramal_com_2, cliente__email

    cliente__tipo = r_cliente.tipo
    cliente__cnpj_cpf = r_cliente.cnpj_cpf
	cliente__rg = r_cliente.rg
    cliente__ie = r_cliente.ie
    cliente__nome = r_cliente.nome
    cliente__endereco = r_cliente.endereco
    cliente__endereco_numero = r_cliente.endereco_numero
    cliente__endereco_complemento = r_cliente.endereco_complemento
    cliente__bairro = r_cliente.bairro
    cliente__cidade = r_cliente.cidade
    cliente__uf = r_cliente.uf
    cliente__cep = r_cliente.cep
    cliente__tel_res = r_cliente.tel_res
    cliente__ddd_res = r_cliente.ddd_res
    cliente__tel_com = r_cliente.tel_com
    cliente__ddd_com = r_cliente.ddd_com
    cliente__ramal_com = r_cliente.ramal_com
    cliente__tel_cel = r_cliente.tel_cel
    cliente__ddd_cel = r_cliente.ddd_cel
    cliente__tel_com_2 = r_cliente.tel_com_2
    cliente__ddd_com_2 = r_cliente.ddd_com_2
    cliente__ramal_com_2 = r_cliente.ramal_com_2
    cliente__email = r_cliente.email

    if isActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos and r_orcamento.st_memorizacao_completa_enderecos <> 0 then 
        cliente__tipo = r_orcamento.endereco_tipo_pessoa
        cliente__cnpj_cpf = r_orcamento.endereco_cnpj_cpf
	    cliente__rg = r_orcamento.endereco_rg
        cliente__ie = r_orcamento.endereco_ie
        cliente__nome = r_orcamento.endereco_nome
        cliente__endereco = r_orcamento.endereco_logradouro
        cliente__endereco_numero = r_orcamento.endereco_numero
        cliente__endereco_complemento = r_orcamento.endereco_complemento
        cliente__bairro = r_orcamento.endereco_bairro
        cliente__cidade = r_orcamento.endereco_cidade
        cliente__uf = r_orcamento.endereco_uf
        cliente__cep = r_orcamento.endereco_cep
        cliente__tel_res = r_orcamento.endereco_tel_res
        cliente__ddd_res = r_orcamento.endereco_ddd_res
        cliente__tel_com = r_orcamento.endereco_tel_com
        cliente__ddd_com = r_orcamento.endereco_ddd_com
        cliente__ramal_com = r_orcamento.endereco_ramal_com
        cliente__tel_cel = r_orcamento.endereco_tel_cel
        cliente__ddd_cel = r_orcamento.endereco_ddd_cel
        cliente__tel_com_2 = r_orcamento.endereco_tel_com_2
        cliente__ddd_com_2 = r_orcamento.endereco_ddd_com_2
        cliente__ramal_com_2 = r_orcamento.endereco_ramal_com_2
        cliente__email = r_orcamento.endereco_email
        end if
%>	
<%	if cliente__tipo = ID_PF then s_aux="CPF" else s_aux="CNPJ"
	s = cnpj_cpf_formata(cliente__cnpj_cpf) 
%>
        <td width="50%" class="MD" align="left"><p class="Rf"><%=s_aux%></p>
		<% if operacao_permitida(OP_LJA_EDITA_CLIENTE_DADOS_CADASTRAIS, s_lista_operacoes_permitidas) then %>
			<a href='javascript:fCLIEdita();' title='clique para editar o cadastro do cliente'><p class="C"><%=s%>&nbsp;</p></a>
		<% else %>
			<a href='javascript:fCLIConsulta();' title='clique para consultar o cadastro do cliente'><p class="C"><%=s%>&nbsp;</p></a>
		<% end if %>
		</td>
		<%
		if cliente__tipo = ID_PF then s = Trim(cliente__rg) else s = Trim(cliente__ie)
			if cliente__tipo = ID_PF then 
%>
	<td align="left" class="MD"><p class="Rf">RG</p><p class="C"><%=s%>&nbsp;</p></td>
<% else %>
	<td align="left" class="MD"><p class="Rf">IE</p><p class="C"><%=s%>&nbsp;</p></td>
<% end if %>
<td align="center" valign="middle" style="width:22px;"><a href='javascript:fCLIConsultaView(<%=chr(34) & r_cliente.id & chr(34) & "," & chr(34) & usuario & chr(34)%>);' title="clique para visualizar o cadastro do cliente"><img id="imgClienteConsultaView" src="../imagem/doc_preview_22.png" /></a></td>
		</tr>
		<tr>
<%
			if Trim(cliente__nome) <> "" then
				s = Trim(cliente__nome)
				end if
		end if
	
	if cliente__tipo = ID_PF then s_aux="NOME DO CLIENTE" else s_aux="RAZÃO SOCIAL DO CLIENTE"
%>
	<td class="MC" align="left" colspan="3"><p class="Rf"><%=s_aux%></p>
	<% if operacao_permitida(OP_LJA_EDITA_CLIENTE_DADOS_CADASTRAIS, s_lista_operacoes_permitidas) then %>
		<a href='javascript:fCLIEdita();' title='clique para editar o cadastro do cliente'><p class="C"><%=s%>&nbsp;</p></a>
	<% else %>
		<a href='javascript:fCLIConsulta();' title='clique para consultar o cadastro do cliente'><p class="C"><%=s%>&nbsp;</p></a>
	<% end if %>
		</td>
		
		

	</tr>
	</table>

<!--  ENDEREÇO DO CLIENTE  -->
<table width="649" class="QS" cellspacing="0">
	<tr>
<%	
	s = formata_endereco(cliente__endereco, cliente__endereco_numero, cliente__endereco_complemento, cliente__bairro, cliente__cidade, cliente__uf, cliente__cep)
%>		
		<td align="left"><p class="Rf">ENDEREÇO</p><p class="C"><%=s%>&nbsp;</p></td>
	</tr>
</table>

<!--  TELEFONE DO CLIENTE  -->
<table width="649" class="QS" cellspacing="0">
	<tr>
<%	s = ""
	if Trim(cliente__tel_res) <> "" then
		s = telefone_formata(Trim(cliente__tel_res))
		s_aux=Trim(cliente__ddd_res)
		if s_aux<>"" then s = "(" & s_aux & ") " & s
		end if
	
	s2 = ""
	if Trim(cliente__tel_com) <> "" then
		s2 = telefone_formata(Trim(cliente__tel_com))
		s_aux = Trim(cliente__ddd_com)
		if s_aux<>"" then s2 = "(" & s_aux & ") " & s2
		s_aux = Trim(cliente__ramal_com)
		if s_aux<>"" then s2 = s2 & "  (R. " & s_aux & ")"
		end if
	if Trim(cliente__tel_cel) <> "" then
		s3 = telefone_formata(Trim(cliente__tel_cel))
		s_aux = Trim(cliente__ddd_cel)
		if s_aux<>"" then s3 = "(" & s_aux & ") " & s3
		end if
	if Trim(cliente__tel_com_2) <> "" then
		s4 = telefone_formata(Trim(cliente__tel_com_2))
		s_aux = Trim(cliente__ddd_com_2)
		if s_aux<>"" then s4 = "(" & s_aux & ") " & s4
		s_aux = Trim(cliente__ramal_com_2)
		if s_aux<>"" then s4 = s4 & "  (R. " & s_aux & ")"
		end if
	
%>

<% if cliente__tipo = ID_PF then %>
	<td class="MD" width="33%" align="left"><p class="Rf">TELEFONE RESIDENCIAL</p><p class="C"><%=s%>&nbsp;</p></td>
	<td class="MD" width="33%" align="left"><p class="Rf">TELEFONE COMERCIAL</p><p class="C"><%=s2%>&nbsp;</p></td>
		<td align="left"><p class="Rf">CELULAR</p><p class="C"><%=s3%>&nbsp;</p></td>

<% else %>
	<td class="MD" width="50%" align="left"><p class="Rf">TELEFONE</p><p class="C"><%=s2%>&nbsp;</p></td>
	<td width="50%" align="left"><p class="Rf">TELEFONE</p><p class="C"><%=s4%>&nbsp;</p></td>

<% end if %>

	</tr>
</table>


<!--  E-MAIL DO CLIENTE  -->
<table width="649" class="QS" cellspacing="0" >
	<tr>
		<td align="left"><p class="Rf">E-MAIL</p><p class="C"><%=Trim(cliente__email)%>&nbsp;</p></td>
	</tr>
</table>

<!--  ENDEREÇO DE ENTREGA  -->
<%
	s = pedido_formata_endereco_entrega(r_orcamento, r_cliente)
%>		
<table width="649" class="QS" cellspacing="0" style="table-layout:fixed">
	<tr>
		<td align="left"><p class="Rf">ENDEREÇO DE ENTREGA</p><p class="C"><%=s%>&nbsp;</p></td>
	</tr>
    <%	if r_orcamento.EndEtg_cod_justificativa <> "" then %>	
    <tr>
		<td align="left" style="word-wrap:break-word"><p class="C" ><%=obtem_descricao_tabela_t_codigo_descricao(GRUPO_T_CODIGO_DESCRICAO__ENDETG_JUSTIFICATIVA,r_orcamento.EndEtg_cod_justificativa)%>&nbsp;</p></td>
	</tr>
    <%end if %>
</table>


<!--  R E L A Ç Ã O   D E   P R O D U T O S  -->
<br>
<br>
<table class="Qx" cellspacing="0">
	<tr bgColor="#FFFFFF">
	<td class="MB" align="left" valign="bottom"><span class="PLTe">Fabr</span></td>
	<td class="MB" align="left" valign="bottom"><span class="PLTe">Produto</span></td>
	<td class="MB" align="left" valign="bottom"><span class="PLTe" style="width:287px;">Descrição/Observações</span></td>
	<td class="MB" align="right" valign="bottom"><span class="PLTd">Qtd</span></td>
	<% if (r_orcamento.permite_RA_status = 1) Or blnTemRA then %>
	<td class="MB" align="right" valign="bottom"><span class="PLTd">Preço</span></td>
	<% end if %>
	<td class="MB" align="right" valign="bottom"><span class="PLTd">VL Lista</span></td>
	<td class="MB" align="right" valign="bottom"><span class="PLTd">Desc</span></td>
	<td class="MB" align="right" valign="bottom"><span class="PLTd">VL Venda</span></td>
	<td class="MB" align="right" valign="bottom"><span class="PLTd">VL Total</span></td>
	</tr>

<% m_TotalDestePedido=0
   m_TotalDestePedidoComRA=0
   n = Lbound(v_item)-1
   for i=1 to MAX_ITENS 
	 n = n+1
	 if n <= Ubound(v_item) then
		with v_item(n)
			s_fabricante=.fabricante
			s_produto=.produto
			s_descricao=.descricao
			s_descricao_html=produto_formata_descricao_em_html(.descricao_html)
			s_obs=.obs
			if (s_descricao_html<>"") And (s_obs<>"") then s_obs=" (" & s_obs & ")"
			s_qtde=.qtde
			s_preco_lista=formata_moeda(.preco_lista)
			if .desc_dado=0 then s_desc_dado="" else s_desc_dado=formata_perc_desc(.desc_dado)
			s_vl_unitario=formata_moeda(.preco_venda)
			s_preco_NF=formata_moeda(.preco_NF)
			m_TotalItem=.qtde * .preco_venda
			m_TotalItemComRA=.qtde * .preco_NF
			s_vl_TotalItem=formata_moeda(m_TotalItem)
			m_TotalDestePedido=m_TotalDestePedido + m_TotalItem
			m_TotalDestePedidoComRA=m_TotalDestePedidoComRA + m_TotalItemComRA
			end with
	 else
		s_fabricante=""
		s_produto=""
		s_descricao=""
		s_descricao_html=""
		s_obs=""
		s_qtde=""
		s_preco_lista=""
		s_desc_dado=""
		s_vl_unitario=""
		s_preco_NF=""
		s_vl_TotalItem=""
		end if

'	A VERSÃO 5.0 DO IE NÃO DESENHA AS MARGENS SE O SPAN NÃO POSSUIR CONTEÚDO
	if s_descricao = "" then s_descricao = "&nbsp;"
	if s_descricao_html = "" then s_descricao_html = "&nbsp;"
	if s_obs = "" then s_obs = "&nbsp;"
	
%>
	<% if (i > MIN_LINHAS_ITENS_IMPRESSAO_ORCAMENTO) And (s_produto = "") then %>
	<tr class="notPrint">
	<% else %>
	<tr>
	<% end if %>
	<td class="MDBE" align="left"><input name="c_fabricante" id="c_fabricante" class="PLLe" style="width:25px;"
		value='<%=s_fabricante%>' readonly tabindex=-1></td>
	<td class="MDB" align="left"><input name="c_produto" id="c_produto" class="PLLe" style="width:54px;"
		value='<%=s_produto%>' readonly tabindex=-1></td>
	<td class="MDB" align="left" style="width:277px;"><span name="c_descricao" id="c_descricao" class="PLLe" style="margin-left:2px;"><%=s_descricao_html%></span>
					<span name="c_obs" id="c_obs" class="PLLe" style="color:navy;"><%=s_obs%></span></td>
	<td class="MDB" align="right"><input name="c_qtde" id="c_qtde" class="PLLd" style="width:21px;"
		value='<%=s_qtde%>' readonly tabindex=-1></td>
	<% if (r_orcamento.permite_RA_status = 1) Or blnTemRA then %>
	<td class="MDB" align="right"><input name="c_vl_NF" id="c_vl_NF" class="PLLd" style="width:62px;"
		value='<%=s_preco_NF%>' readonly tabindex=-1></td>
	<% end if %>
	<td class="MDB" align="right"><input name="c_preco_lista" id="c_preco_lista" class="PLLd" style="width:62px;"
		value='<%=s_preco_lista%>' readonly tabindex=-1></td>
	<td class="MDB" align="right"><input name="c_desc" id="c_desc" class="PLLd" style="width:28px;"
		value='<%=s_desc_dado%>' readonly tabindex=-1></td>
	<td class="MDB" align="right"><input name="c_vl_unitario" id="c_vl_unitario" class="PLLd" style="width:62px;"
		value='<%=s_vl_unitario%>' readonly tabindex=-1></td>
	<td class="MDB" align="right"><input name="c_vl_total" id="c_vl_total" class="PLLd" style="width:70px;" 
		value='<%=s_vl_TotalItem%>' readonly tabindex=-1></td>
	</tr>
<% next %>

	<tr>
	<td colspan="3" align="left">
		<table cellspacing="0" cellpadding="0" width='100%' style="margin-top:4px;">
			<tr>
			<td width="60%" align="left">&nbsp;</td>
			<% if (r_orcamento.permite_RA_status = 1) Or blnTemRA then %>
			<td align="right">
				<table cellspacing="0" cellpadding="0" style="margin-right:2px;">
					<tr>
						<td class="MTBE" align="left"><span class="PLTe">&nbsp;RA</span></td>
						<td class="MTBD" align="right"><input name="c_total_RA" id="c_total_RA" class="PLLd" style="width:70px;color:<%if m_total_RA >=0 then Response.Write " green" else Response.Write " red"%>;" 
							value='<%=formata_moeda(m_total_RA)%>' readonly tabindex=-1></td>
					</tr>
				</table>
			</td>
			<% end if %>

			<td align="right">
				<table cellspacing="0" cellpadding="0" style="margin-right:2px;">
					<tr>
						<td class="MTBE" align="left" nowrap><span class="PLTe">&nbsp;COM(%)</span></td>
						<td class="MTBD" align="left"><input name="c_perc_RT" id="c_perc_RT" class="PLLd" style="width:30px;color:blue;" 
							value='<%=formata_perc_RT(r_orcamento.perc_RT)%>' readonly tabindex=-1></td>
					</tr>
				</table>
			</td>
			</tr>
		</table>
	</td>

	<% if (r_orcamento.permite_RA_status = 1) Or blnTemRA then %>
	<td class="MD" align="left">&nbsp;</td>
	<td class="MDB" align="right">
		<input name="c_total_NF" id="c_total_NF" class="PLLd" style="width:70px;color:blue;" 
				value='<%=formata_moeda(m_TotalDestePedidoComRA)%>' readonly tabindex=-1>
	</td>
	<td colspan="3" class="MD" align="left">&nbsp;</td>
	<% else %>
	<td colspan="4" class="MD" align="left">&nbsp;</td>
	<% end if %>

	<td class="MDB" align="right"><input name="c_total_geral" id="c_total_geral" class="PLLd" style="width:70px;color:blue;" 
		value='<%=formata_moeda(m_TotalDestePedido)%>' readonly tabindex=-1></td>
	</tr>
</table>

<% if r_orcamento.tipo_parcelamento = 0 then %>
<!--  TRATA VERSÃO ANTIGA DA FORMA DE PAGAMENTO   -->
<br>
<table class="Q" cellspacing="0" style="width:649px;">
	<tr>
		<td class="MB" colspan="5" align="left"><p class="Rf">Observações I</p>
			<textarea name="c_obs1" id="c_obs1" class="PLLe notPrint" rows="<%=Cstr(MAX_LINHAS_OBS1)%>" 
				style="width:642px;margin-left:2pt;" 
				readonly tabindex=-1><%=r_orcamento.obs_1%></textarea>
			<span class="PLLe notVisible"><%
				s = substitui_caracteres(r_orcamento.obs_1,chr(13),"<br>")
				if s = "" then s = "&nbsp;"
				Response.Write s %></span>
		</td>
	</tr>
	<tr>
		<td class="MB" colspan="5" align="left"><p class="Rf">Observações II</p>
			<input name="c_obs2" id="c_obs2" class="PLLe" style="width:85px;margin-left:2pt;" 
				readonly tabindex=-1 value='<%=r_orcamento.obs_2%>'>
		</td>
	</tr>
	<tr>
		<td class="MDB" nowrap width="10%" align="left"><p class="Rf">Parcelas</p>
			<input name="c_qtde_parcelas" id="c_qtde_parcelas" class="PLLc" style="width:60px;"
				readonly tabindex=-1 value='<%if (r_orcamento.qtde_parcelas<>0) Or (r_orcamento.forma_pagto<>"") then Response.write Cstr(r_orcamento.qtde_parcelas)%>'>
		</td>
		<td class="MDB" nowrap align="left" valign="top"><p class="Rf">Entrega Imediata</p>
		<% 	if Cstr(r_orcamento.st_etg_imediata) = Cstr(COD_ETG_IMEDIATA_NAO) then
				s = "NÃO"
			elseif Cstr(r_orcamento.st_etg_imediata) = Cstr(COD_ETG_IMEDIATA_SIM) then
				s = "SIM"
			else
				s = ""
				end if
			
			if s <> "" then
				s_aux=formata_data_e_talvez_hora_hhmm(r_orcamento.etg_imediata_data)
				if s_aux <> "" then s = s & " &nbsp; (" & r_orcamento.etg_imediata_usuario & " em " & s_aux & ")"
				end if
			if s="" then s="&nbsp;"
		%>
		<p class="C" style="margin-top:3px;"><%=s%></p>
		</td>
		<td class="MDB" nowrap align="left" valign="top"><p class="Rf">Bem de Uso/Consumo</p>
		<% 	if Cstr(r_orcamento.StBemUsoConsumo) = Cstr(COD_ST_BEM_USO_CONSUMO_NAO) then
				s = "NÃO"
			elseif Cstr(r_orcamento.StBemUsoConsumo) = Cstr(COD_ST_BEM_USO_CONSUMO_SIM) then
				s = "SIM"
			else
				s = ""
				end if
		
			if s="" then s="&nbsp;"
		%>
		<p class="C" style="margin-top:3px;"><%=s%></p>
		</td>
		<td class="MDB" nowrap align="left" valign="top"><p class="Rf">Instalador Instala</p>
		<% 	if Cstr(r_orcamento.InstaladorInstalaStatus) = Cstr(COD_INSTALADOR_INSTALA_NAO) then
				s = "NÃO"
			elseif Cstr(r_orcamento.InstaladorInstalaStatus) = Cstr(COD_INSTALADOR_INSTALA_SIM) then
				s = "SIM"
			else
				s = ""
				end if
		
			if s="" then s="&nbsp;"
		%>
		<p class="C" style="margin-top:3px;"><%=s%></p>
		</td>
		<td class="MB tdGarInd" nowrap align="left" valign="top"><p class="Rf">Garantia Indicador</p>
		<% 	if Cstr(r_orcamento.GarantiaIndicadorStatus) = Cstr(COD_GARANTIA_INDICADOR_STATUS__NAO) then
				s = "NÃO"
			elseif Cstr(r_orcamento.GarantiaIndicadorStatus) = Cstr(COD_GARANTIA_INDICADOR_STATUS__SIM) then
				s = "SIM"
			else
				s = ""
				end if
		
			if s="" then s="&nbsp;"
		%>
		<p class="C" style="margin-top:3px;"><%=s%></p>
		</td>
	</tr>
	<tr>
		<td colspan="5" align="left"><p class="Rf">Forma de Pagamento</p>
			<textarea name="c_forma_pagto" id="c_forma_pagto" class="PLLe notPrint" rows="<%=Cstr(MAX_LINHAS_FORMA_PAGTO)%>"
				style="width:642px;margin-left:2pt;"
				readonly tabindex=-1><%=r_orcamento.forma_pagto%></textarea>
		<span class="PLLe notVisible"><%
			s = substitui_caracteres(r_orcamento.forma_pagto,chr(13),"<br>")
			if s = "" then s = "&nbsp;"
			Response.Write s %></span>
		</td>
	</tr>
</table>
<% else %>
<!--  TRATA NOVA VERSÃO DA FORMA DE PAGAMENTO   -->
<br>
<table class="Q" cellspacing="0" style="width:649px;">
	<tr>
		<td class="MB" colspan="5" align="left"><p class="Rf">Observações </p>
			<textarea name="c_obs1" id="c_obs1" class="PLLe notPrint" rows="<%=Cstr(MAX_LINHAS_OBS1)%>" 
				style="width:642px;margin-left:2pt;" 
				readonly tabindex=-1><%=r_orcamento.obs_1%></textarea>
			<span class="PLLe notVisible"><%
				s = substitui_caracteres(r_orcamento.obs_1,chr(13),"<br>")
				if s = "" then s = "&nbsp;"
				Response.Write s %></span>
		</td>
	</tr>
	<tr>
		<td class="MD" align="left" nowrap><p class="Rf">Nº Nota Fiscal</p>
			<input name="c_obs2" id="c_obs2" class="PLLe" style="width:85px;margin-left:2pt;" 
				readonly tabindex=-1 value='<%=r_orcamento.obs_2%>'>
		</td>
		<td class="MD" nowrap align="left" valign="top"><p class="Rf">Entrega Imediata</p>
		<% 	if Cstr(r_orcamento.st_etg_imediata) = Cstr(COD_ETG_IMEDIATA_NAO) then
				s = "NÃO"
			elseif Cstr(r_orcamento.st_etg_imediata) = Cstr(COD_ETG_IMEDIATA_SIM) then
				s = "SIM"
			else
				s = ""
				end if
			
			if s <> "" then
				s_aux=formata_data_e_talvez_hora_hhmm(r_orcamento.etg_imediata_data)
				if s_aux <> "" then s = s & " &nbsp; (" & r_orcamento.etg_imediata_usuario & " em " & s_aux & ")"
				end if
			if s="" then s="&nbsp;"
		%>
		<p class="C" style="margin-top:3px;"><%=s%></p>
		</td>
		<td class="MD" nowrap align="left" valign="top"><p class="Rf">Bem de Uso/Consumo</p>
		<% 	if Cstr(r_orcamento.StBemUsoConsumo) = Cstr(COD_ST_BEM_USO_CONSUMO_NAO) then
				s = "NÃO"
			elseif Cstr(r_orcamento.StBemUsoConsumo) = Cstr(COD_ST_BEM_USO_CONSUMO_SIM) then
				s = "SIM"
			else
				s = ""
				end if
		
			if s="" then s="&nbsp;"
		%>
		<p class="C" style="margin-top:3px;"><%=s%></p>
		</td>
		<td class="MD" nowrap align="left" valign="top"><p class="Rf">Instalador Instala</p>
		<% 	if Cstr(r_orcamento.InstaladorInstalaStatus) = Cstr(COD_INSTALADOR_INSTALA_NAO) then
				s = "NÃO"
			elseif Cstr(r_orcamento.InstaladorInstalaStatus) = Cstr(COD_INSTALADOR_INSTALA_SIM) then
				s = "SIM"
			else
				s = ""
				end if
		 
			if s="" then s="&nbsp;"
		%>
		<p class="C" style="margin-top:3px;"><%=s%></p>
		</td>
		<td class="tdGarInd" nowrap align="left" valign="top"><p class="Rf">Garantia Indicador</p>
		<% 	if Cstr(r_orcamento.GarantiaIndicadorStatus) = Cstr(COD_GARANTIA_INDICADOR_STATUS__NAO) then
				s = "NÃO"
			elseif Cstr(r_orcamento.GarantiaIndicadorStatus) = Cstr(COD_GARANTIA_INDICADOR_STATUS__SIM) then
				s = "SIM"
			else
				s = ""
				end if
		
			if s="" then s="&nbsp;"
		%>
		<p class="C" style="margin-top:3px;"><%=s%></p>
		</td>
	</tr>
</table>
<br>
<table class="Q" style="width:649px;" cellspacing="0">
  <tr>
	<td align="left"><span class="Rf">Forma de Pagamento</span></td>
  </tr>
  <tr>
	<td align="left">
	  <table width="100%" cellspacing="0" cellpadding="0" border="0">
		<% if Cstr(r_orcamento.tipo_parcelamento) = COD_FORMA_PAGTO_A_VISTA then %>
		<!--  À VISTA  -->
		<tr>
		  <td align="left">
			<table cellspacing="0" cellpadding="0" border="0">
			  <tr>
				<td align="left"><span class="C">À Vista&nbsp&nbsp;(<%=x_opcao_forma_pagamento(r_orcamento.av_forma_pagto)%>)</span></td>
			  </tr>
			</table>
		  </td>
		</tr>
		<% elseif Cstr(r_orcamento.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELA_UNICA then %>
		<!--  PARCELA ÚNICA  -->
		<tr>
		  <td align="left">
			<table cellspacing="0" cellpadding="0" border="0">
			  <tr>
				<td align="left"><span class="C">Parcela Única:&nbsp;&nbsp;<%=SIMBOLO_MONETARIO%>&nbsp;<%=formata_moeda(r_orcamento.pu_valor)%>&nbsp;&nbsp;&nbsp;(<%=x_opcao_forma_pagamento(r_orcamento.pu_forma_pagto)%>)&nbsp;&nbsp;vencendo após&nbsp;<%=formata_inteiro(r_orcamento.pu_vencto_apos)%>&nbsp;dias</span></td>
			  </tr>
			</table>
		  </td>
		</tr>
		<% elseif Cstr(r_orcamento.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_CARTAO then %>
		<!--  PARCELADO NO CARTÃO (INTERNET)  -->
		<tr>
		  <td align="left">
			<table cellspacing="0" cellpadding="0" border="0">
			  <tr>
				<td align="left"><span class="C">Parcelado no Cartão (internet) em&nbsp;&nbsp;<%=Cstr(r_orcamento.pc_qtde_parcelas)%>&nbsp;x&nbsp;&nbsp;<%=SIMBOLO_MONETARIO & " " & formata_moeda(r_orcamento.pc_valor_parcela)%></span></td>
			  </tr>
			</table>
		  </td>
		</tr>
		<% elseif Cstr(r_orcamento.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_CARTAO_MAQUINETA then %>
		<!--  PARCELADO NO CARTÃO (MAQUINETA)  -->
		<tr>
		  <td align="left">
			<table cellspacing="0" cellpadding="0" border="0">
			  <tr>
				<td align="left"><span class="C">Parcelado no Cartão (maquineta) em&nbsp;&nbsp;<%=Cstr(r_orcamento.pc_maquineta_qtde_parcelas)%>&nbsp;x&nbsp;&nbsp;<%=SIMBOLO_MONETARIO & " " & formata_moeda(r_orcamento.pc_maquineta_valor_parcela)%></span></td>
			  </tr>
			</table>
		  </td>
		</tr>
		<% elseif Cstr(r_orcamento.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_COM_ENTRADA then %>
		<!--  PARCELADO COM ENTRADA  -->
		<tr>
		  <td align="left">
			<table cellspacing="0" cellpadding="0" border="0">
			  <tr>
				<td align="left"><span class="C">Entrada:&nbsp;&nbsp;<%=SIMBOLO_MONETARIO%>&nbsp;<%=formata_moeda(r_orcamento.pce_entrada_valor)%>&nbsp;&nbsp;&nbsp;(<%=x_opcao_forma_pagamento(r_orcamento.pce_forma_pagto_entrada)%>)</span></td>
			  </tr>
			  <tr>
				<td align="left"><span class="C">Prestações:&nbsp;&nbsp;<%=formata_inteiro(r_orcamento.pce_prestacao_qtde)%>&nbsp;x&nbsp;<%=SIMBOLO_MONETARIO%>&nbsp;<%=formata_moeda(r_orcamento.pce_prestacao_valor)%>&nbsp;&nbsp;&nbsp;(<%=x_opcao_forma_pagamento(r_orcamento.pce_forma_pagto_prestacao)%>)&nbsp;&nbsp;vencendo a cada&nbsp;<%=formata_inteiro(r_orcamento.pce_prestacao_periodo)%>&nbsp;dias</span></td>
			  </tr>
			</table>
		  </td>
		</tr>
		<% elseif Cstr(r_orcamento.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA then %>
		<!--  PARCELADO SEM ENTRADA  -->
		<tr>
		  <td align="left">
			<table cellspacing="0" cellpadding="0" border="0">
			  <tr>
				<td align="left"><span class="C">1ª Prestação:&nbsp;&nbsp;<%=SIMBOLO_MONETARIO%>&nbsp;<%=formata_moeda(r_orcamento.pse_prim_prest_valor)%>&nbsp;&nbsp;&nbsp;(<%=x_opcao_forma_pagamento(r_orcamento.pse_forma_pagto_prim_prest)%>)&nbsp;&nbsp;vencendo após&nbsp;<%=formata_inteiro(r_orcamento.pse_prim_prest_apos)%>&nbsp;dias</span></td>
			  </tr>
			  <tr>
				<td align="left"><span class="C">Demais Prestações:&nbsp;&nbsp;<%=Cstr(r_orcamento.pse_demais_prest_qtde)%>&nbsp;x&nbsp;<%=SIMBOLO_MONETARIO%>&nbsp;<%=formata_moeda(r_orcamento.pse_demais_prest_valor)%>&nbsp;&nbsp;&nbsp;(<%=x_opcao_forma_pagamento(r_orcamento.pse_forma_pagto_demais_prest)%>)&nbsp;&nbsp;vencendo a cada&nbsp;<%=Cstr(r_orcamento.pse_demais_prest_periodo)%>&nbsp;dias</span></td>
			  </tr>
			</table>
		  </td>
		</tr>
		<% end if %>
	  </table>
	</td>
  </tr>
  <tr>
	<td class="MC" align="left"><p class="Rf">Informações Sobre Análise de Crédito</p>
	  <textarea name="c_forma_pagto" id="c_forma_pagto" class="PLLe notPrint" rows="<%=Cstr(MAX_LINHAS_FORMA_PAGTO)%>"
				style="width:642px;margin-left:2pt;"
				readonly tabindex=-1><%=r_orcamento.forma_pagto%></textarea>
		<span class="PLLe notVisible"><%
			s = substitui_caracteres(r_orcamento.forma_pagto,chr(13),"<br>")
			if s = "" then s = "&nbsp;"
			Response.Write s %></span>
	</td>
  </tr>
</table>
<% end if %>


<!-- ************   SEPARADOR   ************ -->
<table class="notPrint" width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc" align="left">&nbsp;</td></tr>
</table>

<table class="notPrint" width="649" cellpadding="0" cellspacing="0">
<tr><td align="right"><span class="Rc">
	<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
	<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
	</span>
</td></tr>
</table>

<!-- ************   BOTÕES   ************ -->
<table class="notPrint" width="649" cellpadding="0" cellspacing="0">
<tr>
	<%	if url_back <> "" then 
			s="resumo.asp" & "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))
		else 
			s="javascript:history.back()" 
			end if
	%>
	<td align="left"><a name="bVOLTAR" id="bVOLTAR" href="<%=s%>" title="volta para página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a>
	</td>
	<td align="center"><div name="dIMPRIME" id="dIMPRIME">
		<a name="bIMPRIME" id="bIMPRIME" href="javascript:fORCImprime(fORC)" title="vai para a página de impressão do pré-pedido em formulário contínuo">
		<img src="../botao/imprimir.gif" width="176" height="55" border="0"></a></div>
	</td>
	<td align="center">
		<% if IsOrcamentoCancelavel(r_orcamento.st_orcamento) then %>
		<div name='dREMOVE' id='dREMOVE'><a name="bREMOVE" id="bREMOVE" href="javascript:fORCRemove(fORC)" title="cancela este pré-pedido">
			<img src="../botao/remover.gif" width="176" height="55" border="0"></a>
		</div>
		<% end if %>
	</td>
	<td align="right">
		<div name="dMODIFICA" id="dMODIFICA"><a name="bMODIFICA" id="bMODIFICA" href="javascript:fORCModifica(fORC)" title="edita o pré-pedido">
			<img src="../botao/modificar.gif" width="176" height="55" border="0"></a>
		</div>
	</td>
</tr>
<tr>
	<td colspan='4' align="right">
		<% if IsOrcamentoAptoVirarPedido(r_orcamento.st_orcamento) then %>
		<div name="dVIRAPEDIDO" id="dVIRAPEDIDO"><a name="bVIRAPEDIDO" id="bVIRAPEDIDO" href="javascript:fORCVirarPedido(fORC)" title="transforma o pré-pedido em pedido">
			<img src="../botao/transforma.gif" width="176" height="55" border="0"></a>
		</div>
		<% end if %>
	</td>
</tr>
</table>

</form>


<!-- ************   DIRECIONA PARA CADASTRO DE CLIENTES   ************ -->
<form method="post" action="clienteedita.asp" id="fCLI" name="fCLI">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="cliente_selecionado" id="cliente_selecionado" value='<%=r_orcamento.id_cliente%>'>
<input type="hidden" name="operacao_selecionada" id="operacao_selecionada" value='<%=OP_CONSULTA%>'>
<input type="hidden" name="edicao_bloqueada" id="edicao_bloqueada" />
<input type="hidden" name="pagina_retorno" id="pagina_retorno" value='orcamento.asp?orcamento_selecionado=<%=orcamento_selecionado%>&url_back=X'>
</form>


</center>
<div id="divClienteConsultaView"><center><div id="divInternoClienteConsultaView"><img id="imgFechaDivClienteConsultaView" src="../imagem/close_button_32.png" title="clique para fechar o painel de consulta" /><iframe id="iframeClienteConsultaView"></iframe></div></center></div>
</body>

<% end if %>

</html>


<%
'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>