<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/Global.asp"    -->
<!-- #include file = "../global/Braspag.asp"    -->
<!-- #include file = "../global/BraspagCS.asp"    -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ===========================================
'	  P060PagtoReady.asp
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

	dim alerta
	alerta = ""
	
	dim s, usuario, loja, pedido_selecionado, pedido_com_sufixo_nsu, id_pedido_base
	usuario = Trim(Session("usuario_atual"))
	loja = Trim(Session("loja_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 
	If (loja = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

	pedido_selecionado = ucase(Trim(request("pedido_selecionado")))
	if (pedido_selecionado = "") then Response.Redirect("aviso.asp?id=" & ERR_PEDIDO_NAO_ESPECIFICADO)
	s = normaliza_num_pedido(pedido_selecionado)
	if s <> "" then pedido_selecionado = s
	id_pedido_base = retorna_num_pedido_base(pedido_selecionado)

	pedido_com_sufixo_nsu = Trim(Request("pedido_com_sufixo_nsu"))
	if pedido_com_sufixo_nsu = "" then pedido_com_sufixo_nsu = pedido_selecionado

	dim FingerPrint_SessionID
	FingerPrint_SessionID = Trim(Request("FingerPrint_SessionID"))

	dim c_qtde_cartoes, qtde_cartoes
	c_qtde_cartoes = Trim(Request("c_qtde_cartoes"))
	qtde_cartoes = converte_numero(c_qtde_cartoes)
	if qtde_cartoes = 0 then Response.Redirect("aviso.asp?id=" & ERR_QTDE_CARTOES_INVALIDA)

	dim c_fatura_telefone_pais
	c_fatura_telefone_pais = Trim(Request("c_fatura_telefone_pais"))

	dim i, j, vDadosCartao
	redim vDadosCartao(qtde_cartoes)
	for i = 1 to qtde_cartoes
		set vDadosCartao(i) = new cl_BraspagCS_DadosCartao_Checkout
		next

'	RECUPERA DADOS DO FORMULÁRIO
	dim s_name
	for i = 1 to qtde_cartoes
		s_name = "c_cartao_bandeira_" & i
		vDadosCartao(i).bandeira = Trim(Request(s_name))
		s_name = "c_cartao_valor_" & i
		vDadosCartao(i).valor_pagamento = Trim(Request(s_name))
		s_name = "c_opcao_parcelamento_" & i
		vDadosCartao(i).opcao_parcelamento = Trim(Request(s_name))
		s_name = "c_cartao_nome_" & i
		vDadosCartao(i).titular_nome = Trim(Request(s_name))
		s_name = "c_cartao_cpf_cnpj_" & i
		vDadosCartao(i).titular_cpf_cnpj = retorna_so_digitos(Trim(Request(s_name)))
		s_name = "c_cartao_numero_" & i
		vDadosCartao(i).cartao_numero = retorna_so_digitos(Trim(Request(s_name)))
		s_name = "c_cartao_validade_mes_" & i
		vDadosCartao(i).cartao_validade_mes = retorna_so_digitos(Trim(Request(s_name)))
		s_name = "c_cartao_validade_ano_" & i
		vDadosCartao(i).cartao_validade_ano = retorna_so_digitos(Trim(Request(s_name)))
		s_name = "c_cartao_codigo_seguranca_" & i
		vDadosCartao(i).cartao_codigo_seguranca = retorna_so_digitos(Trim(Request(s_name)))
		s_name = "c_cartao_proprio_" & i
		vDadosCartao(i).cartao_proprio = Trim(Request(s_name))
		s_name = "c_fatura_end_logradouro_" & i
		vDadosCartao(i).fatura_end_logradouro = Trim(Request(s_name))
		s_name = "c_fatura_end_numero_" & i
		vDadosCartao(i).fatura_end_numero = Trim(Request(s_name))
		s_name = "c_fatura_end_complemento_" & i
		vDadosCartao(i).fatura_end_complemento = Trim(Request(s_name))
		s_name = "c_fatura_end_bairro_" & i
		vDadosCartao(i).fatura_end_bairro = Trim(Request(s_name))
		s_name = "c_fatura_end_cidade_" & i
		vDadosCartao(i).fatura_end_cidade = Trim(Request(s_name))
		s_name = "c_fatura_end_uf_" & i
		vDadosCartao(i).fatura_end_uf = Trim(Request(s_name))
		s_name = "c_fatura_end_cep_" & i
		vDadosCartao(i).fatura_end_cep = retorna_so_digitos(Trim(Request(s_name)))
		s_name = "c_fatura_telefone_ddd_" & i
		vDadosCartao(i).fatura_tel_ddd = retorna_so_digitos(Trim(Request(s_name)))
		s_name = "c_fatura_telefone_numero_" & i
		vDadosCartao(i).fatura_tel_numero = retorna_so_digitos(Trim(Request(s_name)))
		next
	
'	PROCESSAMENTO PARA OBTER CAMPOS AUXILIARES
	dim v
	for i = 1 to qtde_cartoes
	'	VALOR DO PAGAMENTO NESTE CARTÃO
		vDadosCartao(i).vl_pagamento = converte_numero(vDadosCartao(i).valor_pagamento)
	'	OPÇÃO DE PARCELAMENTO
		if vDadosCartao(i).opcao_parcelamento = "0" then
		'	À VISTA
			vDadosCartao(i).codigo_produto = "0"
			vDadosCartao(i).qtde_parcelas = 1
		elseif InStr(vDadosCartao(i).opcao_parcelamento, "PL|") <> 0 then
		'	PARCELADO ESTABELECIMENTO (LOJA)
			vDadosCartao(i).codigo_produto = "1"
			v = Split(vDadosCartao(i).opcao_parcelamento, "|")
			vDadosCartao(i).qtde_parcelas = converte_numero(v(Ubound(v)))
			if vDadosCartao(i).qtde_parcelas <= 0 then
				alerta=texto_add_br(alerta)
				alerta=alerta & "Quantidade de parcelas inválida no cartão de bandeira " & BraspagDescricaoBandeira(vDadosCartao(i).bandeira)
				end if
		elseif InStr(vDadosCartao(i).opcao_parcelamento, "PC|") <> 0 then
		'	PARCELADO CARTÃO (PELO EMISSOR DO CARTÃO)
			vDadosCartao(i).codigo_produto = "2"
			v = Split(vDadosCartao(i).opcao_parcelamento, "|")
			vDadosCartao(i).qtde_parcelas = converte_numero(v(Ubound(v)))
			if vDadosCartao(i).qtde_parcelas <= 0 then
				alerta=texto_add_br(alerta)
				alerta=alerta & "Quantidade de parcelas inválida no cartão de bandeira " & BraspagDescricaoBandeira(vDadosCartao(i).bandeira)
				end if
		else
			alerta=texto_add_br(alerta)
			alerta=alerta & "Opção de parcelamento inválida no cartão de bandeira " & BraspagDescricaoBandeira(vDadosCartao(i).bandeira)
			end if
	'	DESCRIÇÃO DO PARCELAMENTO
		vDadosCartao(i).descricao_parcelamento = BraspagCSDescricaoParcelamento(vDadosCartao(i).codigo_produto, vDadosCartao(i).qtde_parcelas, vDadosCartao(i).vl_pagamento)
		next

'	CONSISTÊNCIA (BANDEIRA E VALOR IGUAIS)
	if qtde_cartoes > 1 then
		for i = 1 to qtde_cartoes
			for j = 1 to (i-1)
				if (vDadosCartao(i).bandeira = vDadosCartao(j).bandeira) And (vDadosCartao(i).vl_pagamento = vDadosCartao(j).vl_pagamento) then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Os cartões " & Cstr(j) & " e " & Cstr(i) & " são da mesma bandeira (" & BraspagDescricaoBandeira(vDadosCartao(i).bandeira) & ") e estão sendo usados para pagar um valor idêntico (" & SIMBOLO_MONETARIO & " " & formata_moeda(vDadosCartao(i).vl_pagamento) & ")."
					alerta=texto_add_br(alerta)
					alerta=alerta & "Devido a uma limitação do gateway de pagamentos, por favor, escolha um valor de pagamento diferente para cada cartão de mesma bandeira."
				end if
			next
		next
	end if

'	CONSISTÊNCIA GERAL DOS DADOS
	if alerta = "" then
		for i = 1 to qtde_cartoes
			if Trim("" & vDadosCartao(i).bandeira) = "" then
				alerta=texto_add_br(alerta)
				s = "Selecione a bandeira do cartão"
				if qtde_cartoes > 1 then s = Cstr(i) & "º cartão: " & s
				alerta=alerta & s
				end if
			if vDadosCartao(i).vl_pagamento = 0 then
				alerta=texto_add_br(alerta)
				s = "Valor do pagamento é inválido"
				if qtde_cartoes > 1 then s = Cstr(i) & "º cartão: " & s
				alerta=alerta & s
				end if
			if Trim("" & vDadosCartao(i).opcao_parcelamento) = "" then
				alerta=texto_add_br(alerta)
				s = "Opção de parcelamento é inválida"
				if qtde_cartoes > 1 then s = Cstr(i) & "º cartão: " & s
				alerta=alerta & s
				end if
			if vDadosCartao(i).qtde_parcelas = 0 then
				alerta=texto_add_br(alerta)
				s = "Quantidade de parcelas é inválida"
				if qtde_cartoes > 1 then s = Cstr(i) & "º cartão: " & s
				alerta=alerta & s
				end if
			if Trim("" & vDadosCartao(i).titular_nome) = "" then
				alerta=texto_add_br(alerta)
				s = "Informe o nome do titular do cartão"
				if qtde_cartoes > 1 then s = Cstr(i) & "º cartão: " & s
				alerta=alerta & s
				end if
			if Trim("" & vDadosCartao(i).titular_cpf_cnpj) = "" then
				alerta=texto_add_br(alerta)
				s = "Informe o CPF/CNPJ do titular do cartão"
				if qtde_cartoes > 1 then s = Cstr(i) & "º cartão: " & s
				alerta=alerta & s
				end if
			if Trim("" & vDadosCartao(i).cartao_numero) = "" then
				alerta=texto_add_br(alerta)
				s = "Informe o número do cartão"
				if qtde_cartoes > 1 then s = Cstr(i) & "º cartão: " & s
				alerta=alerta & s
				end if
			if Trim("" & vDadosCartao(i).cartao_validade_mes) = "" then
				alerta=texto_add_br(alerta)
				s = "Informe o mês da validade do cartão"
				if qtde_cartoes > 1 then s = Cstr(i) & "º cartão: " & s
				alerta=alerta & s
				end if
			if Trim("" & vDadosCartao(i).cartao_validade_ano) = "" then
				alerta=texto_add_br(alerta)
				s = "Informe o ano da validade do cartão"
				if qtde_cartoes > 1 then s = Cstr(i) & "º cartão: " & s
				alerta=alerta & s
				end if
			if Trim("" & vDadosCartao(i).cartao_codigo_seguranca) = "" then
				alerta=texto_add_br(alerta)
				s = "Informe o código de segurança do cartão"
				if qtde_cartoes > 1 then s = Cstr(i) & "º cartão: " & s
				alerta=alerta & s
				end if
			if Trim("" & vDadosCartao(i).cartao_proprio) = "" then
				alerta=texto_add_br(alerta)
				s = "Informe se o cartão é próprio ou de terceiro"
				if qtde_cartoes > 1 then s = Cstr(i) & "º cartão: " & s
				alerta=alerta & s
				end if
			if Trim("" & vDadosCartao(i).fatura_end_logradouro) = "" then
				alerta=texto_add_br(alerta)
				s = "Informe o endereço da fatura do cartão"
				if qtde_cartoes > 1 then s = Cstr(i) & "º cartão: " & s
				alerta=alerta & s
				end if
			if Trim("" & vDadosCartao(i).fatura_end_numero) = "" then
				alerta=texto_add_br(alerta)
				s = "Informe o número no endereço da fatura do cartão"
				if qtde_cartoes > 1 then s = Cstr(i) & "º cartão: " & s
				alerta=alerta & s
				end if
			if Trim("" & vDadosCartao(i).fatura_end_bairro) = "" then
				alerta=texto_add_br(alerta)
				s = "Informe o bairro do endereço da fatura do cartão"
				if qtde_cartoes > 1 then s = Cstr(i) & "º cartão: " & s
				alerta=alerta & s
				end if
			if Trim("" & vDadosCartao(i).fatura_end_cidade) = "" then
				alerta=texto_add_br(alerta)
				s = "Informe a cidade do endereço da fatura do cartão"
				if qtde_cartoes > 1 then s = Cstr(i) & "º cartão: " & s
				alerta=alerta & s
				end if
			if Trim("" & vDadosCartao(i).fatura_end_uf) = "" then
				alerta=texto_add_br(alerta)
				s = "Informe a UF do endereço da fatura do cartão"
				if qtde_cartoes > 1 then s = Cstr(i) & "º cartão: " & s
				alerta=alerta & s
				end if
			if Trim("" & vDadosCartao(i).fatura_end_cep) = "" then
				alerta=texto_add_br(alerta)
				s = "Informe o CEP do endereço da fatura do cartão"
				if qtde_cartoes > 1 then s = Cstr(i) & "º cartão: " & s
				alerta=alerta & s
				end if
			if Trim("" & vDadosCartao(i).fatura_tel_ddd) = "" then
				alerta=texto_add_br(alerta)
				s = "Informe o DDD do telefone"
				if qtde_cartoes > 1 then s = Cstr(i) & "º cartão: " & s
				alerta=alerta & s
				end if
			if Trim("" & vDadosCartao(i).fatura_tel_numero) = "" then
				alerta=texto_add_br(alerta)
				s = "Informe o número do telefone"
				if qtde_cartoes > 1 then s = Cstr(i) & "º cartão: " & s
				alerta=alerta & s
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
	<title>LOJA</title>
</head>


<% if False then 'APENAS P/ HABILITAR O INTELLISENSE DURANTE O DESENVOLVIMENTO!! %>
<script src="../Global/jquery.js" language="JavaScript" type="text/javascript"></script>
<% end if %>

<script src="<%=URL_FILE__JQUERY%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
window.name="Loja";

function fPAGTOConclui(f) {
	f.action = "P070PagtoMsgExec.asp";
	dPAGTO.style.visibility="hidden";
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

<link href="<%=URL_FILE__E_CSS%>" rel="stylesheet" type="text/css">
<link href="<%=URL_FILE__EPRINTER_CSS%>" rel="stylesheet" type="text/css" media="print">
<link href="<%=URL_FILE__EGWBP_CSS%>" rel="stylesheet" type="text/css">

<style type="text/css">
.tdColMargin
{
	width:15px;
}
</style>

<% if alerta <> "" then %>
<!-- ************************************************************ -->
<!-- **********  PÁGINA PARA EXIBIR MENSAGENS DE ERRO  ********** -->
<!-- ************************************************************ -->
<body onload="bVOLTAR.focus()">
<center>
<br>
<!--  T E L A  -->
<p class="T">A V I S O</p>
<div class="MtAlerta" style="width:600px;font-weight:bold;" align="center"><p style='margin:5px 2px 5px 2px;'><%=alerta%></p></div>
<br><br>
<p class="TracoBottom"></p>
<table cellspacing="0">
<tr>
	<td align="center"><a name="bVOLTAR" href="javascript:history.back()"><img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
</tr>
</table>
</center>
</body>




<% else %>
<!-- ********************************************************************* -->
<!-- **********  PÁGINA PARA EXIBIR RESUMO DO PAGAMENTO  ***************** -->
<!-- ********************************************************************* -->
<body>
<center>

<form id="fPAGTO" name="fPAGTO" method="post" >
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="pedido_selecionado" value="<%=pedido_selecionado%>" />
<input type="hidden" name="c_fatura_telefone_pais" value="55" />
<input type="hidden" name="pedido_com_sufixo_nsu" value="<%=pedido_com_sufixo_nsu%>" />
<input type="hidden" name="FingerPrint_SessionID" value="<%=FingerPrint_SessionID%>" />
<input type="hidden" name="c_qtde_cartoes" value="<%=c_qtde_cartoes%>" />
<%	for i = 1 to qtde_cartoes
		Response.Write "<input type=""hidden"" name=""c_cartao_bandeira_" & i & """ value=""" & vDadosCartao(i).bandeira & """ />" & chr(13)
		Response.Write "<input type=""hidden"" name=""c_cartao_valor_" & i & """ value=""" & vDadosCartao(i).valor_pagamento & """ />" & chr(13)
		Response.Write "<input type=""hidden"" name=""c_opcao_parcelamento_" & i & """ value=""" & vDadosCartao(i).opcao_parcelamento & """ />" & chr(13)
		Response.Write "<input type=""hidden"" name=""c_cartao_nome_" & i & """ value=""" & vDadosCartao(i).titular_nome & """ />" & chr(13)
		Response.Write "<input type=""hidden"" name=""c_cartao_cpf_cnpj_" & i & """ value=""" & vDadosCartao(i).titular_cpf_cnpj & """ />" & chr(13)
		Response.Write "<input type=""hidden"" name=""c_cartao_numero_" & i & """ value=""" & criptografa(vDadosCartao(i).cartao_numero) & """ />" & chr(13)
		Response.Write "<input type=""hidden"" name=""c_cartao_validade_mes_" & i & """ value=""" & vDadosCartao(i).cartao_validade_mes & """ />" & chr(13)
		Response.Write "<input type=""hidden"" name=""c_cartao_validade_ano_" & i & """ value=""" & vDadosCartao(i).cartao_validade_ano & """ />" & chr(13)
		Response.Write "<input type=""hidden"" name=""c_cartao_codigo_seguranca_" & i & """ value=""" & criptografa(vDadosCartao(i).cartao_codigo_seguranca) & """ />" & chr(13)
		Response.Write "<input type=""hidden"" name=""c_cartao_proprio_" & i & """ value=""" & vDadosCartao(i).cartao_proprio & """ />" & chr(13)
		Response.Write "<input type=""hidden"" name=""c_fatura_end_logradouro_" & i & """ value=""" & vDadosCartao(i).fatura_end_logradouro & """ />" & chr(13)
		Response.Write "<input type=""hidden"" name=""c_fatura_end_numero_" & i & """ value=""" & vDadosCartao(i).fatura_end_numero & """ />" & chr(13)
		Response.Write "<input type=""hidden"" name=""c_fatura_end_complemento_" & i & """ value=""" & vDadosCartao(i).fatura_end_complemento & """ />" & chr(13)
		Response.Write "<input type=""hidden"" name=""c_fatura_end_bairro_" & i & """ value=""" & vDadosCartao(i).fatura_end_bairro & """ />" & chr(13)
		Response.Write "<input type=""hidden"" name=""c_fatura_end_cidade_" & i & """ value=""" & vDadosCartao(i).fatura_end_cidade & """ />" & chr(13)
		Response.Write "<input type=""hidden"" name=""c_fatura_end_uf_" & i & """ value=""" & vDadosCartao(i).fatura_end_uf & """ />" & chr(13)
		Response.Write "<input type=""hidden"" name=""c_fatura_end_cep_" & i & """ value=""" & vDadosCartao(i).fatura_end_cep & """ />" & chr(13)
		Response.Write "<input type=""hidden"" name=""c_fatura_telefone_ddd_" & i & """ value=""" & vDadosCartao(i).fatura_tel_ddd & """ />" & chr(13)
		Response.Write "<input type=""hidden"" name=""c_fatura_telefone_numero_" & i & """ value=""" & vDadosCartao(i).fatura_tel_numero & """ />" & chr(13)
		next
%>

<table cellspacing="0" width="649" style="border-bottom:1px solid black">
<tr>
	<td align="center" valign="bottom"><img src="../imagem/<%=BRASPAG_LOGOTIPO_LOJA%>"></td>
</tr>
</table>

<!--  EXIBE RESUMO DO PAGAMENTO  -->
<br>
<br>
<table class="Qx" cellspacing="0" border="0">
	<!-- RESUMO DO PAGAMENTO -->
	<tr>
		<td colspan="6" class="MT" style="border-width:2px;border-color:black;padding:4px;" align="center"><span class="STP" style="font-size:14pt;">Resumo do Pagamento</span></td>
	</tr>
	<% if qtde_cartoes > 1 then %>
	<tr><td colspan="6" style="height:30px;" align="left"></td></tr>
	<% else %>
	<tr><td colspan="6" style="height:15px;" align="left"></td></tr>
	<% end if %>
	<% for i = 1 to qtde_cartoes %>
	<% if qtde_cartoes > 1 then %>
	<tr><td colspan="6" align="center"><span class="PLTd" style="font-size:14pt;"><%=i%>º CARTÃO</span></td></tr>
	<% end if %>
	<tr><td colspan="6" class="MC ME MD" style="height:6px;" align="left"></td></tr>
	<tr bgcolor="#FFFFFF">
		<td class="tdColMargin ME">&nbsp;</td>
		<td rowspan="3" align="center" valign="middle"><img src="../Imagem/Braspag/<%=BraspagObtemNomeArquivoLogoOpcao(vDadosCartao(i).bandeira)%>" border="0" /></td>
		<td rowspan="3" style="width:12px;"></td>
		<td align="right"><span class="PLTd">BANDEIRA</span></td>
		<td align="left"><span class="CARDSpanPgInfo" style="margin-left:6px;"><%=Ucase(BraspagDescricaoBandeira(vDadosCartao(i).bandeira))%></span></td>
		<td class="tdColMargin MD">&nbsp;</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td class="tdColMargin ME">&nbsp;</td>
		<td align="right"><span class="PLTd">Nº CARTÃO</span></td>
		<td align="left"><span class="CARDSpanPgInfo" style="margin-left:6px;"><%=BraspagCSProtegeNumeroCartao(vDadosCartao(i).cartao_numero)%></span></td>
		<td class="tdColMargin MD">&nbsp;</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td class="tdColMargin ME">&nbsp;</td>
		<td align="right"><span class="PLTd">VALOR</span></td>
		<td align="left"><span class="CARDSpanPgInfo" style="margin-left:6px;"><%=vDadosCartao(i).descricao_parcelamento%></span></td>
		<td class="tdColMargin MD">&nbsp;</td>
	</tr>
	<tr><td colspan="6" class="MB ME MD" style="height:6px;" align="left"></td></tr>
	<% if i < qtde_cartoes then %>
	<tr><td colspan="6" style="height:30px;" align="left"></td></tr>
	<% end if %>
	<% next %>
</table>
<br />



<!-- ************   SEPARADOR   ************ -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
<td class="Rc" align="left">&nbsp;</td>
</tr>
</table>


<!-- ************   BOTÕES   ************ -->
<table class="notPrint" width="649" cellpadding="0" cellspacing="0">
<tr>
	<td align="left"><a name="bVOLTAR" href="javascript:history.back();" title="volta para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a>
	</td>
	<td align="right">
		<div name="dPAGTO" id="dPAGTO"><a name="bPAGTO" href="javascript:fPAGTOConclui(fPAGTO)" title="efetua o pagamento">
			<img src="../botao/pagar.gif" width="176" height="55" border="0"></a></div>
	</td>
</tr>
</table>

</form>


</center>

<%' QUANDO O ENVIO É FEITO NO AMBIENTE 'CALL CENTER', OS SCRIPTS DO ANTI-FRAUDE NÃO DEVEM SER EXECUTADOS %>

</body>

<% end if %>

</html>
