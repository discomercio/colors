<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/Braspag.asp"    -->
<!-- #include file = "../global/BraspagCS.asp"    -->

<%
'     =============================================
'	  P070PagtoMsgExec.asp
'     =============================================
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

'	A PÁGINA P080PagtoExec.asp EXECUTA A CHAMADA AO WEB SERVICE DA BRASPAG, QUE EVENTUALMENTE
'	PODE DEMORAR A RESPONDER OU ATÉ MESMO NÃO RESPONDER NUNCA.
'	NESSE CASO, A PÁGINA ANTERIOR É QUE CONTINUA SENDO EXIBIDA NO NAVEGADOR, AGUARDANDO
'	O CONTEÚDO DA NOVA PÁGINA SER ENVIADO. COMO ISSO PODE DEMORAR, FOI CRIADA ESTA PÁGINA
'	INTERMEDIÁRIA COM UMA MENSAGEM SOLICITANDO P/ O USUÁRIO AGUARDAR.

	On Error GoTo 0
	Err.Clear

	dim alerta
	alerta = ""

	dim s, usuario, loja, pedido_selecionado, pedido_com_sufixo_nsu, id_pedido_base

	usuario = BRASPAG_USUARIO_CLIENTE

	pedido_selecionado = ucase(Trim(request("pedido_selecionado")))
	if (pedido_selecionado = "") then Response.Redirect("aviso.asp?id=" & ERR_PEDIDO_NAO_ESPECIFICADO)
	s = normaliza_num_pedido(pedido_selecionado)
	if s <> "" then pedido_selecionado = s
	id_pedido_base = retorna_num_pedido_base(pedido_selecionado)

	pedido_com_sufixo_nsu = Trim(Request("pedido_com_sufixo_nsu"))
	if pedido_com_sufixo_nsu = "" then pedido_com_sufixo_nsu = pedido_selecionado

	dim cnpj_cpf_selecionado
	cnpj_cpf_selecionado = retorna_so_digitos(Request("cnpj_cpf_selecionado"))

	dim FingerPrint_SessionID
	FingerPrint_SessionID = Trim(Request("FingerPrint_SessionID"))

	dim c_qtde_cartoes, qtde_cartoes
	c_qtde_cartoes = Trim(Request("c_qtde_cartoes"))
	qtde_cartoes = converte_numero(c_qtde_cartoes)
	if qtde_cartoes = 0 then Response.Redirect("aviso.asp?id=" & ERR_QTDE_CARTOES_INVALIDA)

	dim c_fatura_telefone_pais
	c_fatura_telefone_pais = Trim(Request("c_fatura_telefone_pais"))

	dim i, vDadosCartao
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
		vDadosCartao(i).cartao_numero = decriptografa(Trim(Request(s_name)))
		s_name = "c_cartao_validade_mes_" & i
		vDadosCartao(i).cartao_validade_mes = retorna_so_digitos(Trim(Request(s_name)))
		s_name = "c_cartao_validade_ano_" & i
		vDadosCartao(i).cartao_validade_ano = retorna_so_digitos(Trim(Request(s_name)))
		s_name = "c_cartao_codigo_seguranca_" & i
		vDadosCartao(i).cartao_codigo_seguranca = decriptografa(Trim(Request(s_name)))
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
	<title><%=SITE_CLIENTE_TITULO_JANELA%></title>
	</head>


<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__SSL_JS%>" Language="JavaScript" Type="text/javascript"></script>

<% if alerta = "" then %>
<script language="JavaScript" type="text/javascript">
	setTimeout('fPAGTO.submit()', 100);
</script>
<% end if %>



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

<style type="text/css">
body::before
{
	content: '';
	border: none;
	margin-top: 0px;
	margin-bottom: 0px;
	padding: 0px;
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
<body>
<center>

<form id="fPAGTO" name="fPAGTO" method="post" action="P080PagtoExec.asp">
<input type="hidden" name="pedido_selecionado" value="<%=pedido_selecionado%>" />
<input type="hidden" name='cnpj_cpf_selecionado' value='<%=cnpj_cpf_selecionado%>'>
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
</form>

<br />
<br />

<table cellpadding="0" cellspacing="0">
	<tr>
	<td valign="bottom" align="right"><span style="color:orangered;font-weight:bold;font-style:italic;font-size:20pt;">Aguarde, processando a transação...</span></td>
	<td style="width:20px;" align="left">&nbsp;</td>
	<td align="left"><img src="../imagem/aguarde.gif"border="0"></td>
	</tr>
</table>

</center>

<% if SITE_CLIENTE_EXIBIR_LOGO_SSL then %>
<script language="JavaScript" type="text/javascript">
	logo_ssl_corner("../imagem/ssl/ssl_corner.gif");
</script>
<% end if %>

</body>
<% end if %>

</html>
