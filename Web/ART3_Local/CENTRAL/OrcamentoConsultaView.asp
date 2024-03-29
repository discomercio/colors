<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ===========================================
'	  OrcamentoConsultaView.asp
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
'
'
'	REVISADO P/ IE10


	On Error GoTo 0
	Err.Clear

	dim s, usuario, orcamento_selecionado, orcamento_selecionado_inicial, pagina_retorno, s_url
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then
		usuario = Trim(Request("usuario"))
		Session("usuario_atual") = usuario
		end if
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO)

	orcamento_selecionado = ucase(Trim(request("orcamento_selecionado")))
	if (orcamento_selecionado = "") then Response.Redirect("aviso.asp?id=" & ERR_ORCAMENTO_NAO_ESPECIFICADO)
	s = normaliza_num_orcamento(orcamento_selecionado)
	if s <> "" then orcamento_selecionado=s
	if Len(orcamento_selecionado) > TAM_MAX_ID_ORCAMENTO then Response.Redirect("aviso.asp?id=" & ERR_ORCAMENTO_INVALIDO)
	
'	MEMORIZA O OR�AMENTO SELECIONADO INICIALMENTE P/ PODER RETORNAR A ELE
	orcamento_selecionado_inicial = Trim(Request("orcamento_selecionado_inicial"))
	
	pagina_retorno = Trim(Request("pagina_retorno"))
	
	dim i, n, s_fabricante, s_produto, s_descricao, s_descricao_html, s_obs, s_qtde, s_preco_lista, s_desc_dado, s_vl_unitario
	dim s_vl_TotalItem, m_TotalItem, m_TotalItemComRA, m_TotalDestePedido, m_TotalDestePedidoComRA
	dim s_preco_NF, m_total_NF, m_total_RA
	
'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	dim max_qtde_itens
	max_qtde_itens = obtem_parametro_PedidoItem_MaxQtdeItens

	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))
	if s_lista_operacoes_permitidas = "" then
		s_lista_operacoes_permitidas = obtem_operacoes_permitidas_usuario(cn, usuario)
		Session("lista_operacoes_permitidas") = s_lista_operacoes_permitidas
		end if
	
	dim s_aux, s2, s3, s4, r_loja, r_cliente
	dim r_orcamento, v_item, alerta, msg_erro
	alerta=""
	if Not le_orcamento(orcamento_selecionado, r_orcamento, msg_erro) then 
		alerta = msg_erro
	else
		if Not le_orcamento_item(orcamento_selecionado, v_item, msg_erro) then alerta = msg_erro
		'Assegura que dados cadastrados anteriormente sejam exibidos corretamente, mesmo se o par�metro da quantidade m�xima de itens tiver sido reduzido
		if VectorLength(v_item) > max_qtde_itens then max_qtde_itens = VectorLength(v_item)
		end if

	dim r_pedido
	if alerta = "" then
		if r_orcamento.st_orc_virou_pedido = 1 then
			if Not le_pedido(r_orcamento.pedido, r_pedido, msg_erro) then alerta = msg_erro
			end if
		end if
	
	if alerta = "" then
		if Not orcamento_calcula_total_NF_e_RA(orcamento_selecionado, m_total_NF, m_total_RA, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
		end if

	dim strTextoIndicador
	dim r_orcamentista_e_indicador
	if alerta = "" then
		call le_orcamentista_e_indicador(r_orcamento.orcamentista, r_orcamentista_e_indicador, msg_erro)
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
	<title>CENTRAL<%=MontaNumOrcamentoExibicaoTitleBrowser(orcamento_selecionado)%></title>
	</head>


<% if False then 'APENAS P/ HABILITAR O INTELLISENSE DURANTE O DESENVOLVIMENTO!! %>
<script src="../Global/jquery.js" language="JavaScript" type="text/javascript"></script>
<% end if %>

<script src="<%=URL_FILE__JQUERY%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
	$(document).ready(function() {
		window.status = "";
		var topo = $('#divConsultaOrcamento').offset().top - parseFloat($('#divConsultaOrcamento').css('margin-top').replace(/auto/, 0)) - parseFloat($('#divConsultaOrcamento').css('padding-top').replace(/auto/, 0));
		$('#divConsultaOrcamento').addClass('divFixo');
	});
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
	fPED.pedido_selecionado.value = s_pedido;
	fPED.submit(); 
}

function fCLIConsulta() {
	window.status = "Aguarde ...";
	fCLI.submit();
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
<link href="<%=URL_FILE__ESCREEN_CSS%>" rel="stylesheet" type="text/css" media="screen">

<style type="text/css">
</style>


<% if alerta <> "" then %>
<!-- ************************************************************ -->
<!-- **********  P�GINA PARA EXIBIR MENSAGENS DE ERRO  ********** -->
<!-- ************************************************************ -->
<body>
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
<!-- **********  P�GINA PARA EXIBIR O OR�AMENTO  ***************** -->
<!-- ************************************************************* -->
<body link="#ffffff" alink="#ffffff" vlink="#ffffff">

<center>

<form method="post" action="PedidoConsultaView.asp" id="fPED" name="fPED">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name='pedido_selecionado' id="pedido_selecionado" value=''>
<input type="hidden" name="usuario" id="usuario" value='<%=usuario%>'>
<input type="hidden" name='pagina_retorno' value='OrcamentoConsultaView.asp?orcamento_selecionado=<%=orcamento_selecionado%>&orcamento_selecionado_inicial=<%=orcamento_selecionado_inicial%>&usuario=<%=usuario%>&<%=MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>'>
</form>

<form id="fORC" name="fORC" method="post">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name='orcamento_selecionado' id="orcamento_selecionado" value='<%=orcamento_selecionado%>'>
<input type="hidden" name="orcamento_selecionado_inicial" id="orcamento_selecionado_inicial" value='<%=orcamento_selecionado_inicial%>'>
<input type="hidden" name="usuario" id="usuario" value='<%=usuario%>'>


<!--  I D E N T I F I C A � � O   D O   O R � A M E N T O -->
<%=MontaHeaderIdentificacaoOrcamento(orcamento_selecionado, r_orcamento, 649)%>
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
	strTextoIndicador = ""
	if r_orcamento.orcamentista <> "" then
		strTextoIndicador = r_orcamento.orcamentista
		if r_orcamentista_e_indicador.desempenho_nota <> "" then
			strTextoIndicador = strTextoIndicador & " (" & r_orcamentista_e_indicador.desempenho_nota & ")"
			end if
		end if
%>
	<td class="MD" align="left"><p class="Rf">LOJA</p><p class="C"><%=s%>&nbsp;</p></td>
	<td width="145" class="MD" align="left"><p class="Rf">OR�AMENTISTA</p><p class="C"><%=strTextoIndicador%>&nbsp;</p></td>
	<td width="145" align="left"><p class="Rf">VENDEDOR</p><p class="C"><%=r_orcamento.vendedor%>&nbsp;</p></td>
	</tr>
	</table>
	
<br>

<!--  CLIENTE   -->
<table width="649" class="Q" cellspacing="0">
	<tr>
<%	s = ""
	set r_cliente = New cl_CLIENTE
	if x_cliente_bd(r_orcamento.id_cliente, r_cliente) then
	
    'le as vari�veis da origem certa: ou do pedido ou do cliente, todas comecam com cliente__
    dim cliente__tipo, cliente__cnpj_cpf, cliente__rg, cliente__ie, cliente__nome
    dim cliente__endereco, cliente__endereco_numero, cliente__endereco_complemento, cliente__bairro, cliente__cidade, cliente__uf, cliente__cep
    dim cliente__tel_res, cliente__ddd_res, cliente__tel_com, cliente__ddd_com, cliente__ramal_com, cliente__tel_cel, cliente__ddd_cel
    dim cliente__tel_com_2, cliente__ddd_com_2, cliente__ramal_com_2, cliente__email, cliente__email_xml, cliente__produtor_rural_status, cliente__contribuinte_icms_status

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
    cliente__email_xml = r_cliente.email_xml
	cliente__produtor_rural_status = r_cliente.produtor_rural_status
	cliente__contribuinte_icms_status = r_cliente.contribuinte_icms_status

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
        cliente__email_xml = r_orcamento.endereco_email_xml
		cliente__produtor_rural_status = r_orcamento.endereco_produtor_rural_status
		cliente__contribuinte_icms_status = r_orcamento.endereco_contribuinte_icms_status
        end if

%>
<%	if cliente__tipo = ID_PF then s_aux="CPF" else s_aux="CNPJ"
	s = cnpj_cpf_formata(cliente__cnpj_cpf) 
%>
		<td align="left" width="33%" class="MD"><p class="Rf"><%=s_aux%></p>
		
			<a href='javascript:fCLIConsulta();' title='clique para consultar o cadastro do cliente'><p class="C"><%=s%>&nbsp;</p></a>
		
		</td>



	<% if cliente__tipo = ID_PF then %>
		<td align="left" width="33%" class="MD"><p class="Rf">RG</p><p class="C"><%=Trim(cliente__rg)%>&nbsp;</p></td>
		<% 
		s_aux = ""
		if converte_numero(Trim(cliente__produtor_rural_status)) = converte_numero(COD_ST_CLIENTE_PRODUTOR_RURAL_SIM) then
			s = converte_numero(cliente__contribuinte_icms_status)
			if s = converte_numero(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO) then
				s_aux = "Sim (N�o contribuinte)"
			elseif s = converte_numero(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM) then
				s_aux = "Sim (IE: " & cliente__ie & ")"
			elseif s = converte_numero(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_ISENTO) then
				s_aux = "Sim (Isento)"
			end if
		elseif cliente__produtor_rural_status = converte_numero(COD_ST_CLIENTE_PRODUTOR_RURAL_NAO) then
			s_aux = "N�o"
		end if
		%>
		<td align="left" width="33%" class="MD"><p class="Rf">PRODUTOR RURAL</p><p class="C"><%=s_aux%>&nbsp;</p></td>
	<% else %>

		<td width="33%" class="MD" align="left"><p class="Rf">IE</p><p class="C"><%=Trim(cliente__ie)%>&nbsp;</p></td>
		<% 
			s_aux = ""
			s = converte_numero(cliente__contribuinte_icms_status)
			if s = converte_numero(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO) then
				s_aux = "N�o"
			elseif s = converte_numero(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM) then
				s_aux = "Sim"
			elseif s = converte_numero(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_ISENTO) then
				s_aux = "Isento"
			end if            
		%>
		<td width="33%" align="left" class="MD"><p class="Rf">CONTRIBUINTE ICMS</p><p class="C"><%=s_aux%>&nbsp;</p></td>

	<% end if %>

		<td align="center" valign="middle" style="width:22px;" class="MB"><a href='javascript:fCLIConsulta();' title="clique para consultar o cadastro do cliente"><img id="imgClienteConsultaView" src="../imagem/doc_preview_22.png" /></a></td>
	</tr>
<%
		
		if Trim(cliente__nome) <> "" then
			s = Trim(cliente__nome)
			end if
		end if
	
	if cliente__tipo = ID_PF then s_aux="NOME DO CLIENTE" else s_aux="RAZ�O SOCIAL DO CLIENTE"
%>
    <tr>
	<td class="MC" align="left" colspan="3"><p class="Rf"><%=s_aux%></p>
	
		<a href='javascript:fCLIConsulta();' title='clique para consultar o cadastro do cliente'><p class="C"><%=s%>&nbsp;</p></a>
	
		</td>
	</tr>
	</table>
<!--  ENDERE�O DO CLIENTE  -->
<table width="649" class="QS" cellspacing="0">
	<tr>
<%	
	s = formata_endereco(cliente__endereco, cliente__endereco_numero, cliente__endereco_complemento, cliente__bairro, cliente__cidade, cliente__uf, cliente__cep)
%>		
		<td align="left"><p class="Rf">ENDERE�O</p><p class="C"><%=s%>&nbsp;</p></td>
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
<table width="649" class="QS" cellspacing="0">
	<tr>
		<td align="left" class="MD" width="50%"><p class="Rf">E-MAIL</p><p class="C"><%=Trim(cliente__email)%>&nbsp;</p></td>
		<td align="left" width="50%"><p class="Rf">E-MAIL (XML)</p><p class="C"><%=Trim(cliente__email_xml)%>&nbsp;</p></td>
	</tr>
</table>

<!--  ENDERE�O DE ENTREGA  -->
<%	
    s = pedido_formata_endereco_entrega(r_orcamento, r_cliente)
%>		
<table width="649" class="QS" cellspacing="0" style="table-layout:fixed">
	<tr>
		<td align="left"><p class="Rf">ENDERE�O DE ENTREGA</p><p class="C"><%=s%>&nbsp;</p></td>
	</tr>
    <%	if r_orcamento.EndEtg_cod_justificativa <> "" then %>	
    <tr>
		<td align="left" style="word-wrap:break-word"><p class="C" ><%=obtem_descricao_tabela_t_codigo_descricao(GRUPO_T_CODIGO_DESCRICAO__ENDETG_JUSTIFICATIVA,r_orcamento.EndEtg_cod_justificativa)%>&nbsp;</p></td>
	</tr>
    <%end if %>
</table>



<!--  R E L A � � O   D E   P R O D U T O S  -->
<br>
<br>
<table class="Qx" cellspacing="0">
	<tr bgColor="#FFFFFF">
	<td class="MB" align="left" valign="bottom"><span class="PLTe">Fabr</span></td>
	<td class="MB" align="left" valign="bottom"><span class="PLTe">Produto</span></td>
	<td class="MB" align="left" valign="bottom"><span class="PLTe" style="width:287px;">Descri��o/Observa��es</span></td>
	<td class="MB" align="right" valign="bottom"><span class="PLTd">Qtd</span></td>
	<% if (r_orcamento.permite_RA_status = 1) Or blnTemRA then %>
	<td class="MB" align="right" valign="bottom"><span class="PLTd">Pre�o</span></td>
	<% end if %>
	<td class="MB" align="right" valign="bottom"><span class="PLTd">VL Lista</span></td>
	<td class="MB" align="right" valign="bottom"><span class="PLTd">Desc</span></td>
	<td class="MB" align="right" valign="bottom"><span class="PLTd">VL Venda</span></td>
	<td class="MB" align="right" valign="bottom"><span class="PLTd">VL Total</span></td>
	</tr>

<% m_TotalDestePedido=0
   m_TotalDestePedidoComRA=0
   n = Lbound(v_item)-1
   for i=1 to max_qtde_itens
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

'	A VERS�O 5.0 DO IE N�O DESENHA AS MARGENS SE O SPAN N�O POSSUIR CONTE�DO
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
<!--  TRATA VERS�O ANTIGA DA FORMA DE PAGAMENTO   -->
<br>
<table class="Q" cellspacing="0" style="width:649px;">
	<tr>
		<td class="MB" colspan="5" align="left"><p class="Rf">Observa��es I</p>
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
		<td class="MB" colspan="5" align="left"><p class="Rf">Observa��es II</p>
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
				s = "N�O"
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
				s = "N�O"
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
				s = "N�O"
			elseif Cstr(r_orcamento.InstaladorInstalaStatus) = Cstr(COD_INSTALADOR_INSTALA_SIM) then
				s = "SIM"
			else
				s = ""
				end if
		
			if s="" then s="&nbsp;"
		%>
		<p class="C" style="margin-top:3px;"><%=s%></p>
		</td>
		<td class="MB" nowrap align="left" valign="top"><p class="Rf">Garantia Indicador</p>
		<% 	if Cstr(r_orcamento.GarantiaIndicadorStatus) = Cstr(COD_GARANTIA_INDICADOR_STATUS__NAO) then
				s = "N�O"
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
<!--  TRATA NOVA VERS�O DA FORMA DE PAGAMENTO   -->
<br>
<table class="Q" cellspacing="0" style="width:649px;">
	<tr>
		<td class="MB" colspan="5" align="left"><p class="Rf">Observa��es I</p>
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
		<td class="MD" align="left" nowrap><p class="Rf">Observa��es II</p>
			<input name="c_obs2" id="c_obs2" class="PLLe" style="width:85px;margin-left:2pt;" 
				readonly tabindex=-1 value='<%=r_orcamento.obs_2%>'>
		</td>
		<td class="MD" nowrap align="left" valign="top"><p class="Rf">Entrega Imediata</p>
		<% 	if Cstr(r_orcamento.st_etg_imediata) = Cstr(COD_ETG_IMEDIATA_NAO) then
				s = "N�O"
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
				s = "N�O"
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
				s = "N�O"
			elseif Cstr(r_orcamento.InstaladorInstalaStatus) = Cstr(COD_INSTALADOR_INSTALA_SIM) then
				s = "SIM"
			else
				s = ""
				end if
		
			if s="" then s="&nbsp;"
		%>
		<p class="C" style="margin-top:3px;"><%=s%></p>
		</td>
		<td nowrap align="left" valign="top"><p class="Rf">Garantia Indicador</p>
		<% 	if Cstr(r_orcamento.GarantiaIndicadorStatus) = Cstr(COD_GARANTIA_INDICADOR_STATUS__NAO) then
				s = "N�O"
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
		<!--  � VISTA  -->
		<tr>
		  <td align="left">
			<table cellspacing="0" cellpadding="0" border="0">
			  <tr>
				<td align="left"><span class="C">� Vista&nbsp&nbsp;(<%=x_opcao_forma_pagamento(r_orcamento.av_forma_pagto)%>)</span></td>
			  </tr>
			</table>
		  </td>
		</tr>
		<% elseif Cstr(r_orcamento.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELA_UNICA then %>
		<!--  PARCELA �NICA  -->
		<tr>
		  <td align="left">
			<table cellspacing="0" cellpadding="0" border="0">
			  <tr>
				<td align="left"><span class="C">Parcela �nica:&nbsp;&nbsp;<%=SIMBOLO_MONETARIO%>&nbsp;<%=formata_moeda(r_orcamento.pu_valor)%>&nbsp;&nbsp;&nbsp;(<%=x_opcao_forma_pagamento(r_orcamento.pu_forma_pagto)%>)&nbsp;&nbsp;vencendo ap�s&nbsp;<%=formata_inteiro(r_orcamento.pu_vencto_apos)%>&nbsp;dias</span></td>
			  </tr>
			</table>
		  </td>
		</tr>
		<% elseif Cstr(r_orcamento.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_CARTAO then %>
		<!--  PARCELADO NO CART�O (INTERNET)  -->
		<tr>
		  <td align="left">
			<table cellspacing="0" cellpadding="0" border="0">
			  <tr>
				<td align="left"><span class="C">Parcelado no Cart�o (internet) em&nbsp;&nbsp;<%=Cstr(r_orcamento.pc_qtde_parcelas)%>&nbsp;x&nbsp;&nbsp;<%=SIMBOLO_MONETARIO & " " & formata_moeda(r_orcamento.pc_valor_parcela)%></span></td>
			  </tr>
			</table>
		  </td>
		</tr>
		<% elseif Cstr(r_orcamento.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_CARTAO_MAQUINETA then %>
		<!--  PARCELADO NO CART�O (MAQUINETA)  -->
		<tr>
		  <td align="left">
			<table cellspacing="0" cellpadding="0" border="0">
			  <tr>
				<td align="left"><span class="C">Parcelado no Cart�o (maquineta) em&nbsp;&nbsp;<%=Cstr(r_orcamento.pc_maquineta_qtde_parcelas)%>&nbsp;x&nbsp;&nbsp;<%=SIMBOLO_MONETARIO & " " & formata_moeda(r_orcamento.pc_maquineta_valor_parcela)%></span></td>
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
				<td align="left"><span class="C">Presta��es:&nbsp;&nbsp;<%=formata_inteiro(r_orcamento.pce_prestacao_qtde)%>&nbsp;x&nbsp;<%=SIMBOLO_MONETARIO%>&nbsp;<%=formata_moeda(r_orcamento.pce_prestacao_valor)%>&nbsp;&nbsp;&nbsp;(<%=x_opcao_forma_pagamento(r_orcamento.pce_forma_pagto_prestacao)%>)&nbsp;&nbsp;vencendo a cada&nbsp;<%=formata_inteiro(r_orcamento.pce_prestacao_periodo)%>&nbsp;dias</span></td>
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
				<td align="left"><span class="C">1� Presta��o:&nbsp;&nbsp;<%=SIMBOLO_MONETARIO%>&nbsp;<%=formata_moeda(r_orcamento.pse_prim_prest_valor)%>&nbsp;&nbsp;&nbsp;(<%=x_opcao_forma_pagamento(r_orcamento.pse_forma_pagto_prim_prest)%>)&nbsp;&nbsp;vencendo ap�s&nbsp;<%=formata_inteiro(r_orcamento.pse_prim_prest_apos)%>&nbsp;dias</span></td>
			  </tr>
			  <tr>
				<td align="left"><span class="C">Demais Presta��es:&nbsp;&nbsp;<%=Cstr(r_orcamento.pse_demais_prest_qtde)%>&nbsp;x&nbsp;<%=SIMBOLO_MONETARIO%>&nbsp;<%=formata_moeda(r_orcamento.pse_demais_prest_valor)%>&nbsp;&nbsp;&nbsp;(<%=x_opcao_forma_pagamento(r_orcamento.pse_forma_pagto_demais_prest)%>)&nbsp;&nbsp;vencendo a cada&nbsp;<%=Cstr(r_orcamento.pse_demais_prest_periodo)%>&nbsp;dias</span></td>
			  </tr>
			</table>
		  </td>
		</tr>
		<% end if %>
	  </table>
	</td>
  </tr>
  <tr>
	<td class="MC" align="left"><p class="Rf">Descri��o da Forma de Pagamento</p>
	  <textarea name="c_forma_pagto" id="c_forma_pagto" class="PLLe notPrint" rows="<%=Cstr(MAX_LINHAS_FORMA_PAGTO)%>"
				style="width:642px;margin-left:2pt;"
				READONLY tabindex=-1><%=r_orcamento.forma_pagto%></textarea>
		<span class="PLLe notVisible"><%
			s = substitui_caracteres(r_orcamento.forma_pagto,chr(13),"<br>")
			if s = "" then s = "&nbsp;"
			Response.Write s %></span>
	</td>
  </tr>
</table>
<% end if %>


<% if orcamento_selecionado <> orcamento_selecionado_inicial then %>
<!-- ************   SEPARADOR   ************ -->
<table class="notPrint" width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc" align="left">&nbsp;</td></tr>
</table>

<!-- ************   BOT�ES   ************ -->
<table class="notPrint" width="649" cellpadding="0" cellspacing="0">
	<tr>
		<td align="center">
			<%	if pagina_retorno <> "" then
					s_url = pagina_retorno
				else
					s_url="OrcamentoConsultaView.asp" & "?orcamento_selecionado=" & orcamento_selecionado_inicial & "&orcamento_selecionado_inicial=" & orcamento_selecionado_inicial & "&usuario=" & usuario & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))
					end if%>
			<a name="bVOLTAR" id="bVOLTAR" href="<%=s_url%>" title="volta para a p�gina anterior">
				<img src="../botao/voltar.gif" width="176" height="55" border="0">
		</td>
	</tr>
</table>
<% else %>
<br />
<br />
<br />
<% end if %>

</form>


<!-- ************   DIRECIONA PARA CADASTRO DE CLIENTES   ************ -->
<form method="post" action="ClienteConsultaView.asp" id="fCLI" name="fCLI">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name='cliente_selecionado' id="cliente_selecionado" value='<%=r_orcamento.id_cliente%>'>
<input type="hidden" name='pagina_retorno' id="pagina_retorno" value='OrcamentoConsultaView.asp?orcamento_selecionado=<%=orcamento_selecionado%>&orcamento_selecionado_inicial=<%=orcamento_selecionado_inicial%>&usuario=<%=usuario%>&<%=MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>'>
<input type="hidden" name="orcamento_selecionado" value="<%=orcamento_selecionado%>">
<input type="hidden" name="orcamento_selecionado_inicial" value="<%=orcamento_selecionado_inicial%>">
<input type="hidden" name="usuario" value="<%=usuario%>">
</form>


</center>
</body>

<% end if %>

</html>

<%
'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>