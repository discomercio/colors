<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =================================================================
'	  RelComissaoIndicadoresNFSeP06BotaoMagico.asp
'     =================================================================
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
	
	Server.ScriptTimeout = MAX_SERVER_SCRIPT_TIMEOUT_EM_SEG
	
	const ID_RELATORIO = "CENTRAL/RelComissaoIndicadoresNFSe"

'	COMO O TRATAMENTO DO RELATÓRIO PODE SER DEMORADO, CASO A SESSÃO EXPIRE E O TRATAMENTO
'	DE SESSÃO EXPIRADA NÃO CONSIGA RESTAURÁ-LA, OBTÉM A IDENTIFICAÇÃO DO USUÁRIO A PARTIR DE
'	UM CAMPO HIDDEN CRIADO NA PÁGINA CHAMADORA EXCLUSIVAMENTE P/ ISSO.
	dim s, s_aux, usuario, msg_erro
	usuario = Trim(Session("usuario_atual"))
	if (usuario = "") then usuario = Trim(Request("c_usuario_sessao"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO)

	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	if Not operacao_permitida(OP_CEN_FLAG_COMISSAO_PAGA, s_lista_operacoes_permitidas) then 
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if
	
'	FILTROS
	dim c_cnpj_nfse
	dim ckb_id_indicador
	dim rb_visao

	c_cnpj_nfse = retorna_so_digitos(Request.Form("c_cnpj_nfse"))
	ckb_id_indicador = Trim(Request.Form("ckb_id_indicador"))
	rb_visao = Trim(Request.Form("rb_visao"))

	dim id_nsu_N1
	id_nsu_N1 = Trim(Request.Form("id_nsu_N1"))

	dim proc_comissao_request_guid
	proc_comissao_request_guid = Trim(Request.Form("proc_comissao_request_guid"))

	dim origem
	origem = Trim(Request.Form("origem"))

	dim alerta
	alerta=""

	dim mensagem
	mensagem = ""

	dim blnErroFatal
	blnErroFatal = False

	dim s_rel_comissao_paga, s_rel_devolucao_descontada, s_rel_perda_descontada
	s_rel_comissao_paga = ""
	s_rel_devolucao_descontada = ""
	s_rel_perda_descontada = ""

'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, cn2, rs, tN1, tN3Ped
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	if Not bdd_conecta_RPIFC(cn2) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	If Not cria_recordset_otimista(tN1, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	If Not cria_recordset_otimista(tN3Ped, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

	s = "SELECT * FROM t_COMISSAO_INDICADOR_NFSe_N1 WHERE (id = " & id_nsu_N1 & ")"
	tN1.Open s, cn
	if tN1.Eof then
		blnErroFatal = True
		alerta=texto_add_br(alerta)
		alerta=alerta & "Falha ao tentar localizar dados do relatório (NSU = " & id_nsu_N1 & ")"
	else
		if c_cnpj_nfse = "" then c_cnpj_nfse = Trim("" & tN1("NFSe_cnpj"))
		if tN1("proc_comissao_status") = 0 then
			blnErroFatal = True
			alerta=texto_add_br(alerta)
			alerta=alerta & "O relatório ainda não processou o pagamento das comissões nos pedidos (NSU = " & id_nsu_N1 & ")"
			end if
		end if

	'OBTÉM O CÓDIGO DA EMPRESA P/ O LANÇAMENTO NO FLUXO DE CAIXA DE ACORDO C/ A LOJA DOS REGISTROS SELECIONADOS
	dim s_id_plano_contas_empresa, qtde_id_plano_contas_empresa, s_lista_lojas, qtde_lojas, alerta_plano_contas_empresa
	s_id_plano_contas_empresa = ""
	qtde_id_plano_contas_empresa = 0
	s_lista_lojas = ""
	qtde_lojas = 0
	alerta_plano_contas_empresa = ""
	if alerta = "" then
		s = "SELECT DISTINCT" & _
				" t_LOJA.id_plano_contas_empresa_comissao_indicador" & _
			" FROM t_COMISSAO_INDICADOR_NFSe_N1 tN1" & _
				" INNER JOIN t_COMISSAO_INDICADOR_NFSe_N2 tN2 ON (tN1.id = tN2.id_comissao_indicador_nfse_n1)" & _
				" INNER JOIN t_COMISSAO_INDICADOR_NFSe_N3_PEDIDO tN3Ped ON (tN2.id = tN3Ped.id_comissao_indicador_nfse_n2)" & _
				" INNER JOIN t_LOJA ON (tN3Ped.loja = t_LOJA.loja)" & _
			" WHERE" & _
				" (tN1.id = " & id_nsu_N1 & ")" & _
				" AND (tN3Ped.st_selecionado = 1)"
		rs.Open s, cn
		do while Not rs.Eof
			s_id_plano_contas_empresa = Trim("" & rs("id_plano_contas_empresa_comissao_indicador"))
			if s_id_plano_contas_empresa <> "" then qtde_id_plano_contas_empresa = qtde_id_plano_contas_empresa + 1
			rs.MoveNext
			loop

		if qtde_id_plano_contas_empresa > 1 then s_id_plano_contas_empresa = ""
		if rs.State <> 0 then rs.Close

		'TRATAMENTO P/ O CASO EM QUE NÃO ENCONTROU A EMPRESA P/ O LANÇAMENTO DO FLUXO DE CAIXA
		if s_id_plano_contas_empresa = "" then
			s = "SELECT DISTINCT" & _
					" tN3Ped.loja" & _
				" FROM t_COMISSAO_INDICADOR_NFSe_N1 tN1" & _
					" INNER JOIN t_COMISSAO_INDICADOR_NFSe_N2 tN2 ON (tN1.id = tN2.id_comissao_indicador_nfse_n1)" & _
					" INNER JOIN t_COMISSAO_INDICADOR_NFSe_N3_PEDIDO tN3Ped ON (tN2.id = tN3Ped.id_comissao_indicador_nfse_n2)" & _
				" WHERE" & _
					" (tN1.id = " & id_nsu_N1 & ")" & _
					" AND (tN3Ped.st_selecionado = 1)" & _
				" ORDER BY" & _
					" tN3Ped.loja"
			if rs.State <> 0 then rs.Close
			rs.Open s, cn
			do while Not rs.Eof
				if Trim("" & rs("loja")) <> "" then
					qtde_lojas = qtde_lojas + 1
					if s_lista_lojas <> "" then s_lista_lojas = s_lista_lojas & ", "
					s_lista_lojas = s_lista_lojas & Trim("" & rs("loja"))
					end if
				rs.MoveNext
				loop
			if rs.State <> 0 then rs.Close
			
			if qtde_lojas = 1 then
				alerta_plano_contas_empresa = "Não foi encontrada a empresa para o lançamento no fluxo de caixa referente à loja: " & s_lista_lojas
			else
				alerta_plano_contas_empresa = "Não foi encontrada a empresa para o lançamento no fluxo de caixa referente às lojas: " & s_lista_lojas
				end if
			end if 'if s_id_plano_contas_empresa = ""
		end if 'if alerta = ""

	dim fluxo_caixa_dt_competencia_default
	dim fluxo_caixa_conta_corrente_default, fluxo_caixa_plano_contas_RT_default
	
	if alerta = "" then
		fluxo_caixa_dt_competencia_default = get_default_valor_texto_bd(usuario, ID_RELATORIO & "|c_fluxo_caixa_dt_competencia")
	
		fluxo_caixa_conta_corrente_default = get_default_valor_texto_bd(usuario, ID_RELATORIO & "|c_fluxo_caixa_conta_corrente")
		if fluxo_caixa_conta_corrente_default = "" then fluxo_caixa_conta_corrente_default = getParametroFromCampoTexto(ID_PARAMETRO_RelComissaoIndicadoresNFSe_PlanoContas_ContaCorrente)
	
		fluxo_caixa_plano_contas_RT_default = get_default_valor_texto_bd(usuario, ID_RELATORIO & "|c_fluxo_caixa_plano_contas_RT")
		if fluxo_caixa_plano_contas_RT_default = "" then fluxo_caixa_plano_contas_RT_default = getParametroFromCampoTexto(ID_PARAMETRO_RelComissaoIndicadoresNFSe_PlanoContas_RT)
		end if 'if alerta = ""





' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________

function fluxo_caixa_conta_corrente_monta_itens_select(byval id_default)
dim x, r, s, strResp, ha_default

	id_default = Trim("" & id_default)
	ha_default = False

	s = "SELECT * FROM t_FIN_CONTA_CORRENTE WHERE (st_ativo = 1) ORDER BY id"
	set r = cn2.Execute(s)

	strResp = ""
	do while Not r.Eof
		x = Trim("" & r("id"))

		if (id_default=x) then
			strResp = strResp & "<option selected"
			ha_default = True
		else
			strResp = strResp & "<option"
			end if

		strResp = strResp & " value='" & x & "'>"
		strResp = strResp & Trim("" & r("banco")) & " &nbsp; " & Trim("" & r("agencia")) & " &nbsp; " & Trim("" & r("conta")) & " &nbsp; " & Trim("" & r("descricao"))
		strResp = strResp & "</option>" & chr(13)
		r.MoveNext
		loop

	if Not ha_default then
		strResp = "<option selected value=''>&nbsp;</option>" & chr(13) & strResp
		end if

	fluxo_caixa_conta_corrente_monta_itens_select = strResp
	r.close
	set r = Nothing
end function


function fluxo_caixa_conta_corrente_monta_descricao(byval id_conta_corrente)
dim r, s, strResp

	fluxo_caixa_conta_corrente_monta_descricao = ""

	id_conta_corrente = Trim("" & id_conta_corrente)
	if id_conta_corrente = "" then exit function

	s = "SELECT * FROM t_FIN_CONTA_CORRENTE WHERE (id = " & id_conta_corrente & ")"
	set r = cn2.Execute(s)

	strResp = ""
	if Not r.Eof then
		strResp = Trim("" & r("banco")) & " &nbsp; " & Trim("" & r("agencia")) & " &nbsp; " & Trim("" & r("conta")) & " &nbsp; " & Trim("" & r("descricao"))
		end if

	fluxo_caixa_conta_corrente_monta_descricao = strResp
	r.close
	set r = Nothing
end function


function fluxo_caixa_empresa_monta_itens_select(byval id_default)
dim x, r, s, strResp, ha_default

	id_default = Trim("" & id_default)
	ha_default = False

	s = "SELECT * FROM t_FIN_PLANO_CONTAS_EMPRESA WHERE (st_ativo = 1) ORDER BY id"
	set r = cn2.Execute(s)

	strResp = ""
	do while Not r.Eof
		x = Trim("" & r("id"))

		if (id_default=x) then
			strResp = strResp & "<option selected"
			ha_default = True
		else
			strResp = strResp & "<option"
			end if

		strResp = strResp & " value='" & x & "'>"
		strResp = strResp & Trim("" & r("id")) & " - " & Trim("" & r("descricao"))
		strResp = strResp & "</option>" & chr(13)
		r.MoveNext
		loop

	if Not ha_default then
		strResp = "<option selected value=''>&nbsp;</option>" & chr(13) & strResp
		end if

	fluxo_caixa_empresa_monta_itens_select = strResp
	r.close
	set r = Nothing
end function


function fluxo_caixa_empresa_monta_descricao(byval id_empresa)
dim r, s, strResp

	fluxo_caixa_empresa_monta_descricao = ""

	id_empresa = Trim("" & id_empresa)
	if id_empresa = "" then exit function

	s = "SELECT * FROM t_FIN_PLANO_CONTAS_EMPRESA WHERE (id = " & id_empresa & ")"
	set r = cn2.Execute(s)

	strResp = ""
	if Not r.Eof then
		strResp = Trim("" & r("id")) & " - " & Trim("" & r("descricao"))
		end if

	fluxo_caixa_empresa_monta_descricao = strResp
	r.close
	set r = Nothing
end function

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
	<title>CENTRAL</title>
	</head>


<% if False then 'APENAS P/ HABILITAR O INTELLISENSE DURANTE O DESENVOLVIMENTO!! %>
<script src="../Global/jquery.js" language="JavaScript" type="text/javascript"></script>
<% end if %>

<script src="<%=URL_FILE__JQUERY%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_I18N%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
	$(function () {
		$("#c_fluxo_caixa_dt_competencia").hUtilUI('datepicker_padrao_peq');
	});

	function fAvisoVoltar(f) {
		f.action = "RelComissaoIndicadoresNFSeP01Filtro.asp?url_back=X";
		f.submit();
	}

	function fRetornar(f) {
		f.action = "RelComissaoIndicadoresNFSeP01Filtro.asp?url_back=X";
		dVOLTAR.style.visibility = "hidden";
		f.submit();
	}

	function fProcessar(f) {
		if (trim($("#c_numero_nfse").val()) == "") {
			alert("Informe o número da NFS-e!");
			$("#c_numero_nfse").select();
			$("#c_numero_nfse").focus();
			return;
		}

		if (retorna_so_digitos($("#c_numero_nfse").val()) != $("#c_numero_nfse").val()) {
			alert("Número da NFS-e contém caracteres inválidos!");
			$("#c_numero_nfse").select();
			$("#c_numero_nfse").focus();
			return;
		}

		if (trim($("#c_fluxo_caixa_dt_competencia").val()) == "") {
			alert("Informe a data de competência para o lançamento no fluxo de caixa!");
			$("#c_fluxo_caixa_dt_competencia").select();
			$("#c_fluxo_caixa_dt_competencia").focus();
			return;
		}

		if (!isDate(f.c_fluxo_caixa_dt_competencia)) {
			alert("Data inválida!");
			$("#c_fluxo_caixa_dt_competencia").select();
			$("#c_fluxo_caixa_dt_competencia").focus();
			return;
		}

		if (trim($('select[name="c_fluxo_caixa_conta_corrente"] option').filter(':selected').val()) == "") {
			alert("Selecione a conta corrente para o lançamento no fluxo de caixa!");
			return;
		}

		if (trim($('select[name="c_fluxo_caixa_empresa"] option').filter(':selected').val()) == "") {
			alert("Selecione a empresa para o lançamento no fluxo de caixa!");
			return;
		}

		// Emily em 01/07/2022: os valores de comissão (RT) e RA Líquido estão sendo somados e o valor total está sendo lançado na conta 1400
		if (trim($("#c_fluxo_caixa_plano_contas_RT").val()) == "") {
			alert("Informe o número do plano de contas da comissão (RT) para o lançamento no fluxo de caixa!");
			$("#c_fluxo_caixa_plano_contas_RT").select();
			$("#c_fluxo_caixa_plano_contas_RT").focus();
			return;
		}

		if (retorna_so_digitos($("#c_fluxo_caixa_plano_contas_RT").val()) != $("#c_fluxo_caixa_plano_contas_RT").val()) {
			alert("Número do plano de contas da comissão (RT) contém caracteres inválidos!");
			$("#c_fluxo_caixa_plano_contas_RT").select();
			$("#c_fluxo_caixa_plano_contas_RT").focus();
			return;
		}

		if (!confirm("Grava o(s) lançamento(s) no fluxo de caixa?")) return;

		dPROCESSAR.style.visibility = "hidden";
		f.action = "RelComissaoIndicadoresNFSeP07ProcessarBotaoMagico.asp";
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
<link href="<%=URL_FILE__JQUERY_UI_CSS%>" rel="stylesheet" type="text/css">

<style type="text/css">
.TdLabel{
	width:200px;
}

.TdInfo{
	width:600px;
}
</style>

<% if alerta <> "" then %>
<!-- ************************************************************ -->
<!-- **********  PÁGINA PARA EXIBIR MENSAGENS DE ERRO  ********** -->
<!-- ************************************************************ -->
<body onload="bVOLTAR.focus();">
<center>
<br />
<!--  T E L A  -->
<form id="fAviso" name="fAviso" method="post">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<!-- FILTROS -->
<input type="hidden" name="c_cnpj_nfse" id="c_cnpj_nfse" value="<%=c_cnpj_nfse%>" />
<input type="hidden" name="ckb_id_indicador" id="ckb_id_indicador" value="<%=ckb_id_indicador%>" />
<input type="hidden" name="rb_visao" id="rb_visao" value="<%=rb_visao%>">
<input type="hidden" name="id_nsu_N1" id="id_nsu_N1" value="<%=id_nsu_N1%>" />
<input type="hidden" name="proc_comissao_request_guid" id="proc_comissao_request_guid" value="<%=proc_comissao_request_guid%>" />


<p class="T">A V I S O</p>
<div class="MtAlerta" style="width:600px;font-weight:bold;" align="center"><p style='margin:5px 2px 5px 2px;'><%=alerta%></p></div>
<br /><br />
<p class="TracoBottom"></p>
<table cellspacing="0">
<tr>
	<td align="center">
        <% if blnErroFatal then %>
        <a name="bVOLTAR" id="bVOLTAR" href="javascript:fAvisoVoltar(fAviso)">
        <% else %>
        <a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()">
        <% end if %>
        <img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
</tr>
</table>
</form>
</center>
</body>

<% else %>

<!-- ***************************************************** -->
<!-- **********  PÁGINA PARA EXIBIR RESULTADO   ********** -->
<!-- ***************************************************** -->
<body onload="window.status='Concluído';f.c_numero_nfse.focus();" link=#000000 alink=#000000 vlink=#000000>

<center>

<form id="f" name="f" method="post">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<!-- FILTROS -->
<input type="hidden" name="c_cnpj_nfse" id="c_cnpj_nfse" value="<%=c_cnpj_nfse%>" />
<input type="hidden" name="ckb_id_indicador" id="ckb_id_indicador" value="<%=ckb_id_indicador%>" />
<input type="hidden" name="rb_visao" id="rb_visao" value="<%=rb_visao%>">
<input type="hidden" name="id_nsu_N1" id="id_nsu_N1" value="<%=id_nsu_N1%>" />
<input type="hidden" name="proc_comissao_request_guid" id="proc_comissao_request_guid" value="<%=proc_comissao_request_guid%>" />
<input type="hidden" name="proc_fluxo_caixa_request_guid" id="proc_fluxo_caixa_request_guid" value="<%=gera_uid%>" />


<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="820" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black;">
	<tr>
		<td align="right" valign="bottom"><span class="PEDIDO">Relatório de Pedidos Indicadores (via NFS-e)</span></td>
	</tr>
</table>
<br />
<br />

<%
%>

<% if blnErroFatal then %>
<div class='MtAlerta' style='width:800px;font-weight:bold;' align='center'>
<p style='margin:5px 2px 5px 2px;'><%=mensagem%></p></div>
<br />
<% end if %>


<% if Not blnErroFatal then %>

<% if tN1("proc_fluxo_caixa_status") <> 0 then %>
<!-- ************   MENSAGEM DE ALERTA INFORMANDO QUE O RELATÓRIO JÁ FOI PROCESSADO ANTERIORMENTE ************ -->
<div class='MtAlerta' style='width:800px;font-weight:bold;' align='center'>
<p style='margin:5px 2px 5px 2px;'>Este relatório já processou os lançamentos no fluxo de caixa em <%=formata_data_hora_sem_seg(tN1("proc_fluxo_caixa_data_hora"))%> por <%=Trim("" & tN1("proc_fluxo_caixa_usuario"))%></p></div>
<br />
<br />
<% end if 'if tN1("proc_fluxo_caixa_status") <> 0 %>


<!-- ************   MENSAGEM  ************ -->
<table class="Qx" style="width:800px;" cellpadding="1" cellspacing="0">
	<tr>
		<td class="MT TdLabel" align="right"><span class="Cd">NSU do Relatório:</span></td>
		<td class="MTBD TdInfo" align="left"><span class="C"><%=CStr(id_nsu_N1)%></span></td>
	</tr>
	<tr>
		<td class="MDBE TdLabel" align="right"><span class="Cd">CNPJ NFS-e:</span></td>
		<td class="MDB TdInfo" align="left"><span class="C"><%=cnpj_cpf_formata(c_cnpj_nfse)%></span></td>
	</tr>
	<tr>
		<td class="MDBE TdLabel" align="right"><span class="Cd">Comissão Processada em:</span></td>
		<td class="MDB TdInfo" align="left"><span class="C"><%=formata_data_e_talvez_hora_hhmm(tN1("proc_comissao_data_hora"))%> por <%=Trim("" & tN1("proc_comissao_usuario"))%></span></td>
	</tr>
	<tr>
		<td class="MDBE TdLabel" align="right"><span class="Cd">Lançamentos Processados em:</span></td>
		<% if tN1("proc_fluxo_caixa_status") = 1 then %>
		<td class="MDB TdInfo" align="left"><span class="C"><%=formata_data_e_talvez_hora_hhmm(tN1("proc_fluxo_caixa_data_hora"))%> por <%=Trim("" & tN1("proc_fluxo_caixa_usuario"))%></span></td>
		<% else %>
		<td class="MDB TdInfo" align="left"><span class="C" style="color:red;">Não processado</span></td>
		<% end if %>
	</tr>
	<tr>
		<%
			s = "SELECT" & _
						" tN3Ped.*" & _
					" FROM t_COMISSAO_INDICADOR_NFSe_N3_PEDIDO tN3Ped" & _
						" INNER JOIN t_COMISSAO_INDICADOR_NFSe_N2 tN2 ON (tN3Ped.id_comissao_indicador_nfse_n2 = tN2.id)" & _
						" INNER JOIN t_COMISSAO_INDICADOR_NFSe_N1 tN1 ON (tN2.id_comissao_indicador_nfse_n1 = tN1.id)" & _
					" WHERE" & _
						" (tN1.id = " & id_nsu_N1 & ")" & _
						" AND (tN3Ped.st_selecionado = 1)" & _
						" AND (tN3Ped.id_cfg_tabela_origem = " & ID_CFG_TABELA_ORIGEM_T_PEDIDO & ")" & _
					" ORDER BY" & _
						" tN3Ped.id"
			if tN3Ped.State <> 0 then tN3Ped.Close
			tN3Ped.Open s, cn
			do while Not tN3Ped.Eof
				if s_rel_comissao_paga <> "" then s_rel_comissao_paga = s_rel_comissao_paga & ", "
				s_rel_comissao_paga = s_rel_comissao_paga & Trim("" & tN3Ped("pedido"))
				tN3Ped.MoveNext
				loop

		s_aux = s_rel_comissao_paga
		if s_aux = "" then s_aux = "(nenhum pedido)"
		%>
		<td class="MDBE TdLabel" valign="middle" align="right"><span class="Cd">Comissão Paga:</span></td>
		<td class="MDB TdInfo" align="left"><span class="C"><%=s_aux%></span></td>
	</tr>
	<tr>
		<%
			s = "SELECT" & _
						" tN3Ped.*" & _
					" FROM t_COMISSAO_INDICADOR_NFSe_N3_PEDIDO tN3Ped" & _
						" INNER JOIN t_COMISSAO_INDICADOR_NFSe_N2 tN2 ON (tN3Ped.id_comissao_indicador_nfse_n2 = tN2.id)" & _
						" INNER JOIN t_COMISSAO_INDICADOR_NFSe_N1 tN1 ON (tN2.id_comissao_indicador_nfse_n1 = tN1.id)" & _
					" WHERE" & _
						" (tN1.id = " & id_nsu_N1 & ")" & _
						" AND (tN3Ped.st_selecionado = 1)" & _
						" AND (tN3Ped.id_cfg_tabela_origem = " & ID_CFG_TABELA_ORIGEM_T_PEDIDO_ITEM_DEVOLVIDO & ")" & _
					" ORDER BY" & _
						" tN3Ped.id"
			if tN3Ped.State <> 0 then tN3Ped.Close
			tN3Ped.Open s, cn
			do while Not tN3Ped.Eof
				if s_rel_devolucao_descontada <> "" then s_rel_devolucao_descontada = s_rel_devolucao_descontada & ", "
				s_rel_devolucao_descontada = s_rel_devolucao_descontada & Trim("" & tN3Ped("pedido"))
				tN3Ped.MoveNext
				loop

		%>
		<% s_aux = s_rel_devolucao_descontada
			if s_aux = "" then s_aux = "(nenhum pedido)"
		%>
		<td class="MDBE TdLabel" valign="middle" align="right"><span class="Cd">Devolução Descontada:</span></td>
		<td class="MDB TdInfo" align="left"><span class="C"><%=s_aux%></span></td>
	</tr>
	<tr>
		<%
			s = "SELECT" & _
						" tN3Ped.*" & _
					" FROM t_COMISSAO_INDICADOR_NFSe_N3_PEDIDO tN3Ped" & _
						" INNER JOIN t_COMISSAO_INDICADOR_NFSe_N2 tN2 ON (tN3Ped.id_comissao_indicador_nfse_n2 = tN2.id)" & _
						" INNER JOIN t_COMISSAO_INDICADOR_NFSe_N1 tN1 ON (tN2.id_comissao_indicador_nfse_n1 = tN1.id)" & _
					" WHERE" & _
						" (tN1.id = " & id_nsu_N1 & ")" & _
						" AND (tN3Ped.st_selecionado = 1)" & _
						" AND (tN3Ped.id_cfg_tabela_origem = " & ID_CFG_TABELA_ORIGEM_T_PEDIDO_PERDA & ")" & _
					" ORDER BY" & _
						" tN3Ped.id"
			if tN3Ped.State <> 0 then tN3Ped.Close
			tN3Ped.Open s, cn
			do while Not tN3Ped.Eof
				if s_rel_perda_descontada <> "" then s_rel_perda_descontada = s_rel_perda_descontada & ", "
				s_rel_perda_descontada = s_rel_perda_descontada & Trim("" & tN3Ped("pedido"))
				tN3Ped.MoveNext
				loop
		%>
		<% s_aux = s_rel_perda_descontada
			if s_aux = "" then s_aux = "(nenhum pedido)"
		%>
		<td class="MDBE TdLabel" valign="middle" align="right"><span class="Cd">Perda Descontada:</span></td>
		<td class="MDB TdInfo" align="left"><span class="C"><%=s_aux%></span></td>
	</tr>
	<tr>
		<% s_aux = formata_moeda(tN1("vl_total_geral_selecionado_RT"))
			if s_aux = "" then s_aux = "?"
		%>
		<td class="MDBE TdLabel" align="right"><span class="Cd">Comissão (RT):</span></td>
		<td class="MDB TdInfo" align="left"><span class="C"><%=s_aux%></span></td>
	</tr>
	<tr>
		<% s_aux = formata_moeda(tN1("vl_total_geral_selecionado_RA_liquido"))
			if s_aux = "" then s_aux = "?"
		%>
		<td class="MDBE TdLabel" align="right"><span class="Cd">RA Líquido:</span></td>
		<td class="MDB TdInfo" align="left"><span class="C"><%=s_aux%></span></td>
	</tr>
	<tr>
		<% s_aux = formata_moeda(tN1("vl_total_geral_selecionado_RT") + tN1("vl_total_geral_selecionado_RA_liquido"))
			if s_aux = "" then s_aux = "?"
		%>
		<td class="MDBE TdLabel" align="right"><span class="Cd">Total Comissão (RT + RA Líquido):</span></td>
		<td class="MDB TdInfo" align="left"><span class="C"><%=s_aux%></span></td>
	</tr>
</table>
<br />
<br />
<% if tN1("proc_fluxo_caixa_status") = 0 then %>
<table class="Qx" style="width:800px;" cellpadding="1" cellspacing="0">
	<tr>
		<td class="MT TdLabel" valign="middle" align="right"><span class="Cd">Número NFS-e:</span></td>
		<td class="MTBD TdInfo" align="left"><input type="text" name="c_numero_nfse" id="c_numero_nfse" maxlength="9" class="PLLe" style="font-size:11pt;width:300px;" onblur="this.value=trim(this.value);" onkeypress="if (digitou_enter(true)) f.c_fluxo_caixa_dt_competencia.focus(); filtra_numerico();" /></td>
	</tr>
	<tr>
		<td class="MDBE TdLabel" valign="middle" align="right"><span class="Cd">Data Competência:</span></td>
		<td class="MDB TdInfo" align="left"><input type="text" name="c_fluxo_caixa_dt_competencia" id="c_fluxo_caixa_dt_competencia" maxlength="10" class="PLLe" style="font-size:11pt;width:100px;" onblur="if (!isDate(this)) {alert('Data inválida!'); this.focus();}" onkeypress="if (digitou_enter(true)) f.c_fluxo_caixa_plano_contas_RT.focus(); filtra_data();" value="<%=fluxo_caixa_dt_competencia_default%>" /></td>
	</tr>
	<tr>
		<td class="MDBE TdLabel" valign="middle" align="right"><span class="Cd">Conta Corrente:</span></td>
		<td class="MDB TdInfo" align="left">
				<select id="c_fluxo_caixa_conta_corrente" name="c_fluxo_caixa_conta_corrente" style="width:500px;margin:6px;">
					<%=fluxo_caixa_conta_corrente_monta_itens_select(fluxo_caixa_conta_corrente_default)%>
				</select>
		</td>
	</tr>
	<tr>
		<td class="MDBE TdLabel" valign="middle" align="right"><span class="Cd">Empresa:</span></td>
		<td class="MDB TdInfo" align="left">
				<select id="c_fluxo_caixa_empresa" name="c_fluxo_caixa_empresa" style="width:500px;margin:6px;">
					<%=fluxo_caixa_empresa_monta_itens_select(s_id_plano_contas_empresa)%>
				</select>
		</td>
	</tr>
	<tr>
		<td class="MDBE TdLabel" valign="middle" align="right"><span class="Cd">Plano Contas:</span></td>
		<td class="MDB TdInfo" align="left"><input type="text" name="c_fluxo_caixa_plano_contas_RT" id="c_fluxo_caixa_plano_contas_RT" maxlength="6" class="PLLe" style="font-size:11pt;width:300px;" onblur="this.value=trim(this.value);" onkeypress="filtra_numerico();" value="<%=fluxo_caixa_plano_contas_RT_default%>" /></td>
	</tr>
<% if alerta_plano_contas_empresa <> "" then %>
	<tr>
		<td colspan="2" style="height:4px;">&nbsp;</td>
	</tr>
	<tr>
		<td colspan="2" class="MT" align="center">
			<span style="color:red;font-weight:bold;"><%=alerta_plano_contas_empresa%></span>
		</td>
	</tr>
<% end if %>
</table>
<br />
<% else 'if tN1("proc_fluxo_caixa_status") = 0 %>
<table class="Qx" style="width:800px;" cellpadding="1" cellspacing="0">
	<tr>
		<td class="MT TdLabel" valign="middle" align="right"><span class="Cd">Número NFS-e:</span></td>
		<td class="MTBD TdInfo" align="left"><span class="C"><%=Trim("" & tN1("NFSe_numero"))%></span></td>
	</tr>
	<tr>
		<td class="MDBE TdLabel" valign="middle" align="right"><span class="Cd">Data Competência:</span></td>
		<td class="MDB TdInfo" align="left"><span class="C"><%=formata_data(tN1("fluxo_caixa_dt_competencia"))%></span></td>
	</tr>
	<tr>
		<td class="MDBE TdLabel" valign="middle" align="right"><span class="Cd">Conta Corrente:</span></td>
		<td class="MDB TdInfo" align="left"><span class="C"><%=fluxo_caixa_conta_corrente_monta_descricao(tN1("fluxo_caixa_id_conta_corrente"))%></span></td>
	</tr>
	<tr>
		<td class="MDBE TdLabel" valign="middle" align="right"><span class="Cd">Empresa:</span></td>
		<td class="MDB TdInfo" align="left"><span class="C"><%=fluxo_caixa_empresa_monta_descricao(tN1("fluxo_caixa_id_plano_contas_empresa"))%></span></td>
	</tr>
	<tr>
		<td class="MDBE TdLabel" valign="middle" align="right"><span class="Cd">Plano Contas:</span></td>
		<td class="MDB TdInfo" align="left"><span class="C"><%=Trim("" & tN1("fluxo_caixa_comissao_id_plano_contas_conta"))%></span></td>
	</tr>
</table>
<br />
<% end if 'if tN1("proc_fluxo_caixa_status") = 0 %>

<% end if 'if Not blnErroFatal %>


<!-- ************   SEPARADOR   ************ -->
<table class="notPrint" width="820" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black;">
<tr><td class="Rc" align="left">&nbsp;</td></tr>
</table>

<!-- ************   LINKS: PÁGINA INICIAL / ENCERRA SESSÃO   ************ -->
<table class="notPrint" width="820" cellpadding="0" cellspacing="0">
<tr><td align="right"><span class="Rc">
	<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
	<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
	</span>
</td></tr>
</table>

<!-- ************   BOTÕES   ************ -->
<%
dim acao_botao_voltar, s_title
if origem = "QUERY" then
	acao_botao_voltar = "javascript:history.back()"
	s_title = "Retornar"
else
	acao_botao_voltar = "javascript:fRetornar(f)"
	s_title = "Retornar para o início do relatório"
	end if
%>
<table class="notPrint" width="820" cellspacing="0">
<% if (tN1("proc_fluxo_caixa_status") = 0) And (Not blnErroFatal) then %>
<tr>
	<td align="left"><div name="dVOLTAR" id="dVOLTAR"><a name="bVOLTAR" id="bVOLTAR" href="<%=acao_botao_voltar%>" title="<%=s_title%>">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a></div>
	</td>
	<td align="right"><div name="dPROCESSAR" id="dPROCESSAR"><a name="bPROCESSAR" id="bPROCESSAR" href="javascript:fProcessar(f)" title="Gravar lançamentos no fluxo de caixa">
		<img src="../botao/processar.gif" width="176" height="55" border="0"></a></div>
	</td>
</tr>
<% else %>
<tr>
	<td align="center"><div name="dVOLTAR" id="dVOLTAR"><a name="bVOLTAR" id="bVOLTAR" href="<%=acao_botao_voltar%>" title="<%=s_title%>">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a></div>
	</td>
</tr>
<% end if %>
</table>
</form>

</center>
</body>

<% end if %>

</html>


<%
	if tN1.State <> 0 then tN1.Close
	set tN1 = nothing

	if tN3Ped.State <> 0 then tN3Ped.Close
	set tN3Ped = nothing

'	FECHA CONEXAO COM O BANCO DE DADOS
	cn2.Close
	set cn2 = nothing

	cn.Close
	set cn = nothing
%>