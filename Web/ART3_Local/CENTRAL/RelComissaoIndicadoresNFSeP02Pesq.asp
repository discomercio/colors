<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/Global.asp"    -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ========================================================
'	  RelComissaoIndicadoresNFSeP02Pesq.asp
'     ========================================================
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

	dim usuario
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

	dim s, s_filtro, qtde_indicadores_encontrados
	dim alerta
	alerta = ""
	qtde_indicadores_encontrados = 0

'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	FILTROS
	dim c_cnpj_nfse
	dim ckb_st_entrega_entregue, c_dt_entregue_inicio, c_dt_entregue_termino
	dim ckb_comissao_paga_sim, ckb_comissao_paga_nao
	dim ckb_st_pagto_pago, ckb_st_pagto_nao_pago, ckb_st_pagto_pago_parcial
	dim rb_visao
	
	c_cnpj_nfse = retorna_so_digitos(Request.Form("c_cnpj_nfse"))
	ckb_st_entrega_entregue = Trim(Request.Form("ckb_st_entrega_entregue"))
	c_dt_entregue_inicio = Trim(Request.Form("c_dt_entregue_inicio"))
	c_dt_entregue_termino = Trim(Request.Form("c_dt_entregue_termino"))

	ckb_comissao_paga_sim = Trim(Request.Form("ckb_comissao_paga_sim"))
	ckb_comissao_paga_nao = Trim(Request.Form("ckb_comissao_paga_nao"))

	ckb_st_pagto_pago = Trim(Request.Form("ckb_st_pagto_pago"))
	ckb_st_pagto_nao_pago = Trim(Request.Form("ckb_st_pagto_nao_pago"))
	ckb_st_pagto_pago_parcial = Trim(Request.Form("ckb_st_pagto_pago_parcial"))

	rb_visao = Trim(Request.Form("rb_visao"))
	
	if c_cnpj_nfse = "" then
		alerta = "CNPJ não foi informado"
	elseif Not cnpj_cpf_ok(c_cnpj_nfse) then
		alerta = "CNPJ informado é inválido"
		end if

	dim dtMinDtInicialFiltroPeriodo, intMaxDiasDtInicialFiltroPeriodo
	dim strMinDtInicialFiltroPeriodoYYYYMMDD, strMinDtInicialFiltroPeriodoDDMMYYYY
	if operacao_permitida(OP_CEN_RESTRINGE_DT_INICIAL_FILTRO_PERIODO, s_lista_operacoes_permitidas) then
		intMaxDiasDtInicialFiltroPeriodo = obtem_max_dias_dt_inicial_filtro_periodo()
		dtMinDtInicialFiltroPeriodo = Date - intMaxDiasDtInicialFiltroPeriodo
		strMinDtInicialFiltroPeriodoYYYYMMDD = formata_data_yyyymmdd(dtMinDtInicialFiltroPeriodo)
		strMinDtInicialFiltroPeriodoDDMMYYYY = formata_data(dtMinDtInicialFiltroPeriodo)
	else
		strMinDtInicialFiltroPeriodoYYYYMMDD = ""
		strMinDtInicialFiltroPeriodoDDMMYYYY = ""
		end if






' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________
' _____________________________________
' CONSULTA EXECUTA
'
sub consulta_executa
dim s, s_cor, s_sql
dim x, cab_table, cab
dim s_cnpj_nfse
dim n_reg
dim r

	cab_table = "<table cellspacing='0' cellpadding='0'>" & chr(13)
	cab = "	<tr style='background:azure'>" & chr(13) & _
		  "		<td class='MT TdCheckbox' valign='bottom' style='vertical-align:bottom;' nowrap>&nbsp;</td>" & chr(13) & _
		  "		<td class='MTBD TdCnpjNFSe' valign='bottom' style='vertical-align:bottom;' nowrap><span class='R TdSpn' style='font-weight:bold;'>CNPJ (NFS-e)</span></td>" & chr(13) & _
		  "		<td class='MTBD TdRazaoSocialNFSe' style='vertical-align:bottom;' valign='bottom'><span class='R TdSpn' style='font-weight:bold;'>Razão Social (NFS-e)</span></td>" & chr(13) & _
		  "		<td class='MTBD TdApelido' style='vertical-align:bottom;' valign='bottom'><span class='R TdSpn' style='font-weight:bold;'>Apelido</span></td>" & chr(13) & _
		  "		<td class='MTBD TdCnpjCadastro' style='vertical-align:bottom;' valign='bottom'><span class='R TdSpn' style='font-weight:bold;'>CNPJ/CPF (Cadastro)</span></td>" & chr(13) & _
		  "		<td class='MTBD TdRazaoSocialCadastro' style='vertical-align:bottom;' valign='bottom'><span class='R TdSpn' style='font-weight:bold;'>Razão Social (Cadastro)</span></td>" & chr(13) & _
		  "		<td class='MTBD TdStatus' style='vertical-align:bottom;' valign='bottom'><span class='R TdSpn' style='font-weight:bold;'>Status</span></td>" & chr(13) & _
		  "		<td class='MTBD TdVendedor' style='vertical-align:bottom;' valign='bottom'><span class='R TdSpn' style='font-weight:bold;'>Vendedor Associado</span></td>" & chr(13) & _
		  "	</tr>" & chr(13)

	x = cab_table & _
		cab

	n_reg = 0
	
	s_sql = "SELECT * FROM t_ORCAMENTISTA_E_INDICADOR WHERE (comissao_NFSe_cnpj = '" & retorna_so_digitos(c_cnpj_nfse) & "') ORDER BY status, apelido"
	set r = cn.Execute(s_sql)
	do while Not r.Eof
		n_reg = n_reg + 1

		x = x & "	<tr>" & chr(13)

		'Checkbox
		x = x & "		<td class='MDBE TdCheckbox'><input type='checkbox' class='CKB' name='ckb_id_indicador' value='" & Trim("" & r("Id")) & "'></td>" & chr(13)

		'CNPJ (NFS-e)
		s = Trim("" & r("comissao_NFSe_cnpj"))
		if s = "" then
			s = "&nbsp;"
		else
			s = cnpj_cpf_formata(s)
			end if
		x = x & "		<td class='MDB TdCnpjNFSe'><span class='Cn TdSpn'>" & s & "</span></td>" & chr(13)

		'Razão social (NFS-e)
		s = Trim("" & r("comissao_NFSe_razao_social"))
		if s = "" then s = "&nbsp;"
		x = x & "		<td class='MDB TdRazaoSocialNFSe'><span class='Cn TdSpn'>" & s & "</span></td>" & chr(13)

		'Apelido
		s = Trim("" & r("apelido"))
		if s = "" then s = "&nbsp;"
		x = x & "		<td class='MDB TdApelido'><span class='Cn TdSpn'>" & s & "</span></td>" & chr(13)

		'CNPJ (Cadastro)
		s = Trim("" & r("cnpj_cpf"))
		if s = "" then
			s = "&nbsp;"
		else
			s = cnpj_cpf_formata(s)
			end if
		x = x & "		<td class='MDB TdCnpjCadastro'><span class='Cn TdSpn'>" & s & "</span></td>" & chr(13)

		'Razão social (Cadastro)
		s = Trim("" & r("razao_social_nome_iniciais_em_maiusculas"))
		if s = "" then s = "&nbsp;"
		x = x & "		<td class='MDB TdRazaoSocialCadastro'><span class='Cn TdSpn'>" & s & "</span></td>" & chr(13)

		'Status
		s = Trim("" & r("status"))
		if s = "A" then
			s = "Ativo"
			s_cor = "green"
		else
			s = "Inativo"
			s_cor = "red"
			end if
		x = x & "		<td class='MDB TdStatus'><span class='Cn TdSpn' style='color:" & s_cor & ";'>" & s & "</span></td>" & chr(13)

		'Vendedor
		s = Trim("" & r("vendedor"))
		if s = "" then s = "&nbsp;"
		x = x & "		<td class='MDB TdVendedor'><span class='Cn TdSpn'>" & s & "</span></td>" & chr(13)

		x = x & "	</tr>" & chr(13)

		r.MoveNext
		loop

	if n_reg = 0 then
		x = x & "	<tr>" & chr(13) & _
				"		<td class='MDBE ALERTA' align='center' colspan='8'><span class='ALERTA'>&nbsp;NENHUM INDICADOR ENCONTRADO&nbsp;</span></TD>" & chr(13) & _
				"	</tr>" & chr(13)
		end if

	'Fecha tabela
	x = x & "</table>" & chr(13)
	Response.write x

	qtde_indicadores_encontrados = n_reg

	if r.State <> 0 then r.Close
	set r=nothing

end sub
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



<script src="<%=URL_FILE__JQUERY%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_I18N%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__AJAX_JS%>" language="JavaScript" type="text/javascript"></script>

<script type="text/javascript">
	$(function () {
		$(".aviso").css('display', 'none');

		$(".CKB").each(function () {
			if (this.checked) {
				$(this).closest("tr").css("background-color", "palegreen");
			}
			else {
				$(this).closest("tr").css("background-color", "");
			}
		});

		$(".CKB").change(function () {
			if (this.checked) {
				$(this).closest("tr").css("background-color", "palegreen");
			}
			else {
				$(this).closest("tr").css("background-color", "");
			}
		});
	});
</script>

<script language="JavaScript" type="text/javascript">
	function fFILTROConfirma(f) {
		if (!$("input:checkbox[name='ckb_id_indicador']").is(":checked")) {
			alert("Nenhum indicador foi selecionado!");
			return;
		}

		dCONFIRMA.style.visibility = "hidden";

		f.action = "RelComissaoIndicadoresNFSeP03Filtro.asp";

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
<link href="<%=URL_FILE__JQUERY_UI_CSS%>" rel="stylesheet" type="text/css">

<style type="text/css">

	.aviso {
		font-family: Arial, Helvetica, sans-serif;
		font-size: 8pt;
		font-weight: bold;
		font-style: normal;
		margin: 0pt 0pt 0pt 0pt;
		color: #f00;
		display: none;
	}

	.TdSpn {
		display: block;
		padding:2px;
		font-weight:bold;
	}
	.TdCheckbox {
		width: 20px;
		vertical-align:middle;
	}
	.TdCnpjNFSe {
		width: 110px;
		text-align:left;
		vertical-align:middle;
	}
	.TdRazaoSocialNFSe {
		width: 190px;
		text-align:left;
		vertical-align:middle;
		border-right:2px solid #C0C0C0;
	}
	.TdApelido {
		width: 160px;
		text-align:left;
		vertical-align:middle;
	}
	.TdCnpjCadastro {
		width: 110px;
		text-align:left;
		vertical-align:middle;
	}
	.TdRazaoSocialCadastro {
		width: 190px;
		text-align:left;
		vertical-align:middle;
	}
	.TdStatus {
		width: 50px;
		vertical-align:middle;
	}
	.TdVendedor {
		width: 90px;
		vertical-align:middle;
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
<body>
<center>

<form id="fFILTRO" name="fFILTRO" method="post" onsubmit="if (!fFILTROConfirma(fFILTRO)) return false;">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="c_MinDtInicialFiltroPeriodoYYYYMMDD" id="c_MinDtInicialFiltroPeriodoYYYYMMDD" value='<%=strMinDtInicialFiltroPeriodoYYYYMMDD%>'>
<input type="hidden" name="c_MinDtInicialFiltroPeriodoDDMMYYYY" id="c_MinDtInicialFiltroPeriodoDDMMYYYY" value='<%=strMinDtInicialFiltroPeriodoDDMMYYYY%>'>
<input type="hidden" name="c_cnpj_nfse" id="c_cnpj_nfse" value="<%=c_cnpj_nfse%>" />
<input type="hidden" name="ckb_st_entrega_entregue" id="ckb_st_entrega_entregue" value="<%=ckb_st_entrega_entregue%>" />
<input type="hidden" name="c_dt_entregue_inicio" id="c_dt_entregue_inicio" value="<%=c_dt_entregue_inicio%>" />
<input type="hidden" name="c_dt_entregue_termino" id="c_dt_entregue_termino" value="<%=c_dt_entregue_termino%>" />
<input type="hidden" name="ckb_comissao_paga_sim" id="ckb_comissao_paga_sim" value="<%=ckb_comissao_paga_sim%>" />
<input type="hidden" name="ckb_comissao_paga_nao" id="ckb_comissao_paga_nao" value="<%=ckb_comissao_paga_nao%>" />
<input type="hidden" name="ckb_st_pagto_pago" id="ckb_st_pagto_pago" value="<%=ckb_st_pagto_pago%>" />
<input type="hidden" name="ckb_st_pagto_nao_pago" id="ckb_st_pagto_nao_pago" value="<%=ckb_st_pagto_nao_pago%>" />
<input type="hidden" name="ckb_st_pagto_pago_parcial" id="ckb_st_pagto_pago_parcial" value="<%=ckb_st_pagto_pago_parcial%>" />
<input type="hidden" name="rb_visao" id="rb_visao" value="<%=rb_visao%>" />

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="940" cellpadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Relatório de Pedidos Indicadores (via NFS-e)</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>

<!-- FILTROS -->
<%
	s_filtro = "<table width='940' cellpadding='0' cellspacing='0' style='margin-top:8px;margin-bottom:8px;' border='0'>" & chr(13)

'	EMITENTE DA NFS-e
	s = ""
	if cnpj_cpf_ok(c_cnpj_nfse) then s = cnpj_cpf_formata(c_cnpj_nfse)
	if s = "" then s = "N.I."
	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' nowrap><span class='N'>CNPJ (emitente NFS-e):&nbsp;</span></td>" & chr(13) & _
				"		<td align='left' valign='top' width='99%'><span class='N'>" & s & "</span></td>" & chr(13) & _
				"	</tr>" & chr(13)

	s_filtro = s_filtro & "</table>" & chr(13)
	Response.Write s_filtro
%>

<table width="940" cellpadding="4" cellspacing="0" style="border-top:1px solid black">
<tr><td class="Rc" align="left">&nbsp;</td></tr>
</table>


<!--  LISTA DE INDICADORES PARA SELECIONAR  -->
<br />
<% consulta_executa %>


<!-- ************   SEPARADOR   ************ -->
<table width="940" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc" align="left">&nbsp;</td></tr>
</table>
<br />


<table width="940" cellspacing="0">
<% if qtde_indicadores_encontrados > 0 then %>
<tr>
	<td align="left"><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()" title="volta para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
	<td align="right"><div name="dCONFIRMA" id="dCONFIRMA"><a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fFILTROConfirma(fFILTRO)" title="prossegue para a próxima etapa">
		<img src="../botao/proximo.gif" width="176" height="55" border="0"></a></div>
	</td>
</tr>
<% else %>
<tr>
	<td align="center"><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()" title="volta para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
</tr>
<% end if %>
</table>
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