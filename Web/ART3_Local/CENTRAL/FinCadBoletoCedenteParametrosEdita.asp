<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/Global.asp"    -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ==========================================
'	  FinCadBoletoCedenteParametrosEdita.asp
'     ==========================================
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
	dim s, strSql, usuario, operacao_selecionada
	dim id_conta_corrente_selecionado, id_boleto_cedente_selecionado

	usuario = trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 
	
'	REGISTRO A EDITAR
	operacao_selecionada = trim(request("operacao_selecionada"))
	
	id_conta_corrente_selecionado = ""
	id_boleto_cedente_selecionado = ""
	
	if operacao_selecionada=OP_INCLUI then
		id_conta_corrente_selecionado = retorna_so_digitos(request("id_selecionado"))
	else
		id_boleto_cedente_selecionado = retorna_so_digitos(request("id_selecionado"))
		end if

	if (operacao_selecionada<>OP_INCLUI) And (operacao_selecionada<>OP_CONSULTA) then Response.Redirect("aviso.asp?id=" & ERR_OPERACAO_NAO_ESPECIFICADA)

	if operacao_selecionada=OP_INCLUI then
		if (id_conta_corrente_selecionado="") Or (converte_numero(id_conta_corrente_selecionado)=0) then Response.Redirect("aviso.asp?id=" & ERR_ID_NAO_INFORMADO) 
	else
		if (id_boleto_cedente_selecionado="") Or (converte_numero(id_boleto_cedente_selecionado)=0) then Response.Redirect("aviso.asp?id=" & ERR_ID_NAO_INFORMADO) 
		end if

'	CONECTA COM O BANCO DE DADOS
	dim cn,rs
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	if operacao_selecionada=OP_INCLUI then
		strSql = "SELECT " & _
					"*" & _
				" FROM t_FIN_CONTA_CORRENTE" & _
				" WHERE" & _
					" (id = " & id_conta_corrente_selecionado & ")"
		set rs = cn.Execute(strSql)
		if Err <> 0 then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
		if rs.EOF then Response.Redirect("aviso.asp?id=" & ERR_ID_NAO_CADASTRADO)
	elseif operacao_selecionada=OP_CONSULTA then
		strSql = "SELECT " & _
					"*" & _
				" FROM t_FIN_BOLETO_CEDENTE" & _
				" WHERE" & _
					" (id = " & id_boleto_cedente_selecionado & ")"
		set rs = cn.Execute(strSql)
		if Err <> 0 then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
		if rs.EOF then Response.Redirect("aviso.asp?id=" & ERR_ID_NAO_CADASTRADO)
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

<script src="<%=URL_FILE__JQUERY%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_I18N%>" Language="JavaScript" type="text/javascript"></script>
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

function consiste_endereco_cadastro( f ) {

	if (trim(f.endereco.value)=="") {
		alert('Preencha o endereço!!');
		f.endereco.focus();
		return false;
		}

	if (trim(f.endereco_numero.value)=="") {
		alert('Preencha o número do endereço!!');
		f.endereco_numero.focus();
		return false;
		}
		
	if (trim(f.bairro.value)=="") {
		alert('Preencha o bairro!!');
		f.bairro.focus();
		return false;
		}

	if (trim(f.cidade.value)=="") {
		alert('Preencha a cidade!!');
		f.cidade.focus();
		return false;
		}

	s=trim(f.uf.value);
	if ((s=="")||(!uf_ok(s))) {
		alert('UF inválida!!');
		f.uf.focus();
		return false;
		}

	if (trim(f.cep.value)=="") {
		alert('Informe o CEP!!');
		return false;
		}
		
	if (!cep_ok(f.cep.value)) {
		alert('CEP inválido!!');
		f.cep.focus();
		return false;
		}
		
	return true;
}

function RemoveRegistro( f ) {
var b;
	b=window.confirm('Confirma a exclusão?');
	if (b){
		f.operacao_selecionada.value=OP_EXCLUI;
		dREMOVE.style.visibility="hidden";
		window.status = "Aguarde ...";
		f.submit();
		}
}

function AtualizaRegistro(f) {
	if (trim(f.c_apelido.value) == "") {
		alert('Informe um nome curto (apelido) para o cedente!!');
		f.c_apelido.focus();
		return;
		}
	if (trim(f.c_loja_default_boleto_plano_contas.value) == "") {
		alert('Informe o nº da loja padrão para se obter o plano de contas para o qual os lançamentos serão vinculados na situação em que não for possível determinar o nº da loja associada ao boleto.');
		f.c_loja_default_boleto_plano_contas.focus();
		return;
		}
	if (trim(f.c_num_banco.value)=="") {
		alert('Informe o número do banco!!');
		f.c_num_banco.focus();
		return;
		}
	if (trim(f.c_nome_banco.value)=="") {
		alert('Informe o nome do banco!!');
		f.c_nome_banco.focus();
		return;
		}
	if (trim(f.c_agencia.value)=="") {
		alert('Informe o número da agência!!');
		f.c_agencia.focus();
		return;
		}
	if (trim(f.c_digito_agencia.value)=="") {
		alert('Informe o dígito da agência!!');
		f.c_digito_agencia.focus();
		return;
		}
	if (trim(f.c_conta.value)=="") {
		alert('Informe o número da conta corrente!!');
		f.c_conta.focus();
		return;
		}
	if (trim(f.c_digito_conta.value)=="") {
		alert('Informe o dígito da conta corrente!!');
		f.c_digito_conta.focus();
		return;
		}
	if (trim(f.c_carteira.value)=="") {
		alert('Informe o número da carteira!!');
		f.c_carteira.focus();
		return;
		}
	if (trim(f.c_codigo_empresa.value)=="") {
		alert('Informe o código da empresa!!');
		f.c_codigo_empresa.focus();
		return;
		}
	if (trim(f.c_nome_empresa.value)=="") {
		alert('Informe o nome da empresa!!');
		f.c_nome_empresa.focus();
		return;
		}
	if (trim(f.c_juros_mora.value)=="") {
		alert('Informe o percentual do juros de mora!!');
		f.c_juros_mora.focus();
		return;
		}
	if (trim(f.c_perc_multa.value)=="") {
		alert('Informe o percentual da multa!!');
		f.c_perc_multa.focus();
		return;
		}
		
//  Verifica se o endereço do cadastro está devidamente preenchido
	if (!consiste_endereco_cadastro(fCAD)) return;
		
//  PARA O CASO DE TER CLICADO NO BOTÃO BACK APÓS TER CLICADO NA OPERAÇÃO EXCLUIR
	f.operacao_selecionada.value=f.operacao_selecionada_original.value;
	dATUALIZA.style.visibility="hidden";
	window.status = "Aguarde ...";
	f.submit();
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
#rb_st_ativo {
	margin: 0pt 4pt 1pt 6pt;
	}
</style>


<%	if operacao_selecionada=OP_INCLUI then
		s = "fCAD.c_num_banco.focus()"
	else
		s = "focus()"
		end if
%>
<body id="corpoPagina" onload="<%=s%>">

<center>

<!-- #include file = "../global/JanelaBuscaCEP.htm"    -->


<!--  FORMULÁRIO DE CADASTRO  -->
<br />
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
<%	if operacao_selecionada=OP_INCLUI then
		s = "Boleto - Configuração de Nova Conta do Cedente"
	else
		s = "Boleto - Consulta/Edição de Conta do Cedente"
		end if
%>
	<td align="center" valign="bottom"><span class="PEDIDO"><%=s%></span></td>
</tr>
</table>
<br>


<!--  CAMPOS DO CADASTRO  -->
<form id="fCAD" name="fCAD" method="post" action="FinCadBoletoCedenteParametrosAtualiza.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name='operacao_selecionada_original' id="operacao_selecionada_original" value='<%=operacao_selecionada%>'>
<input type="hidden" name='operacao_selecionada' id="operacao_selecionada" value='<%=operacao_selecionada%>'>
<input type="hidden" name='id_conta_corrente_selecionado' id="id_conta_corrente_selecionado" value='<%=id_conta_corrente_selecionado%>'>
<input type="hidden" name='id_boleto_cedente_selecionado' id="id_boleto_cedente_selecionado" value='<%=id_boleto_cedente_selecionado%>'>

<!-- ************   BANCO / CARTEIRA   ************ -->
<table width="649" class="Q" cellspacing="0">
	<tr>
		<%	if operacao_selecionada=OP_CONSULTA then 
			'	t_FIN_BOLETO_CEDENTE
				s = Trim("" & rs("num_banco")) 
			else 
			'	t_FIN_CONTA_CORRENTE
				s = Right(Trim("" & rs("banco")), 3)
				end if %>
		<td class="MD" width="20%" align="left">
			<p class="R">Nº BANCO</p>
			<p class="C">
				<input id="c_num_banco" name="c_num_banco" class="TA" type="text" maxlength="3" style="width:75px;" value="<%=s%>" 
					onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCAD.c_nome_banco.focus(); filtra_numerico();">
			</p>
		</td>
		<%	if operacao_selecionada=OP_CONSULTA then 
			'	t_FIN_BOLETO_CEDENTE
				s = Trim("" & rs("nome_banco")) 
			else 
			'	t_FIN_CONTA_CORRENTE
				s = ""
				if Trim("" & rs("banco")) = "237" then s = "BRADESCO"
				end if %>
		<td align="left">
			<p class="R">NOME DO BANCO</p>
			<p class="C">
				<input id="c_nome_banco" name="c_nome_banco" class="TA" type="text" maxlength="15" style="width:250px;" value="<%=s%>" 
					onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCAD.c_agencia.focus(); filtra_nome_identificador();">
			</p>
		</td>
	</tr>
</table>

<!-- ************   AGÊNCIA / CONTA   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
		<%	if operacao_selecionada=OP_CONSULTA then 
			'	t_FIN_BOLETO_CEDENTE
				s = Trim("" & rs("agencia")) 
			else 
			'	t_FIN_CONTA_CORRENTE
				s = Right(Trim("" & rs("agencia_sem_digito")),5)
				end if %>
		<td class="MD" width="20%" align="left">
			<p class="R">AGÊNCIA</p>
			<p class="C">
				<input id="c_agencia" name="c_agencia" class="TA" type="text" maxlength="5" style="width:100px;" value="<%=s%>" 
					onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCAD.c_digito_agencia.focus(); filtra_numerico();">
			</p>
		</td>
		<%	if operacao_selecionada=OP_CONSULTA then 
			'	t_FIN_BOLETO_CEDENTE
				s = Trim("" & rs("digito_agencia")) 
			else
			'	t_FIN_CONTA_CORRENTE
				s = Trim("" & rs("digito_agencia"))
				end if %>
		<td class="MD" width="10%" align="left">
			<p class="R">DÍGITO</p>
			<p class="C">
				<input id="c_digito_agencia" name="c_digito_agencia" class="TA" type="text" maxlength="1" style="width:40px;" value="<%=s%>" 
					onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCAD.c_conta.focus(); filtra_numerico();">
			</p>
		</td>
		<%	if operacao_selecionada=OP_CONSULTA then 
			'	t_FIN_BOLETO_CEDENTE
				s = Trim("" & rs("conta")) 
			else 
			'	t_FIN_CONTA_CORRENTE
				s = Right(Trim("" & rs("conta_sem_digito")),7)
				end if %>
		<td class="MD" width="20%" align="left">
			<p class="R">CONTA CORRENTE</p>
			<p class="C">
				<input id="c_conta" name="c_conta" class="TA" type="text" maxlength="7" style="width:100px;" value="<%=s%>" 
					onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCAD.c_digito_conta.focus(); filtra_numerico();">
			</p>
		</td>
		<%	if operacao_selecionada=OP_CONSULTA then 
			'	t_FIN_BOLETO_CEDENTE
				s = Trim("" & rs("digito_conta")) 
			else 
			'	t_FIN_CONTA_CORRENTE
				s = Trim("" & rs("digito_conta")) 
				end if %>
		<td class="MD" width="10%" align="left">
			<p class="R">DÍGITO</p>
			<p class="C">
				<input id="c_digito_conta" name="c_digito_conta" class="TA" type="text" maxlength="1" style="width:40px;" value="<%=s%>" 
					onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCAD.c_carteira.focus(); filtra_numerico();">
			</p>
		</td>
		<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("carteira")) else s=""%>
		<td align="left">
			<p class="R">CARTEIRA</p>
			<p class="C">
				<input id="c_carteira" name="c_carteira" class="TA" type="text" maxlength="3" style="width:75px;" value="<%=s%>" 
					onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCAD.c_codigo_empresa.focus(); filtra_numerico();">
			</p>
		</td>
	</tr>
</table>

<!-- ************   CÓDIGO DA EMPRESA / NOME DA EMPRESA   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
		<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("codigo_empresa")) else s=""%>
		<td class="MD" width="30%" align="left">
			<p class="R">CÓDIGO DA EMPRESA</p>
			<p class="C">
				<input id="c_codigo_empresa" name="c_codigo_empresa" class="TA" type="text" maxlength="20" style="width:170px;" value="<%=s%>" 
					onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCAD.c_nome_empresa.focus(); filtra_numerico();">
			</p>
		</td>
		<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("nome_empresa")) else s=""%>
		<td align="left">
			<p class="R">NOME DA EMPRESA</p>
			<p class="C">
				<input id="c_nome_empresa" name="c_nome_empresa" class="TA" type="text" maxlength="30" style="width:350px;" value="<%=s%>" 
					onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCAD.c_apelido.focus(); filtra_nome_identificador();">
			</p>
		</td>
	</tr>
</table>

<!-- ************   APELIDO / LOJA DEFAULT (BOLETO: PLANO CONTAS)   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
		<%	if operacao_selecionada=OP_CONSULTA then
				s = Trim("" & rs("apelido"))
			else
				s = ""
				end if
		 %>
		<td class="MD" width="30%" align="left">
			<p class="R">APELIDO</p>
			<p class="C">
				<input id="c_apelido" name="c_apelido" class="TA" type="text" maxlength="10" style="width:100px;" value="<%=s%>" 
					onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCAD.c_loja_default_boleto_plano_contas.focus(); filtra_nome_identificador();">
			</p>
		</td>
		<% if operacao_selecionada=OP_CONSULTA then
				s = Trim("" & rs("loja_default_boleto_plano_contas"))
			else
				s = ""
				end if
		 %>
		<td align="left">
			<p class="R">Nº LOJA PADRÃO (BOLETO: PLANO CONTAS)</p>
			<p class="C">
				<input id="c_loja_default_boleto_plano_contas" name="c_loja_default_boleto_plano_contas" class="TA" type="text" maxlength="3" style="width:100px;" value="<%=s%>" 
					onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCAD.c_juros_mora.focus(); filtra_numerico();">
			</p>
		</td>
	</tr>
</table>

<!-- ************   JUROS MORA / PERCENTUAL DE MULTA / NSU REMESSA  ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
		<%if operacao_selecionada=OP_CONSULTA then s=formata_perc(rs("juros_mora")) else s=""%>
		<td class="MD" width="30%" align="left">
			<p class="R">JUROS DE MORA AO MÊS (%)</p>
			<p class="C">
				<input id="c_juros_mora" name="c_juros_mora" class="TA" type="text" maxlength="6" style="width:100px;" value="<%=s%>" 
					onblur="this.value=formata_numero(this.value,2);"
					onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCAD.c_perc_multa.focus(); filtra_percentual();">
			</p>
		</td>
		<%if operacao_selecionada=OP_CONSULTA then s=formata_perc(rs("perc_multa")) else s=""%>
		<td class="MD" width="30%" align="left">
			<p class="R">PERC MULTA (%)</p>
			<p class="C">
				<input id="c_perc_multa" name="c_perc_multa" class="TA" type="text" maxlength="6" style="width:100px;" value="<%=s%>" 
					onblur="this.value=formata_numero(this.value,2);"
					onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCAD.c_nsu_arq_remessa.focus(); filtra_percentual();">
			</p>
		</td>
		<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("nsu_arq_remessa")) else s="0"%>
		<td align="left">
			<p class="R">Nº SEQUENCIAL DE REMESSA</p>
			<p class="C">
				<input id="c_nsu_arq_remessa" name="c_nsu_arq_remessa" class="TA" type="text" maxlength="7" style="width:100px;" value="<%=s%>" 
					onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCAD.c_qtde_dias_protestar_apos_padrao.focus(); filtra_numerico();">
			</p>
		</td>
<%
	dim st_ativo
	st_ativo=false
	if operacao_selecionada=OP_CONSULTA then
		if Cstr(rs("st_ativo")) = Cstr(COD_FIN_ST_ATIVO__ATIVO) then st_ativo=true
	elseif operacao_selecionada=OP_INCLUI then
		st_ativo=true
		end if
%>
<!-- ************   STATUS ATIVO / PROTESTAR APÓS N DIAS  ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
		<td width="30%" class="MD" align="left">
		<p class="R">STATUS</p>
		<p class="C">
			<input type="radio" id="rb_st_ativo" name="rb_st_ativo" 
				value="<%=COD_FIN_ST_ATIVO__INATIVO%>" 
				class="TA" <%if Not st_ativo then Response.Write(" checked")%>
				><span onclick="fCAD.rb_st_ativo[0].click();" 
				style="cursor:default;color:<%=finStAtivoCor(COD_FIN_ST_ATIVO__INATIVO)%>;"
				><%=finStAtivoDescricao(COD_FIN_ST_ATIVO__INATIVO)%></span
				>&nbsp;
			<input type="radio" id="rb_st_ativo" name="rb_st_ativo" 
				value="<%=COD_FIN_ST_ATIVO__ATIVO%>" 
				class="TA" <%if st_ativo then Response.Write(" checked")%>
				><span onclick="fCAD.rb_st_ativo[1].click();" 
				style="cursor:default;color:<%=finStAtivoCor(COD_FIN_ST_ATIVO__ATIVO)%>;"
				><%=finStAtivoDescricao(COD_FIN_ST_ATIVO__ATIVO)%></span
				>&nbsp;</p>
		</td>
		<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("qtde_dias_protestar_apos_padrao")) else s=""%>
		<td align="left">
			<p class="R">PROTESTAR APÓS (VALOR PADRÃO EM DIAS)</p>
			<p class="C">
				<input id="c_qtde_dias_protestar_apos_padrao" name="c_qtde_dias_protestar_apos_padrao" class="TA" type="text" maxlength="2" style="width:75px;" value="<%=s%>" 
					onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCAD.c_segunda_mensagem_padrao.focus(); filtra_numerico();">
			</p>
		</td>
	</tr>
</table>

<!-- ************   TEXTOS PADRÃO PARA OS CAMPOS DE MENSAGEM   ************ -->
<table width="649" class="QS" cellspacing="0">
	<!-- ************   2ª MENSAGEM OU SACADOR/AVALISTA (TEXTO PADRÃO)   ************ -->
	<tr>
		<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("segunda_mensagem_padrao")) else s=""%>
		<td align="left">
			<p class="R">2ª MENSAGEM OU SACADOR/AVALISTA (TEXTO PADRÃO)</p>
			<p class="C">
				<input id="c_segunda_mensagem_padrao" name="c_segunda_mensagem_padrao" class="TA" type="text" maxlength="60" style="width:635px;" value="<%=s%>" 
					onblur="this.value=trim(this.value);"
					onkeypress="if (digitou_enter(true)) fCAD.c_mensagem_1_padrao.focus(); filtra_nome_identificador();">
			</p>
		</td>
	</tr>
	<!-- ************   MENSAGEM 1 (TEXTO PADRÃO)   ************ -->
	<tr>
		<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("mensagem_1_padrao")) else s=""%>
		<td class="MC" align="left">
			<p class="R">MENSAGEM 1 (TEXTO PADRÃO)</p>
			<p class="C">
				<input id="c_mensagem_1_padrao" name="c_mensagem_1_padrao" class="TA" type="text" maxlength="80" style="width:635px;" value="<%=s%>" 
					onblur="this.value=trim(this.value);"
					onkeypress="if (digitou_enter(true)) fCAD.c_mensagem_2_padrao.focus(); filtra_nome_identificador();">
			</p>
		</td>
	</tr>
	<!-- ************   MENSAGEM 2 (TEXTO PADRÃO)   ************ -->
	<tr>
		<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("mensagem_2_padrao")) else s=""%>
		<td class="MC" align="left">
			<p class="R">MENSAGEM 2 (TEXTO PADRÃO)</p>
			<p class="C">
				<input id="c_mensagem_2_padrao" name="c_mensagem_2_padrao" class="TA" type="text" maxlength="80" style="width:635px;" value="<%=s%>" 
					onblur="this.value=trim(this.value);"
					onkeypress="if (digitou_enter(true)) fCAD.c_mensagem_3_padrao.focus(); filtra_nome_identificador();">
			</p>
		</td>
	</tr>
	<!-- ************   MENSAGEM 3 (TEXTO PADRÃO)   ************ -->
	<tr>
		<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("mensagem_3_padrao")) else s=""%>
		<td class="MC" align="left">
			<p class="R">MENSAGEM 3 (TEXTO PADRÃO)</p>
			<p class="C">
				<input id="c_mensagem_3_padrao" name="c_mensagem_3_padrao" class="TA" type="text" maxlength="80" style="width:635px;" value="<%=s%>" 
					onblur="this.value=trim(this.value);"
					onkeypress="if (digitou_enter(true)) fCAD.c_mensagem_4_padrao.focus(); filtra_nome_identificador();">
			</p>
		</td>
	</tr>
	<!-- ************   MENSAGEM 4 (TEXTO PADRÃO)   ************ -->
	<tr>
		<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("mensagem_4_padrao")) else s=""%>
		<td class="MC" align="left">
			<p class="R">MENSAGEM 4 (TEXTO PADRÃO)</p>
			<p class="C">
				<input id="c_mensagem_4_padrao" name="c_mensagem_4_padrao" class="TA" type="text" maxlength="80" style="width:635px;" value="<%=s%>" 
					onblur="this.value=trim(this.value);"
					onkeypress="if (digitou_enter(true)) fCAD.endereco.focus(); filtra_nome_identificador();">
			</p>
		</td>
	</tr>
</table>

<!-- ************   ENDEREÇO   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
	<td width="100%" align="left"><p class="R">ENDEREÇO</p><p class="C">
		<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("endereco")) else s=""%>
		<input id="endereco" name="endereco" class="TA" value="<%=s%>" maxlength="60" style="width:635px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCAD.endereco_numero.focus(); filtra_nome_identificador();"></p></td>
	</tr>
</table>

<!-- ************   Nº/COMPLEMENTO   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
	<td class="MD" width="50%" align="left"><p class="R">Nº</p><p class="C">
		<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("endereco_numero")) else s=""%>
		<input id="endereco_numero" name="endereco_numero" class="TA" value="<%=s%>" maxlength="20" style="width:310px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCAD.endereco_complemento.focus(); filtra_nome_identificador();"></p></td>
	<td align="left"><p class="R">COMPLEMENTO</p><p class="C">
		<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("endereco_complemento")) else s=""%>
		<input id="endereco_complemento" name="endereco_complemento" class="TA" value="<%=s%>" maxlength="60" style="width:310px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCAD.bairro.focus(); filtra_nome_identificador();"></p></td>
	</tr>
</table>

<!-- ************   BAIRRO/CIDADE   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
	<td class="MD" width="50%" align="left"><p class="R">BAIRRO</p><p class="C">
		<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("bairro")) else s=""%>
		<input id="bairro" name="bairro" class="TA" value="<%=s%>" maxlength="72" style="width:310px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCAD.cidade.focus(); filtra_nome_identificador();"></p></td>
	<td align="left"><p class="R">CIDADE</p><p class="C">
		<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("cidade")) else s=""%>
		<input id="cidade" name="cidade" class="TA" value="<%=s%>" maxlength="60" style="width:310px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCAD.uf.focus(); filtra_nome_identificador();"></p></td>
	</tr>
</table>

<!-- ************   UF/CEP   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
	<td class="MD" width="50%" align="left"><p class="R">UF</p><p class="C">
		<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("uf")) else s=""%>
		<input id="uf" name="uf" class="TA" value="<%=s%>" maxlength="2" size="3" onkeypress="if (digitou_enter(true) && tem_info(this.value) && uf_ok(this.value)) bATUALIZA.focus();" onblur="this.value=trim(this.value); if (!uf_ok(this.value)) {alert('UF inválida!!');this.focus();} else this.value=ucase(this.value);"></p></td>
	<td width="25%" align="left"><p class="R">CEP</p><p class="C">
		<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("cep")) else s=""%>
		<input id="cep" name="cep" readonly tabindex=-1 class="TA" value="<%=cep_formata(s)%>" maxlength="9" size="11" onkeypress="if (digitou_enter(true) && cep_ok(this.value)) bATUALIZA.focus(); filtra_cep();" onblur="if (!cep_ok(this.value)) {alert('CEP inválido!!');this.focus();} else this.value=cep_formata(this.value);"></p></td>
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
		s = "<td align='center'>" & _
				"<div name='dREMOVE' id='dREMOVE'>" & _
					"<a href='javascript:RemoveRegistro(fCAD)' title='exclui do banco de dados'>" & _
						"<img src='../botao/remover.gif' width=176 height=55 border=0>" & _
					"</a>" & _
				"</div>" & _
			"</td>"
		end if
	%><%=s%>
	<td align="right"><div name="dATUALIZA" id="dATUALIZA">
		<a name="bATUALIZA" id="bATUALIZA" href="javascript:AtualizaRegistro(fCAD)" title="atualiza o cadastro">
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