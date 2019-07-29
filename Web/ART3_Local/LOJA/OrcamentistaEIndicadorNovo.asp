<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/Global.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =============================================================
'	  O R C A M E N T I S T A E I N D I C A D O R N O V O . A S P
'     =============================================================
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
	dim s, usuario, loja, id_selecionado, operacao_selecionada, tipo_PJ_PF, i
	dim s_label, s_parametro, s_selected
	usuario = trim(Session("usuario_atual"))
	loja = Trim(Session("loja_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 
	If (loja = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))
	  
	if Not operacao_permitida(OP_LJA_CADASTRAMENTO_INDICADOR, s_lista_operacoes_permitidas) then 
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if
	
'	INDICADOR A CADASTRAR
	id_selecionado = ucase(trim(request("id_selecionado")))
	operacao_selecionada = trim(request("operacao_selecionada"))
	tipo_PJ_PF = trim(Request.Form("rb_tipo"))
	
	if (id_selecionado="") then Response.Redirect("aviso.asp?id=" & ERR_ORCAMENTISTA_INDICADOR_NAO_ESPECIFICADO) 
	if (operacao_selecionada<>OP_INCLUI) And (operacao_selecionada<>OP_CONSULTA) then Response.Redirect("aviso.asp?id=" & ERR_OPERACAO_NAO_ESPECIFICADA)
	
'	CONECTA COM O BANCO DE DADOS
	dim cn,rs,r
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	set rs = cn.Execute("SELECT * FROM t_ORCAMENTISTA_E_INDICADOR WHERE (apelido='" & id_selecionado & "')")
	if Err <> 0 then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
	
	if operacao_selecionada=OP_INCLUI then
		if Not rs.EOF then Response.Redirect("aviso.asp?id=" & ERR_ORCAMENTISTA_INDICADOR_JA_CADASTRADO)
		set r = cn.Execute("SELECT * FROM t_USUARIO WHERE (usuario = '" & id_selecionado & "')")
		if Not r.Eof then Response.Redirect("aviso.asp?id=" & ERR_ID_JA_EM_USO_POR_USUARIO)
	elseif operacao_selecionada=OP_CONSULTA then
		if rs.EOF then Response.Redirect("aviso.asp?id=" & ERR_ORCAMENTISTA_INDICADOR_NAO_CADASTRADO)
		tipo_PJ_PF = Trim("" & rs("tipo"))
		end if

%>


<%=DOCTYPE_LEGADO%>

<html>


<head>
	<title>LOJA</title>
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

<%  dim strScript
	strScript = _
		"<script language='JavaScript'>" & chr(13) & _
		"var MAX_TAM_OBS = " & Cstr(MAX_TAM_OBS_ORCAMENTISTA_INDICADOR) & ";" & chr(13) & _
		"</script>" & chr(13)
	
	Response.Write strScript
%>

<% if tipo_PJ_PF = ID_PF then %>
<script language="JavaScript" type="text/javascript">
var tipo_PJ_PF = ID_PF;
</script>
<% else %>
<script language="JavaScript" type="text/javascript">
var tipo_PJ_PF = ID_PJ;
</script>
<% end if %>

<script language="JavaScript" type="text/javascript">
var fCepPopup;

$(function () {
	// Trata o problema em que os campos do formulário são limpos após retornar à esta página c/ o history.back() pela 2ª vez quando ocorre erro de consistência
	if (trim(fCAD.c_FormFieldValues.value) != "") {
		stringToForm(fCAD.c_FormFieldValues.value, $('#fCAD'));
	}
});

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

function AtualizaItem( f ) {
var s, s_senha;

//  CNPJ/CPF + RAZÃO SOCIAL/NOME
	if (tipo_PJ_PF == ID_PF) {
		if (trim(f.razao_social_nome.value)=="") {
			alert('Preencha o nome!!');
			f.razao_social_nome.focus();
			return;
			}
		if (trim(f.cnpj_cpf.value)=='') {
			alert('Preencha o CPF!!');
			f.cnpj_cpf.focus();
			return;
			}
		if (!cpf_ok(f.cnpj_cpf.value)) {
			alert('CPF inválido!!');
			f.cnpj_cpf.focus();
			return;
			}
		if (trim(f.ie_rg.value)=='') {
			alert('Preencha o RG!!');
			f.ie_rg.focus();
			return;
			}
		}
	else {
		if (trim(f.razao_social_nome.value)=="") {
			alert('Preencha a razão social!!');
			f.razao_social_nome.focus();
			return;
			}
		if (trim(f.cnpj_cpf.value)=='') {
			alert('Preencha o CNPJ!!');
			f.cnpj_cpf.focus();
			return;
			}
		if (!cnpj_ok(f.cnpj_cpf.value)) {
			alert('CNPJ inválido!!');
			f.cnpj_cpf.focus();
			return;
			}
		if (trim(f.ie_rg.value)=='') {
			alert('Preencha a IE!!');
			f.ie_rg.focus();
			return;
			}
	}

//  NOME FANTASIA
	if (trim(f.c_nome_fantasia.value) == '') {
	    alert('Preencha o nome fantasia!!');
	    f.c_nome_fantasia.focus();
	    return;
	}

//  CEP
	if (trim(f.cep.value)=='') {
		alert('Preencha o endereço usando o cadastro de CEP!!');
		f.cep.focus();
		return;
		}
		
//  CEP VÁLIDO?
	if (!cep_ok(f.cep.value)) {
		alert('CEP inválido!!');
		f.cep.focus();
		return;
		}
		
//  ENDEREÇO
	if (trim(f.endereco.value)=='') {
		alert('Preencha o endereco!!');
		f.endereco.focus();
		return;
		}

//  Nº DO ENDEREÇO
	if (trim(f.endereco_numero.value)=='') {
		alert('Preencha o número do endereco!!');
		f.endereco_numero.focus();
		return;
		}

//  BAIRRO
	if (trim(f.bairro.value)=='') {
		alert('Preencha o bairro!!');
		f.bairro.focus();
		return;
		}

//  CIDADE
	if (trim(f.cidade.value)=='') {
		alert('Preencha a cidade!!');
		f.cidade.focus();
		return;
		}
		
//  UF
	if (!uf_ok(f.uf.value)) {
		alert('UF inválida!!');
		f.uf.focus();
		return;
		}

//  DDD
	if (trim(f.ddd.value)=='') {
		alert('Preencha o DDD!!');
		f.ddd.focus();
		return;
		}

//  TELEFONE
	if (trim(f.telefone.value)=='') {
		alert('Preencha o telefone!!');
		f.telefone.focus();
		return;
		}

//  FAX
	if (trim(f.fax.value)=='') {
		alert('Preencha o fax!!');
		f.fax.focus();
		return;
		}

//  DDD CEL
	if (trim(f.ddd_cel.value)=='') {
		alert('Preencha o DDD do celular!!');
		f.ddd_cel.focus();
		return;
		}

//  CEL
	if (trim(f.tel_cel.value)=='') {
		alert('Preencha o celular!!');
		f.tel_cel.focus();
		return;
		}

//  CONTATO
	if (trim(f.contato.value)=='') {
		alert('Preencha o contato!!');
		f.contato.focus();
		return;
		}
		
//  TELEFONE / FAX VÁLIDOS?
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
//  Nº CELULAR VÁLIDO?
	if (!ddd_ok(f.ddd_cel.value)) {
		alert('DDD do celular é inválido!!');
		f.ddd_cel.focus();
		return;
		}
	if (!telefone_ok(f.tel_cel.value)) {
		alert('Telefone celular inválido!!');
		f.tel_cel.focus();
		return;
		}
	if ((trim(f.ddd_cel.value)!="")||(trim(f.tel_cel.value)!="")) {
		if (trim(f.ddd_cel.value)=="") {
			alert('Preencha o DDD do celular!!');
			f.ddd_cel.focus();
			return;
			}
		if (trim(f.tel_cel.value)=="") {
			alert('Preencha o telefone celular!!');
			f.tel_cel.focus();
			return;
			}
		}
		
    //  DADOS BANCÁRIOS
	if ((trim(f.banco.value) != "") || (trim(f.agencia.value) != "") || (trim(f.conta.value) != "") || (trim(f.favorecido.value) != "")) {
	    if (trim(f.banco.value) == "") {
	        alert('Preencha o número do banco!!');
	        f.banco.focus();
	        return;
	    }
	    if (trim(f.agencia.value) == "") {
	        alert('Preencha o número da agência!!');
	        f.agencia.focus();
	        return;
	    }

	    if (trim(f.conta.value) == "") {
	        alert('Preencha o número da conta!!');
	        f.conta.focus();
	        return;
	    }
	    if (trim(f.banco.value) != "745") {
	        if (trim(f.conta_dv.value) == "") {
	            alert('Preencha o dígito verificador da conta!!');
	            f.conta_dv.focus();
	            return;
	        }
	    }
	    if (trim(f.banco.value) == "104") {
	        if (trim(f.tipo_operacao.value) == "") {
	            alert('Contas da Caixa Econômica Federal exigem preenchimento do tipo de operação!!')
	            f.tipo_operacao.focus();
	            return;
	        }
	    }
	    if (trim(f.tipo_conta.value) == "") {
	        alert('Preencha o tipo de conta!!');
	        f.tipo_conta.focus();
	        return;
	    }
	    if (trim(f.favorecido.value) == "") {
	        alert('Preencha o favorecido!!');
	        f.favorecido.focus();
	        return;
	    }
	}

    //CNPF/CPF DO FAVORECIDO  
	if ((trim(f.favorecido_cnpjcpf.value) == '')) {
	    alert('Preencha o CPF/CNPJ do favorecido!');
	    f.favorecido_cnpjcpf.focus();
	    return;
	}
	if (cnpj_cpf_ok(f.favorecido_cnpjcpf.value) == false) {
	    alert('CPF/CNPJ inválido!');
	    f.favorecido_cnpjcpf.focus();
	    return;
	}

    //  SENHA
	if (f.rb_acesso[0].checked) {
	    s_senha = f.senha.value;
	    if (s_senha == "") {
	        alert('Preencha a senha!!');
	        f.senha.focus();
	        return;
	    }

	    if (s_senha.length < 5) {
	        alert('A senha deve possuir no mínimo 5 caracteres!!');
	        f.senha.focus();
	        return;
	    }

	    if (s_senha != f.senha2.value) {
	        alert('A confirmação da senha não confere!!');
	        f.senha2.focus();
	        return;
	    }
	}


//  E-MAIL
	if ((trim(f.c_email.value) == '') && (trim(f.c_email2.value) == '') && (trim(f.c_email3.value) == '')) {
		alert('Informe no mínimo um endereço de email!!');
		f.c_email.focus();
		return;
		}

//  E-MAIL VÁLIDO?
	if (trim(f.c_email.value) != "") {
		if (!email_ok(f.c_email.value)) {
			alert('Email inválido!!');
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

	if (trim(f.c_email3.value) != "") {
		if (!email_ok(f.c_email3.value)) {
			alert("Email inválido!!");
			f.c_email3.focus();
			return;
			}
		}
	
//  FORMA COMO CONHECEU A BONSHOP
	if (trim(f.c_forma_como_conheceu_codigo.value) == '') {
		alert('Selecione a forma como conheceu a Bonshop!!');
		f.c_forma_como_conheceu_codigo.focus();
		return;
	}
	
	s = "" + f.c_obs.value;
	if (s.length > MAX_TAM_OBS) {
		alert('Conteúdo de "Observações" excede em ' + (s.length-MAX_TAM_OBS) + ' caracteres o tamanho máximo de ' + MAX_TAM_OBS + '!!');
		f.c_obs.focus();
		return;
		}

	fCAD.c_FormFieldValues.value = formToString($("#fCAD"));

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

<style TYPE="text/css">
#loja,#vendedor {
	margin: 4pt 0pt 4pt 10pt;
	vertical-align: top;
	}
#rb_acesso,#rb_status {
	margin-left:10pt;
	}
#rb_estabelecimento 
{
	margin-left:10pt;
}
#lbl_estabelecimento 
{
	font-size:9pt;
}
</style>

<%	if operacao_selecionada=OP_INCLUI then
		s = "fCAD.razao_social_nome.focus();"
	else
		s = "focus();"
		end if
%>
<body id="corpoPagina" onload="<%=s%>">

<center>

<!-- #include file = "../global/JanelaBuscaCEP.htm"    -->


<!--  CADASTRO DO INDICADOR -->
<br />
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
<%	if operacao_selecionada=OP_INCLUI then
		s = "Cadastro de Novo Indicador"
	else
		s = "Consulta/Edição de Indicador Cadastrado"
		end if
%>
	<td align="center" valign="bottom"><span class="PEDIDO"><%=s%></span></td>
</tr>
</table>
<br>


<!--  CAMPOS DO CADASTRO  -->
<form id="fCAD" name="fCAD" method="post" action="OrcamentistaEIndicadorNovoConfirma.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="operacao_selecionada" id="operacao_selecionada" value='<%=operacao_selecionada%>'>
<input type="hidden" name="id_selecionado" id="id_selecionado" value='<%=id_selecionado%>'>
<input type="hidden" name="tipo_PJ_PF" id="tipo_PJ_PF" value='<%=tipo_PJ_PF%>'>
<input type="hidden" name="c_FormFieldValues" id="c_FormFieldValues" value="" />

<!-- ************   NOME/RAZÃO SOCIAL   ************ -->
<table width="649" class="Q" cellSpacing="0">
	<tr>
		<td class="MD" width="30%" align="left"><p class="R">APELIDO</p><p class="C"><input id="c_apelido" name="c_apelido" class="TA" value="<%=id_selecionado%>" readonly size="25" style="text-align:center; color:#0000ff"></p></td>
<%if tipo_PJ_PF=ID_PJ then s_label = "RAZÃO SOCIAL" else s_label="NOME" %>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("razao_social_nome")) else s=""%>
		<td width="70%" align="left"><p class="R"><%=s_label%></p><p class="C"><input id="razao_social_nome" name="razao_social_nome" class="TA" type="text" maxlength="60" size="60" value="<%=s%>" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCAD.c_responsavel_principal.focus(); filtra_nome_identificador();"></p></td>
	</tr>
</table>

<!-- ************  RESPONSÁVEL PRINCIPAL   ************ -->
<table width="649" class="QS" cellSpacing="0">
	<tr>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("responsavel_principal")) else s=""%>
		<td align="left"><p class="R">PRINCIPAL</p><p class="C"><input id="c_responsavel_principal" name="c_responsavel_principal" class="TA" type="text" maxlength="60" size="60" value="<%=s%>" onkeypress="if (digitou_enter(true)) fCAD.c_nome_fantasia.focus();"></p></td>
	</tr>
</table>

<!-- ************   NOME FANTASIA   ************ -->
<table width="649" class="QS" cellSpacing="0">
	<tr>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("nome_fantasia")) else s=""%>
		<td align="left"><p class="R">NOME FANTASIA</p><p class="C"><input id="c_nome_fantasia" name="c_nome_fantasia" class="TA" type="text" maxlength="60" size="60" value="<%=s%>" onkeypress="if (digitou_enter(true)) fCAD.cnpj_cpf.focus(); filtra_nome_identificador();"></p></td>
	</tr>
</table>

<!-- ************   CNPJ/CPF + IE/RG   ************ -->
<table width="649" class="QS" cellSpacing="0">
	<tr>
<%if tipo_PJ_PF=ID_PJ then s_label = "CNPJ" else s_label="CPF" %>
<%if operacao_selecionada=OP_CONSULTA then s=cnpj_cpf_formata(Trim("" & rs("cnpj_cpf"))) else s=""%>
	<td class="MD" width="50%" align="left"><p class="R"><%=s_label%></p><p class="C">
		<input id="cnpj_cpf" name="cnpj_cpf" class="TA" value="<%=s%>" maxlength="18" size="24" 
		<% if tipo_PJ_PF = ID_PJ then %>
			onblur="if (!cnpj_ok(this.value)) {alert('CNPJ inválido!!'); this.focus();} else this.value=cnpj_formata(this.value);" onkeypress="if (digitou_enter(true) && cnpj_ok(this.value)) fCAD.ie_rg.focus(); filtra_cnpj();"
		<% else %>
			onblur="if (!cpf_ok(this.value)) {alert('CPF inválido!!'); this.focus();} else this.value=cpf_formata(this.value);" onkeypress="if (digitou_enter(true) && cpf_ok(this.value)) fCAD.ie_rg.focus(); filtra_cpf();"
		<% end if %>
		></p></td>
<%if tipo_PJ_PF=ID_PJ then s_label = "IE" else s_label="RG" %>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("ie_rg")) else s=""%>
		<td width="50%" align="left"><p class="R"><%=s_label%></p><p class="C"><input id="ie_rg" name="ie_rg" class="TA" type="text" maxlength="20" size="25" value="<%=s%>" onkeypress="if (digitou_enter(true)) fCAD.endereco.focus(); filtra_nome_identificador();"></p></td>
	</tr>
</table>

<!-- ************   ENDEREÇO   ************ -->
<table width="649" class="QS" cellSpacing="0">
	<tr>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("endereco")) else s=""%>
		<td width="100%" align="left"><p class="R">ENDEREÇO</p><p class="C"><input id="endereco" name="endereco" class="TA" value="<%=s%>" maxlength="60" style="width:635px;" onkeypress="if (digitou_enter(true)) fCAD.endereco_numero.focus(); filtra_nome_identificador();"></p></td>
	</tr>
</table>

<!-- ************   Nº/COMPLEMENTO   ************ -->
<table width="649" class="QS" cellSpacing="0">
	<tr>
	<td class="MD" width="50%" align="left"><p class="R">Nº</p><p class="C">
		<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("endereco_numero")) else s=""%>
		<input id="endereco_numero" name="endereco_numero" class="TA" value="<%=s%>" maxlength="20" style="width:310px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCAD.endereco_complemento.focus(); filtra_nome_identificador();"></p></td>
	<td align="left"><p class="R">COMPLEMENTO</p><p class="C">
		<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("endereco_complemento")) else s=""%>
		<input id="endereco_complemento" name="endereco_complemento" class="TA" value="<%=s%>" maxlength="60" style="width:310px;" onkeypress="if (digitou_enter(true)) fCAD.bairro.focus(); filtra_nome_identificador();"></p></td>
	</tr>
</table>

<!-- ************   BAIRRO/CIDADE   ************ -->
<table width="649" class="QS" cellSpacing="0">
	<tr>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("bairro")) else s=""%>
		<td width="50%" class="MD" align="left"><p class="R">BAIRRO</p><p class="C"><input id="bairro" name="bairro" class="TA" value="<%=s%>" maxlength="72" style="width:310px;" onkeypress="if (digitou_enter(true)) fCAD.cidade.focus(); filtra_nome_identificador();"></p></td>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("cidade")) else s=""%>
		<td width="50%" align="left"><p class="R">CIDADE</p><p class="C"><input id="cidade" name="cidade" class="TA" value="<%=s%>" maxlength="60" style="width:310px;" onkeypress="if (digitou_enter(true)) fCAD.uf.focus(); filtra_nome_identificador();"></p></td>
	</tr>
</table>

<!-- ************   UF/CEP   ************ -->
<table width="649" class="QS" cellSpacing="0">
	<tr>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("uf")) else s=""%>
		<td class="MD"  width="50%" align="left"><p class="R">UF</p><p class="C"><input id="uf" name="uf" class="TA" value="<%=s%>" maxlength="2" size="3" onkeypress="if (digitou_enter(true) && uf_ok(this.value)) fCAD.ddd.focus();" onblur="this.value=trim(this.value); if (!uf_ok(this.value)) {alert('UF inválida!!');this.focus();} else this.value=ucase(this.value);"></p></td>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("cep")) else s=""%>
		<td width="25%" align="left"><p class="R">CEP</p><p class="C"><input id="cep" name="cep" readonly tabindex=-1 class="TA" value="<%=cep_formata(s)%>" maxlength="9" size="11" onkeypress="if (digitou_enter(true) && cep_ok(this.value)) fCAD.ddd.focus(); filtra_cep();" onblur="if (!cep_ok(this.value)) {alert('CEP inválido!!');this.focus();} else this.value=cep_formata(this.value);"></p></td>
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

<!-- ************   DDD/TELEFONE/FAX/NEXTEL   ************ -->
<table width="649" class="QS" cellSpacing="0">
	<tr>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("ddd")) else s=""%>
		<td width="15%" class="MD" align="left"><p class="R">DDD</p><p class="C"><input id="ddd" name="ddd" class="TA" value="<%=s%>" maxlength="4" size="5" onkeypress="if (digitou_enter(true) && ddd_ok(this.value)) fCAD.telefone.focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inválido!!');this.focus();}"></p></td>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("telefone")) else s=""%>
		<td width="25%" class="MD" align="left"><p class="R">TELEFONE</p><p class="C"><input id="telefone" name="telefone" class="TA" value="<%=telefone_formata(s)%>" maxlength="11" size="12" onkeypress="if (digitou_enter(true) && telefone_ok(this.value)) fCAD.fax.focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('Telefone inválido!!');this.focus();} else this.value=telefone_formata(this.value);"></p></td>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("fax")) else s=""%>
		<td width="25%" class="MD" align="left"><p class="R">FAX</p><p class="C"><input id="fax" name="fax" class="TA" value="<%=telefone_formata(s)%>" maxlength="11" size="12" onkeypress="if (digitou_enter(true) && telefone_ok(this.value)) fCAD.c_nextel.focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('Fax inválido!!');this.focus();} else this.value=telefone_formata(this.value);"></p></td>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("nextel")) else s=""%>
		<td align="left"><p class="R">NEXTEL</p><p class="C"><input id="c_nextel" name="c_nextel" class="TA" value="<%=s%>" maxlength="15" size="12" onkeypress="if (digitou_enter(true)) fCAD.ddd_cel.focus(); filtra_nextel();" onblur="this.value=trim(this.value);"></p></td>
	</tr>
</table>

<!-- ************   TEL CEL / CONTATO   ************ -->
<table width="649" class="QS" cellSpacing="0">
	<tr>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("ddd_cel")) else s=""%>
		<td width="15%" class="MD" align="left" nowrap><p class="R">DDD (CEL)</p><p class="C"><input id="ddd_cel" name="ddd_cel" class="TA" value="<%=s%>" maxlength="2" size="3" onkeypress="if (digitou_enter(true) && ddd_ok(this.value)) fCAD.tel_cel.focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inválido!!');this.focus();}"></p></td>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("tel_cel")) else s=""%>
		<td width="25%" class="MD" align="left"><p class="R">TELEFONE (CEL)</p><p class="C"><input id="tel_cel" name="tel_cel" class="TA" value="<%=telefone_formata(s)%>" maxlength="10" size="11" onkeypress="if (digitou_enter(true) && telefone_ok(this.value)) fCAD.contato.focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('Telefone inválido!!');this.focus();} else this.value=telefone_formata(this.value);"></p></td>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("contato")) else s=""%>
		<td align="left"><p class="R">CONTATO</p><p class="C"><input id="contato" name="contato" class="TA" value="<%=s%>" maxlength="40" size="55" onkeypress="if (digitou_enter(true)) fCAD.banco.focus(); filtra_nome_identificador();"></p></td>
	</tr>
</table>

<!-- ************   BANCO/AGÊNCIA/CONTA   ************ -->
<table width="649" class="QS" cellSpacing="0">
	<tr>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("banco")) else s=""%>
		<td width="15%" class="MD" nowrap align="left"><p class="R">BANCO</p><p class="C"><input id="banco" name="banco" class="TA" value="<%=s%>" maxlength="4" size="3" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCAD.agencia.focus(); filtra_numerico();" onblur="this.value=trim(this.value);tipoOperacao();"></p></td>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("agencia")) else s=""%>
		<td width="17%" class="MD" align="left"><p class="R">AGÊNCIA</p><p class="C"><input id="agencia" name="agencia" class="TA" value="<%=s%>" maxlength="8" size="5" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCAD.agencia_dv.focus(); filtra_agencia_bancaria();" onblur="this.value=trim(this.value);"></p></td>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("agencia_dv")) else s=""%>
		<td width="5%" class="MD" align="left"><p class="R">DÍG.</p><p class="C"><input id="agencia_dv" name="agencia_dv" class="TA" value="<%=s%>" maxlength="1" size="1" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCAD.conta.focus();" onblur="this.value=trim(this.value);"></p></td>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("conta_operacao")) else s=""%>
		<td width="15%" class="MD" align="left"><p class="R">TIPO OPERAÇÃO</p><p class="C"><input id="tipo_operacao" name="tipo_operacao" class="TA" value="<%=s%>" maxlength="3" size="12" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCAD.tipo_conta.focus();" onblur="this.value=trim(this.value);"></p></td>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("conta")) else s=""%>
		<td width="17%" class="MD" align="left"><p class="R">CONTA</p><p class="C"><input id="conta" name="conta" class="TA" value="<%=s%>" maxlength="12" size="12" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCAD.conta_dv.focus(); filtra_conta_bancaria();" onblur="this.value=trim(this.value);"></p></td>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("conta_dv")) else s=""%>
		<td width="5%" class="MD" align="left"><p class="R">DÍG.</p><p class="C"><input id="conta_dv" name="conta_dv" class="TA" value="<%=s%>" maxlength="2" size="1" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCAD.tipo_operacao.focus();" onblur="this.value=trim(this.value);"></p></td>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("tipo_conta")) else s=""%>
		<td width="15%" align="left"><p class="R">TIPO CONTA</p><p class="C">
            <%s_selected="" %>
            <select name="tipo_conta" id="tipo_conta">
                <%if s="" then  s_selected=" selected"%>
                <option value=""<%=s_selected%>>&nbsp;</option>
                <%s_selected=""
                    if s="C" then s_selected=" selected" %>
                <option value="C"<%=s_selected%>>Corrente</option>
                <%s_selected=""
                    if s="P" then s_selected=" selected" %>
                <option value="P"<%=s_selected%>>Poupança</option>
            </select> </p></td>
	</tr>
</table>

<!-- ************   FAVORECIDO    *******************  -->
<table width="649" class="QS" cellspacing="0">
    <tr>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("favorecido")) else s=""%>
		<td width="1%" align="left"><p class="R">FAVORECIDO</p><p class="C"><input id="favorecido" name="favorecido" class="TA" value="<%=s%>" maxlength="40" size="80" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCAD.favorecido_cnpjcpf.focus(); filtra_nome_identificador();" onblur="this.value=trim(this.value);"></p></td>
    </tr>
</table>

<!-- ************   CPF CNPJ FAVORECIDO/ SENHA / CONFIRMAÇÃO DA SENHA   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
<%if operacao_selecionada=OP_CONSULTA then s=cnpj_cpf_formata(Trim("" & rs("favorecido_cnpj_cpf"))) else s="" %>
 <td class="MD" width="40%" align="left"><p class="R">CPF/CNPJ DO FAVORECIDO</p><p class="C"><input id="favorecido_cnpjcpf" name="favorecido_cnpjcpf" class="TA" type="text" maxlength="18" size="25" value="<%=s%>" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCAD.senha.focus();"
        			onblur="if (retorna_so_digitos(this.value).length==14) { this.value=cnpj_formata(this.value);} else if (retorna_so_digitos(this.value).length==11){ this.value=cpf_formata(this.value);} else alert('Formato de CPF/CNPJ inválido!');"></p></td>

		<td class="MD" width="30%" align="left"><p class="R">SENHA</p><p class="C"><input id="senha" name="senha" class="TA" type="password" maxlength="15" size="18" value="" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCAD.senha2.focus();"></p></td>
		<td width="30%" align="left"><p class="R">SENHA (CONFIRMAÇÃO)</p><p class="C"><input id="senha2" name="senha2" class="TA" type="password" maxlength="15" size="18" value="" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCAD.loja.focus();"></p></td>
	</tr>
</table>

<!-- ************   ACESSO AO SISTEMA/STATUS   ************ -->
<table width="649" class="QS" cellSpacing="0">
	<tr>
		<td width="35%" align="left"><p class="R">ACESSO AO SISTEMA</p><p class="C">
			<input type="radio" id="rb_acesso_liberado" name="rb_acesso" value="1" 
				class="TA"
				><span onclick="fCAD.rb_acesso[0].click();" style="cursor:default; color:#006600">Liberado</span>
			<br><input type="radio" id="rb_acesso_bloqueado" name="rb_acesso" value="0" 
				class="TA"
				><span onclick="fCAD.rb_acesso[1].click();" style="cursor:default; color:#ff0000">Bloqueado</span>
			</p></td>
        </tr>
    </table>

<!-- ************   E-MAILS   ************ -->
<table width="649" class="QS" cellSpacing="0">
	<tr>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("email")) else s=""%>
		<td align="left"><p class="R">E-MAIL (1)</p><p class="C">
			<input id="c_email" name="c_email" class="TA" value="<%=s%>" maxlength="60" 
			style="text-align:left;" size="74"
			onkeypress="if (digitou_enter(true)) fCAD.c_email2.focus(); filtra_email();"
			onblur="this.value=trim(this.value);">
		</p></td>
	</tr>
	<tr>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("email2")) else s=""%>
		<td class="MC" align="left"><p class="R">E-MAIL (2)</p><p class="C">
			<input id="c_email2" name="c_email2" class="TA" value="<%=s%>" maxlength="60" 
			style="text-align:left;" size="74"
			onkeypress="if (digitou_enter(true)) fCAD.c_email3.focus(); filtra_email();"
			onblur="this.value=trim(this.value);">
		</p></td>
	</tr>
	<tr>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("email3")) else s=""%>
		<td class="MC" align="left"><p class="R">E-MAIL (3)</p><p class="C">
			<input id="c_email3" name="c_email3" class="TA" value="<%=s%>" maxlength="60" 
			style="text-align:left;" size="74"
			onkeypress="if (digitou_enter(true)) fCAD.c_obs.focus(); filtra_email();"
			onblur="this.value=trim(this.value);">
		</p></td>
	</tr>
</table>

<!-- ************   TIPO DE ESTABELECIMENTO   ************ -->
<table width="649" class="QS" cellSpacing="0">
	<tr>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("tipo_estabelecimento")) else s=""%>
		<td width="100%" style="padding-bottom:4px;" align="left">
		<p class="R">ESTABELECIMENTO</p>
		<input type="radio" id="rb_estabelecimento_casa" name="rb_estabelecimento" value="<%=COD_PARCEIRO_TIPO_ESTABELECIMENTO__CASA%>" class="TA"<%if s = COD_PARCEIRO_TIPO_ESTABELECIMENTO__CASA then Response.Write(" checked")%>><span id="lbl_estabelecimento" onclick="fCAD.rb_estabelecimento[0].click();" style="cursor:default;" class="C">Casa</span>
		<br><input type="radio" id="rb_estabelecimento_escritorio" name="rb_estabelecimento" value="<%=COD_PARCEIRO_TIPO_ESTABELECIMENTO__ESCRITORIO%>" class="TA"<%if s = COD_PARCEIRO_TIPO_ESTABELECIMENTO__ESCRITORIO then Response.Write(" checked")%>><span id="lbl_estabelecimento" onclick="fCAD.rb_estabelecimento[1].click();" style="cursor:default;" class="C">Escritório</span>
		<br><input type="radio" id="rb_estabelecimento_loja" name="rb_estabelecimento" value="<%=COD_PARCEIRO_TIPO_ESTABELECIMENTO__LOJA%>" class="TA"<%if s = COD_PARCEIRO_TIPO_ESTABELECIMENTO__LOJA then Response.Write(" checked")%>><span id="lbl_estabelecimento" onclick="fCAD.rb_estabelecimento[2].click();" style="cursor:default;" class="C">Loja</span>
		<br><input type="radio" id="rb_estabelecimento_oficina" name="rb_estabelecimento" value="<%=COD_PARCEIRO_TIPO_ESTABELECIMENTO__OFICINA%>" class="TA"<%if s = COD_PARCEIRO_TIPO_ESTABELECIMENTO__OFICINA then Response.Write(" checked")%>><span id="lbl_estabelecimento" onclick="fCAD.rb_estabelecimento[3].click();" style="cursor:default;" class="C">Oficina</span>
		
		</td>
	</tr>
</table>

<!-- ************   FORMA COMO CONHECEU A BONSHOP   ************ -->
<table width="649" class="QS" cellSpacing="0">
	<tr>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("forma_como_conheceu_codigo")) else s=""%>
		<td align="left"><p class="R">FORMA COMO CONHECEU A BONSHOP</p><p class="C">
			<select id='c_forma_como_conheceu_codigo' name='c_forma_como_conheceu_codigo' style='width:490px;margin-left:4pt;margin-top:4pt;margin-bottom:4pt;' onkeyup='if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;'>
			<%=codigo_descricao_monta_itens_select(GRUPO_T_CODIGO_DESCRICAO__CAD_ORCAMENTISTA_E_INDICADOR__FORMA_COMO_CONHECEU, s)%>
			</select>
		</p></td>
	</tr>
</table>

<!-- ************   VENDEDORES   **************** -->
<table width="649" class="QS" cellspacing="0">
	<tr>
		<td align="left" class="MB" colspan="2"><p class="R">VENDEDORES</p></td>
	</tr>
    <tr>
        <td align="left"><p class="R" style="margin-bottom:3px;margin-top:3px">NOME</p></td>
    </tr>
<% for i = 1 to CADASTRO_INDICADOR_QTDE_MAX_VENDEDORES %>
    <tr>
        <td align="left" width="40%">
            <input id="c_indicador_contato_<%=i%>" name="c_indicador_contato" class="TA" value="" style="text-align: left;margin-left: 5px;border:1px solid #c0c0c0;" maxlength="60" size="40" />
            <input type="hidden" name="contato_id" id="contato_id_<%=i%>" value="" />

        </td>
    </tr>
<% next %>
</table>

<!-- ************   OBS   ************ -->
<table width="649" class="QS" cellSpacing="0">
	<tr>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("obs")) else s=""%>
		<td align="left"><p class="R">OBSERVAÇÕES</p><p class="C">
			<textarea name="c_obs" id="c_obs" class="PLLe" rows="<%=Cstr(MAX_LINHAS_OBS_ORCAMENTISTA_INDICADOR)%>" 
				style="width:635px;margin-left:1pt;" onkeypress="limita_tamanho(this,MAX_TAM_OBS);" onblur="this.value=trim(this.value);"
				><%=s%></textarea>
		</p></td>
	</tr>
</table>

<!-- ************   SEPARADOR   ************ -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc" align="left">&nbsp;</td></tr>
</table>
<br>


<table class="notPrint" width="649" cellSpacing="0">
<tr>
	<td align="left"><a href="javascript:history.back()" title="cancela o cadastramento">
		<img src="../botao/cancelar.gif" width="176" height="55" border="0"></a></td>
	<td align="right"><div name="dATUALIZA" id="dATUALIZA">
		<a name="bATUALIZA" id="bATUALIZA" href="javascript:AtualizaItem(fCAD)" title="confirma o cadastro do indicador">
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