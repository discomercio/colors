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
'	  O R C A M E N T I S T A E I N D I C A D O R E D I T A . A S P
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
	
	Const WIDTH_MAIN_TABLE = 689

'	EXIBIÇÃO DE BOTÕES DE PESQUISA DE CEP
	dim blnPesquisaCEPAntiga, blnPesquisaCEPNova
	
	blnPesquisaCEPAntiga = False
	blnPesquisaCEPNova = True
	
	dim alerta
	alerta = ""
	
'	OBTEM O ID
	dim s, usuario, id_selecionado, operacao_selecionada, tipo_PJ_PF, url_origem, sid, i
	dim s_checked, s_ckb_value, s_ckb_id, s_span_id, s_color, s_lista_id_forma_pagto, s_selected
	dim s_label, s_parametro, chave, senha_descripto, s_disabled
    dim cont, inc
	usuario = trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))
	
'	ORÇAMENTISTA/INDICADOR A EDITAR
	id_selecionado = ucase(trim(request("id_selecionado")))
	operacao_selecionada = trim(request("operacao_selecionada"))
	tipo_PJ_PF = trim(Request.Form("rb_tipo"))
	
	if (id_selecionado="") then Response.Redirect("aviso.asp?id=" & ERR_ORCAMENTISTA_INDICADOR_NAO_ESPECIFICADO) 
	if (operacao_selecionada<>OP_INCLUI) And (operacao_selecionada<>OP_CONSULTA) then Response.Redirect("aviso.asp?id=" & ERR_OPERACAO_NAO_ESPECIFICADA)
	
'	FOI UM RELATÓRIO QUE ORIGINOU A EDIÇÃO DO INDICADOR?
	dim pagina_relatorio_originou_edicao
	pagina_relatorio_originou_edicao = Trim(Request.Form("pagina_relatorio_originou_edicao"))

'	CONECTA COM O BANCO DE DADOS
	dim cn,rs,r,t, rs2, msg_erro
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	dim vCodDescrOpDadosBancarios, sDescricao, blnDescricaoCadastrada
	call carrega_em_vetor_t_codigo_descricao(GRUPO_T_CODIGO_DESCRICAO__ORCAMENTISTA_INDICADOR__OP_DADOS_BANCARIOS, vCodDescrOpDadosBancarios)

	dim blnUsuarioDeptoFinanceiro, vDeptoSetorUsuario
	blnUsuarioDeptoFinanceiro = False
	if alerta = "" then
		if Not obtem_Usuario_x_DeptoSetor(usuario, vDeptoSetorUsuario, msg_erro) then
			alerta=texto_add_br(alerta)
			alerta = alerta & msg_erro
		else
			for i=LBound(vDeptoSetorUsuario) to UBound(vDeptoSetorUsuario)
				if (vDeptoSetorUsuario(i).StInativo = 0) then
					if (vDeptoSetorUsuario(i).Id = ID_DEPTO_SETOR__FIN_FINANCEIRO) Or (vDeptoSetorUsuario(i).Id = ID_DEPTO_SETOR__FIN_CREDITO) then
						blnUsuarioDeptoFinanceiro = True
						exit for
						end if
					end if
				next
			end if
		end if

	if (Not operacao_permitida(OP_CEN_CADASTRO_ORCAMENTISTAS_E_INDICADORES, s_lista_operacoes_permitidas)) And _
	   (Not operacao_permitida(OP_CEN_REL_CHECAGEM_NOVOS_PARCEIROS, s_lista_operacoes_permitidas)) And _
	   (Not blnUsuarioDeptoFinanceiro) then 
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	set rs = cn.Execute("SELECT * FROM t_ORCAMENTISTA_E_INDICADOR WHERE (apelido='" & id_selecionado & "')")
	if Err <> 0 then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)

	dim blnChecadoStatusBloqueado
	blnChecadoStatusBloqueado=True
	
	if operacao_selecionada=OP_INCLUI then
		if Not rs.EOF then Response.Redirect("aviso.asp?id=" & ERR_ORCAMENTISTA_INDICADOR_JA_CADASTRADO)
		set r = cn.Execute("SELECT * FROM t_USUARIO WHERE (usuario = '" & id_selecionado & "')")
		if Not r.Eof then Response.Redirect("aviso.asp?id=" & ERR_ID_JA_EM_USO_POR_USUARIO)
		blnChecadoStatusBloqueado=False
	elseif operacao_selecionada=OP_CONSULTA then
		if rs.EOF then Response.Redirect("aviso.asp?id=" & ERR_ORCAMENTISTA_INDICADOR_NAO_CADASTRADO)
		tipo_PJ_PF = Trim("" & rs("tipo"))
		if CLng(rs("checado_status"))=0 then blnChecadoStatusBloqueado=False
		end if

	dim r_loja_indicador
	if alerta = "" then
		set r_loja_indicador = New cl_LOJA
		if operacao_selecionada <> OP_INCLUI then
			if Not x_loja_bd(Trim("" & rs("loja")), r_loja_indicador) then
				alerta=texto_add_br(alerta)
				alerta=alerta & "A loja cadastrada para o indicador (" & Trim("" & rs("loja")) & ") não foi encontrada"
				end if
			end if
		end if

	dim blnVisivelIdMagentoB2B
	blnVisivelIdMagentoB2B = False
	if alerta = "" then
		if operacao_selecionada=OP_INCLUI then
			blnVisivelIdMagentoB2B = True
		else
			if (r_loja_indicador.unidade_negocio = COD_UNIDADE_NEGOCIO_LOJA__AC) _
				Or ((Trim("" & rs("id_magento_b2b")) <> "") And (Trim("" & rs("id_magento_b2b")) <> "0")) then
				blnVisivelIdMagentoB2B = True
				end if
			end if
		end if 'if alerta = ""

    url_origem = Request("url_origem")

	dim strJS
	strJS = "<script language='JavaScript'>" & chr(13) & _
			"var vOpDadosBancarios = new Array();" & chr(13) & _
			"vOpDadosBancarios[0] = new oOpDadosBancarios('','',0,'');" & chr(13)

	for i=LBound(vCodDescrOpDadosBancarios) to UBound(vCodDescrOpDadosBancarios)
		with vCodDescrOpDadosBancarios(i)
			if Trim("" & .codigo) <> "" then
				strJS = strJS & _
						"vOpDadosBancarios[vOpDadosBancarios.length] = new oOpDadosBancarios(" & _
							"'" & .grupo & "'" & _
							",'" & .codigo & "'" & _
							"," & CStr(.st_inativo) & _
							",'" & .descricao & "'" & _
						");" & chr(13)
				end if
			end with
		next

	if blnUsuarioDeptoFinanceiro then s = "true" else s = "false"
	strJS = strJS & _
			"var blnUsuarioDeptoFinanceiro = " & s & ";" & chr(13)

	strJS = strJS & _
			"</script>" & chr(13)





' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________

' ____________________________________________________________________________
' MONTA SQL CHECKBOX RESTRICAO FORMA PAGTO
function monta_sql_restricao_forma_pagto(byval id_orcamentista_e_indicador, byval tipo_cliente)
dim s_sql, s_sql_aux

'	IMPORTANTE: A FORMA DE PAGAMENTO SOMENTE ESTÁ BLOQUEADA SE HOUVER O REGISTRO CUJO CAMPO
'	==========  'st_restricao_ativa' SEJA IGUAL A 1. SE O CAMPO ESTIVER COM O VALOR ZERO OU
'				SE O REGISTRO NÃO EXISTIR, SIGNIFICA QUE A FORMA DE PAGAMENTO ESTÁ LIBERADA.
	s_sql_aux = _
		"SELECT" & _
			" st_restricao_ativa" & _
		" FROM t_ORCAMENTISTA_E_INDICADOR_RESTRICAO_FORMA_PAGTO tOIRFP" & _
		" WHERE" & _
			" (tOIRFP.id_forma_pagto = tFP.id)" & _
			" AND (id_orcamentista_e_indicador = '" & id_orcamentista_e_indicador & "')" & _
			" AND (tipo_cliente = '" & tipo_cliente & "')" & _
			" AND (st_restricao_ativa <> 0)"
	
	s_sql = _
		"SELECT" & _
			" tFP.id AS id_forma_pagto," & _
			" tFP.descricao," & _
			" tFP.ordenacao," & _
			" Coalesce((" & s_sql_aux & "), 0) AS st_restricao_ativa" & _
		" FROM t_FORMA_PAGTO tFP" & _
		" WHERE" & _
			" ((hab_a_vista <> 0) OR (hab_entrada <> 0) OR (hab_prestacao <> 0))" & _
		" ORDER BY" & _
			" tFP.ordenacao"
	monta_sql_restricao_forma_pagto = s_sql
end function
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
<script src="<%=URL_FILE__JQUERY_UI_I18N%>" Language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__AJAX_JS%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JANELACEP_JS%>" language="JavaScript" type="text/javascript"></script>

<script type="text/javascript">
function oOpDadosBancarios(grupo, codigo, st_inativo, descricao) {
	this.grupo = grupo;
	this.codigo = codigo;
	this.st_inativo = st_inativo;
	this.descricao = descricao;
}
</script>

<%
	Response.Write strJS
%>

<script type="text/javascript">
	var TAM_MIN_SENHA = <%=TAM_MIN_SENHA%>;
	var s_ckb_id, s_spn_id;

	$(function() {
		$(".CKB_PF, .CKB_PJ").each(function() {
			s_ckb_id = $(this).attr('id');
			s_spn_id = s_ckb_id.replace("ckb_", "spn_");
			if ($(this).is(':checked')) {
				$("#" + s_spn_id).css('color', 'red');
			}
			else {
				$("#" + s_spn_id).css('color', 'darkgreen');
			}
		});
		
		$(".CKB_PF, .CKB_PJ").change(function() {
			s_ckb_id = $(this).attr('id');
			s_spn_id = s_ckb_id.replace("ckb_", "spn_");
			if ($(this).is(':checked')) {
				$("#" + s_spn_id).css('color', 'red');
			}
			else {
				$("#" + s_spn_id).css('color', 'darkgreen');
			}
		});

		$(".CKB_COM_CAR").each(function () {
			s_ckb_id = $(this).attr('id');
			s_spn_id = s_ckb_id.replace("ckb_", "spn_");
			if ($(this).is(':checked')) {
				$("#" + s_spn_id).css('color', 'darkgreen');
			}
			else {
				$("#" + s_spn_id).css('color', 'dimgray');
			}
		});

		$(".CKB_COM_CAR").change(function () {
			s_ckb_id = $(this).attr('id');
			s_spn_id = s_ckb_id.replace("ckb_", "spn_");
			if ($(this).is(':checked')) {
				$("#" + s_spn_id).css('color', 'darkgreen');
			}
			else {
				$("#" + s_spn_id).css('color', 'dimgray');
			}
		});
	});
</script>

<%	dim strScript
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

function opcaoDadosBancariosEditado() {
	return (converte_numero($('input[name="rb_opcao_dados_bancarios"]:checked').val()) != converte_numero($("#rb_opcao_dados_bancarios_original").val()));
}

function dadosContaBancariaEditado() {
	if ($("#banco").val() != $("#banco_original").val()) return true;
	if ($("#agencia").val() != $("#agencia_original").val()) return true;
	if ($("#agencia_dv").val() != $("#agencia_dv_original").val()) return true;
	if ($("#tipo_operacao").val() != $("#tipo_operacao_original").val()) return true;
	if ($("#conta").val() != $("#conta_original").val()) return true;
	if ($("#conta_dv").val() != $("#conta_dv_original").val()) return true;
	if ($("#tipo_conta option:selected").val() != $("#tipo_conta_original").val()) return true;
	if ($("#favorecido").val() != $("#favorecido_original").val()) return true;
	if (retorna_so_digitos($("#favorecido_cnpjcpf").val()) != retorna_so_digitos($("#favorecido_cnpjcpf_original").val())) return true;
	return false;
}

function dadosChavePixEditado() {
	if (converte_numero($("#c_pix_tipo_chave option:selected").val()) != converte_numero($("#c_pix_tipo_chave_original").val())) return true;
	if ($("#c_pix_chave").val() != $("#c_pix_chave_original").val()) return true;
	if ($("#c_pix_favorecido").val() != $("#c_pix_favorecido_original").val()) return true;
	return false;
}

function dadosComissaoCartaoEditado() {
	if ($("#ckb_comissao_cartao_status").is(':checked')) {
		if (converte_numero($("#ckb_comissao_cartao_status_original").val()) != 1) return true;
	}
	else {
		if (converte_numero($("#ckb_comissao_cartao_status_original").val()) != 0) return true;
	}
	if (retorna_so_digitos($("#c_comissao_cartao_cpf").val()) != retorna_so_digitos($("#c_comissao_cartao_cpf_original").val())) return true;
	if ($("#c_comissao_cartao_titular").val() != $("#c_comissao_cartao_titular_original").val()) return true;
	return false;
}

function dadosComissaoNFSeEditado() {
	if (retorna_so_digitos($("#c_comissao_NFSe_cnpj").val()) != retorna_so_digitos($("#c_comissao_NFSe_cnpj_original").val())) return true;
	if ($("#c_comissao_NFSe_razao_social").val() != $("#c_comissao_NFSe_razao_social_original").val()) return true;
	return false;
}

function RemoveItem( f ) {
var b;
	b=window.confirm('Confirma a exclusão do orçamentista / indicador?');
	if (b){
		f.operacao_selecionada.value=OP_EXCLUI;
		dREMOVE.style.visibility="hidden";
		window.status = "Aguarde ...";
		f.submit();
		}
}

function AtualizaItem( f ) {
var s, s_senha, s_msg, sDescricaoOpDadosBancarios, cont;

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
		}

//  NOME FANTASIA
	if (trim(f.c_nome_fantasia.value) == '') {
	    alert('Preencha o nome fantasia!!');
	    f.c_nome_fantasia.focus();
	    return;
	}

//  CEP
	if (!cep_ok(f.cep.value)) {
		alert('CEP inválido!!');
		f.cep.focus();
		return;
		}
//  UF
	if (!uf_ok(f.uf.value)) {
		alert('UF inválida!!');
		f.uf.focus();
		return;
		}
//  TELEFONE / FAX
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
//  Nº CELULAR
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
	if ((trim(f.banco.value) != "") || (trim(f.agencia.value) != "") || (trim(f.conta.value) != "") || (trim(f.favorecido.value) != "") || (trim(f.favorecido_cnpjcpf.value) != "")) {
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
	}

//  SENHA
	if (f.rb_acesso[0].checked) {
		s_senha = trim(f.senha.value);
		if (s_senha == "") {
			alert('Preencha a senha!!');
			f.senha.focus();
			return;
		}

		if (s_senha != trim(f.senha2.value)) {
			alert('A confirmação da senha não confere!!');
			f.senha2.focus();
			return;
		}

		// Validações realizadas somente p/ inclusão de novo parceiro ou se alterou a senha
		if ((f.operacao_selecionada.value == OP_INCLUI) || (f.rb_acesso_original.value != "1") || (f.senha.value != f.senha_original.value)) {
			if (s_senha.length < TAM_MIN_SENHA) {
				alert('A senha deve possuir no mínimo ' + TAM_MIN_SENHA + ' caracteres!!');
				f.senha.focus();
				return;
			}

			if (!(tem_digito(s_senha) && tem_letra(s_senha))) {
				alert("A senha deve conter no mínimo 1 letra e 1 dígito numérico");
				f.senha.focus();
				return;
			}
		}
	}

//  LOJA
	if (trim(f.loja.value)=='') {
		alert('Selecione a loja!!');
		f.loja.focus();
		return;
		}
//  ATENDIDO POR
	if (trim(f.vendedor.value)=='') {
		alert('Selecione o vendedor!!');
		f.vendedor.focus();
		return;
		}
//  ACESSO AO SISTEMA
	if ((!f.rb_acesso[0].checked)&&(!f.rb_acesso[1].checked)) {
		alert('Indique se o acesso ao sistema será liberado ou não!!');
		return;
		}
//  STATUS
	if ((!f.rb_status[0].checked)&&(!f.rb_status[1].checked)) {
		alert('Indique o status!!');
		return;
		}
//  PERMITE RA
	if ((!f.rb_permite_RA_status[0].checked) && (!f.rb_permite_RA_status[1].checked)) {
		alert('Informe se o RA é permitido ou não!!');
		return;
	}

//  EMAIL
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
	if (trim(f.c_email3.value) != "") {
		if (!email_ok(f.c_email3.value)) {
			alert("Email inválido!!");
			f.c_email3.focus();
			return;
		}
	}

//  FORMA COMO CONHECEU A DIS
	if (trim(f.c_forma_como_conheceu_codigo_original.value) != "") {
		if (trim(f.c_forma_como_conheceu_codigo.value) == "") {
			alert("Selecione a forma como conheceu a DIS!!");
			return;
			}
	}
	
	s = "" + f.c_obs.value;
	if (s.length > MAX_TAM_OBS) {
		alert('Conteúdo de "Observações" excede em ' + (s.length-MAX_TAM_OBS) + ' caracteres o tamanho máximo de ' + MAX_TAM_OBS + '!!');
		f.c_obs.focus();
		return;
	}

	for (cont = 1; cont <= 10; cont++) {
	    if ($("#desc_descricao" + cont).val() != "" || $("#desc_valor" + cont).val() != "") {
	        if ($("#desc_descricao" + cont).val() == "") {
	            alert("O campo 'descrição' do valor correspondente deve ser preenchido!");
	            $("#desc_descricao" + cont).css('background-color', '#FA8072');
	            $("#desc_descricao" + cont).css('border-color', '#000');
	            $("#desc_descricao" + cont).focus();
	            return;
	        }
	        if ($("#desc_valor" + cont).val() == "") {
	            alert("O campo 'valor' da descrição correspondente deve ser preenchido!");
	            $("#desc_valor" + cont).css('background-color', '#FA8072');
	            $("#desc_valor" + cont).css('border-color', '#000');
	            $("#desc_valor" + cont).focus();
	            return;
	        }
	    }
	}

	s = trim("" + f.c_id_magento_b2b.value);
	if (s != "") {
		if (retorna_so_digitos(s) != s.replace(".", "")) {
			alert("ID Magento B2B informado está em formato inválido!");
			f.c_id_magento_b2b.focus();
			return;
		}
	}

	// Valida dados do Pix
	if (($('input[name="rb_opcao_dados_bancarios"]:checked').val() == "<%=COD_ORCAMENTISTA_INDICADOR__OP_DADOS_BANCARIOS__CHAVE_PIX%>") || (converte_numero($('#c_pix_tipo_chave option:selected').val()) != 0) || (trim($('#c_pix_chave').val()) != "") || (trim($('#c_pix_favorecido').val()) != "")) {
		if (converte_numero($('#c_pix_tipo_chave option:selected').val()) == 0) {
			alert("Selecione o tipo de chave Pix!");
			$('#c_pix_tipo_chave').focus();
			return;
		}
		if (trim($('#c_pix_chave').val()) == "") {
			alert("Informe a chave Pix!");
			$('#c_pix_chave').focus();
			return;
		}
		if (trim($('#c_pix_favorecido').val()) == "") {
			alert("Informe o nome do favorecido nos dados do Pix!");
			$('#c_pix_favorecido').focus();
			return;
		}
		// Valida a chave Pix de acordo com o tipo de chave
		if ($('#c_pix_tipo_chave option:selected').val() == "<%=COD_ORCAMENTISTA_INDICADOR__PIX_TIPO_CHAVE__CNPJ_CPF%>") {
			if (!cnpj_cpf_ok($('#c_pix_chave').val())) {
				alert("Chave Pix informada é um CNPJ/CPF inválido!");
				$('#c_pix_chave').focus();
				return;
			}
		}
		else if ($('#c_pix_tipo_chave option:selected').val() == "<%=COD_ORCAMENTISTA_INDICADOR__PIX_TIPO_CHAVE__CELULAR%>") {
			if (retorna_so_digitos($('#c_pix_chave').val()).length < 10) {
				alert("Chave Pix informada é um número de telefone inválido!\nVerifique se o DDD foi informado!");
				$('#c_pix_chave').focus();
				return;
			}
		}
		else if ($('#c_pix_tipo_chave option:selected').val() == "<%=COD_ORCAMENTISTA_INDICADOR__PIX_TIPO_CHAVE__EMAIL%>") {
			if (!email_ok($('#c_pix_chave').val())) {
				alert("Chave Pix informada é um e-mail inválido!");
				$('#c_pix_chave').focus();
				return;
			}
		}
	}

	// Verifica se foi feita a opção pelo tipo de dados bancários (somente se o pagamento da comissão não for via cartão)
	var blnStatusAtivoComissaoCartao = false;
	if (blnUsuarioDeptoFinanceiro) {
		blnStatusAtivoComissaoCartao = $("#ckb_comissao_cartao_status").is(':checked');
	}
	else {
		if ($("#ckb_comissao_cartao_status").val() == "1") blnStatusAtivoComissaoCartao = true;
	}

	if (!blnStatusAtivoComissaoCartao) {
		sDescricaoOpDadosBancarios = "";
		if (converte_numero($('input[name="rb_opcao_dados_bancarios"]:checked').val()) == 0) {
			s_msg = "";
			for (var i = 0; i < vOpDadosBancarios.length; i++) {
				if ((vOpDadosBancarios[i].codigo != "") && (vOpDadosBancarios[i].st_inativo == 0)) {
					if (s_msg != "") s_msg += "\n        ou";
					s_msg += "\n        " + vOpDadosBancarios[i].descricao;
				}
			}
			alert("Selecione uma opção de dados bancários:" + s_msg);
			return;
		}
		else {
			for (var i = 0; i < vOpDadosBancarios.length; i++) {
				if (vOpDadosBancarios[i].codigo == converte_numero($('input[name="rb_opcao_dados_bancarios"]:checked').val())) {
					sDescricaoOpDadosBancarios = vOpDadosBancarios[i].descricao;
					break;
				}
			}
		}

		// Confirmação sobre os dados bancários
		// Obs: a confirmação é solicitada somente quando o pagamento da comissão não for via cartão e (a) se houve edição ou (b) se o parceiro está sendo cadastrado agora
		if (f.operacao_selecionada.value == OP_INCLUI) {
			s_msg = "Confirma que o pagamento da comissão será realizado usando '" + sDescricaoOpDadosBancarios + "'?\nCertifique-se de que os dados estão corretos!";
			if (!confirm(s_msg)) return;
		}
		else {
			if (opcaoDadosBancariosEditado() || dadosContaBancariaEditado() || dadosChavePixEditado() || dadosComissaoCartaoEditado() || dadosComissaoNFSeEditado()) {
				s_msg = "Confirma que o pagamento da comissão será realizado usando '" + sDescricaoOpDadosBancarios + "'?\nCertifique-se de que os dados estão corretos!";
				if (!confirm(s_msg)) return;
			}
		}
	} // if (!blnStatusAtivoComissaoCartao)

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

<script type="text/javascript">

    //Every resize of window
    $(window).resize(function () {
        sizeDivEtiqueta();
    });

    //Every scroll of window
    $(window).scroll(function () {
        sizeDivEtiqueta();
    });

    //Dynamically assign height
    function sizeDivEtiqueta() {
        var newTop = $(window).scrollTop() + "px";
        $("#div_etiqueta").css("top", newTop);
        $("#etiqueta_layout").css("top", newTop);
    }

    function CopiaEndereco() {

        $("#etq_endereco").val($("#endereco").val());
        $("#etq_endereco_numero").val($("#endereco_numero").val());
        $("#etq_endereco_complemento").val($("#endereco_complemento").val());
        $("#etq_bairro").val($("#bairro").val());
        $("#etq_cidade").val($("#cidade").val());
        $("#etq_uf").val($("#uf").val());
        $("#etq_cep").val($("#cep").val());

    }

    function AbreJanelaEtiqueta() {
        if ($("#etq_endereco").val() == "") {
            alert("Preencha o endereço a ser impresso na etiqueta!");
            fCAD.etq_endereco.focus();
            return;
        }
        if ($("#etq_endereco_numero").val() == "") {
            alert("Preencha o número do endereço a ser impresso na etiqueta!");
            fCAD.etq_endereco_numero.focus();
            return;
        }
        if ($("#etq_cidade").val() == "") {
            alert("Preencha a cidade a ser impressa na etiqueta!");
            fCAD.etq_cidade.focus();
            return;
        }
        if ($("#etq_uf").val() == "") {
            alert("Preencha a UF a ser impressa na etiqueta!");
            fCAD.etq_uf.focus();
            return;
        }
        if ($("#etq_ddd_1").val() != "" || $("#etq_tel_1").val() != "") {
            if ($("#etq_ddd_1").val() == "") {
                alert("Preencha o DDD do telefone");
                fCAD.etq_ddd_1.focus();
                return;
            }
            if ($("#etq_tel_1").val() == "") {
                alert("Preencha o telefone");
                fCAD.etq_tel_1.focus();
                return;
            }
        }
        if ($("#etq_ddd_2").val() != "" || $("#etq_tel_2").val() != "") {
            if ($("#etq_ddd_2").val() == "") {
                alert("Preencha o DDD do telefone");
                fCAD.etq_ddd_2.focus();
                return;
            }
            if ($("#etq_tel_2").val() == "") {
                alert("Preencha o telefone");
                fCAD.etq_tel_2.focus();
                return;
            }
        }
        if ($("#c_nome_fantasia").val() != "") {
            $("#etq_nome_fantasia").text($("#c_nome_fantasia").val());
        }
        else {
            alert("É necessário preencher o campo 'Nome Fantasia' que será a identificação do indicador na etiqueta!");
            fCAD.c_nome_fantasia.focus();
            return;
        }

        // torna a etiqueta visível
        $("#div_etiqueta").css('display', 'block');
        $("#etiqueta_layout").css('display', 'block');

        if ($("#etq_endereco_complemento").val() == "") {
            $("#separa_complemento").text("");
        }
        else {
            $("#separa_complemento").text(" - ");
        }
        if ($("#etq_bairro").val() == "") {
            $("#separa_bairro").text("");
        }
        else {
            $("#separa_bairro").text(" - ");
        }
        if ($("#etq_cep").val() == "") {
            $("#separa_cep").text("");
        }
        else {
            $("#separa_cep").text(" - ");
        }
        if ($("#etq_ddd_1").val() == "") {
            $("#spn_label_fone").text("");
            $("#spn_fecha_ddd_1").text("");
        }
        else {
            $("#spn_label_fone").text("Fone: (");
            $("#spn_fecha_ddd_1").text(") ");
        }
        if ($("#etq_ddd_2").val() == "") {
            $("#separa_tel").text("");
            $("#spn_abre_ddd_2").text("");
            $("#spn_fecha_ddd_2").text("");
        }
        else {
            $("#separa_tel").text(" / ");
            $("#spn_abre_ddd_2").text("(");
            $("#spn_fecha_ddd_2").text(") ");
        }
        if ($("#etq_email").val() == "") {
            $("#spn_label_email").text("");
        }
        else {
            $("#spn_label_email").text("Email: ");
        }

        $("#spn_etq_endereco").text($("#etq_endereco").val());
        $("#spn_etq_numero").text($("#etq_endereco_numero").val());
        $("#spn_etq_complemento").text($("#etq_endereco_complemento").val());
        $("#spn_etq_bairro").text($("#etq_bairro").val());
        $("#spn_etq_cidade").text($("#etq_cidade").val());
        $("#spn_etq_uf").text($("#etq_uf").val());
        $("#spn_etq_cep").text($("#etq_cep").val());
        $("#spn_etq_ddd_1").text($("#etq_ddd_1").val());
        $("#spn_etq_tel_1").text($("#etq_tel_1").val());
        $("#spn_etq_ddd_2").text($("#etq_ddd_2").val());
        $("#spn_etq_tel_2").text($("#etq_tel_2").val());
        $("#spn_etq_email").text($("#etq_email").val());

        if ($("#etq_ddd_1").val() == $("#etq_ddd_2").val()) {
            $("#spn_abre_ddd_2").text("");
            $("#spn_fecha_ddd_2").text("");
            $("#spn_etq_ddd_2").text("");
        }
        if ($("#etq_ddd_2").val() != "") {
            if ($("#etq_ddd_1").val() == "") {
                $("#spn_etq_ddd_1").text($("#etq_ddd_2").val());
                $("#spn_etq_tel_1").text($("#etq_tel_2").val());
                $("#spn_fecha_ddd_1").text(") ");
                $("#spn_etq_ddd_2").text("");
                $("#spn_etq_tel_2").text("");
                $("#separa_tel").text("");
                $("#spn_abre_ddd_2").text("");
                $("#spn_fecha_ddd_2").text("");
                $("#spn_label_fone").text("Fone: (");
            }
        }

    }

    function fechaEtiqueta() {
        $("#div_etiqueta").css('display', 'none');
        $("#etiqueta_layout").css('display', 'none');
    }

    function mostraOcultaDadosEtiqueta() {
        if ($("#Etq1").is(':visible')) {
            $("#Etq1").hide();
            $("#Etq2").hide();
            $("#Etq3").hide();
            $("#Etq4").hide();
            $("#Etq5").hide();
            $("#Etq6").hide();
            $("#imgEtiqueta").attr({ src: '../imagem/plus.gif' });
        }
        else {
            $("#Etq1").show();
            $("#Etq2").show();
            $("#Etq3").show();
            $("#Etq4").show();
            $("#Etq5").show();
            $("#Etq6").show();
            $("#imgEtiqueta").attr({ src: '../imagem/minus.gif' });
        }
    }

    function mostraOcultaDescontos() {
        if ($("#tblDesc").is(':visible')) {
            $("#tblDesc").hide();
            $("#imgDescontos").attr({ src: '../imagem/plus.gif' });
        }
        else {
            $("#tblDesc").show();
            $("#imgDescontos").attr({ src: '../imagem/minus.gif' });
        }
    }

</script>
<script type="text/javascript">
    function limpaRegistro(pos) {
        $("#desc_descricao" + pos).val("");
        $("#desc_valor" + pos).val("");
    }

    function calcTotal() {
        var i, total, n;
        total=0;

        for (i = 1; i <= fCAD.desc_valor.length; i++) {
            n = converte_numero($("#desc_valor"+i).val());
            
            if (n == "") {
                n = 0;
                n = parseFloat(n);
            }
            
            total += n;
        }
        $("#spn_total").text("<%=SIMBOLO_MONETARIO%> " + formata_moeda(total));
    }

    function tipoOperacao() {
        if (trim($("#banco").val()) == "104") {
            $("#tipo_operacao").attr('disabled', false);
        }
        else {
            $("#tipo_operacao").val("");
            $("#tipo_operacao").attr('disabled', true);
        }
    }

    function verificaNumero(e) {
        if (e.which != 8 && e.which != 0 && (e.which < 48 || e.which > 57)) {
            return false;
        }
    }
</script>

<script type="text/javascript">
    $(function () {
        fCAD.hdd_endereco.value = fCAD.etq_endereco.value;
        fCAD.hdd_numero.value = fCAD.etq_endereco_numero.value;
        fCAD.hdd_complemento.value = fCAD.etq_endereco_complemento.value;
        fCAD.hdd_bairro.value = fCAD.etq_bairro.value;
        fCAD.hdd_cidade.value = fCAD.etq_cidade.value;
        fCAD.hdd_uf.value = fCAD.etq_uf.value;
        fCAD.hdd_cep.value = fCAD.etq_cep.value;
        fCAD.hdd_email.value = fCAD.etq_email.value;
        fCAD.hdd_ddd_1.value = fCAD.etq_ddd_1.value;
        fCAD.hdd_ddd_2.value = fCAD.etq_ddd_2.value;
        fCAD.hdd_tel_1.value = fCAD.etq_tel_1.value;
        fCAD.hdd_tel_2.value = fCAD.etq_tel_2.value;

        $("#Etq1").hide();
        $("#Etq2").hide();
        $("#Etq3").hide();
        $("#Etq4").hide();
        $("#Etq5").hide();
        $("#Etq6").hide();

        $("#tblDesc").hide();

        $("#div_etiqueta").css('filter', 'alpha(opacity=30)');

        calcTotal();

        tipoOperacao();

        $("#tipo_operacao").keypress(verificaNumero);
    });
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
<link href="<%=URL_FILE__JQUERY_UI_CSS%>" rel="stylesheet" type="text/css">

<style type="text/css">
#loja,#vendedor {
	margin: 4pt 0pt 4pt 10pt;
	vertical-align: top;
	}
#rb_acesso,#rb_status,#rb_permite_RA_status {
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
.TitTipoCli 
{
	font-size: 9pt;
	color:black;
	margin-left:15px;
}
.CKB_PF
{
	margin-left:15px;
}
.CKB_PJ
{
	margin-left:15px;
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
<table cellSpacing="0">
<tr>
	<td align="center"><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()"><img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
</tr>
</table>
</center>
</body>



<% else %>

<%	if operacao_selecionada=OP_INCLUI then
		s = "fCAD.razao_social_nome.focus()"
	else
		s = "focus()"
		end if
%>
<body id="corpoPagina" onload="<%=s%>">

<center>

<div id="div_etiqueta" style="width:100%;height:100%;position:absolute;left:0;top:0;display:none;background-color:#000;opacity:0.3"></div>
    <div id="etiqueta_layout" style="display:none;z-index:100;position:absolute;width:500px;height:150px;background-color:#fff;left:50%;top:50%;margin-left:-250px;margin-top:20%;box-shadow:2px 2px 2px #000;border-radius:8px;">
        <a href="javascript:fechaEtiqueta();" title="Fechar" style="font-size:21pt;font-weight:bolder;color:#555;position:relative;right:-240px;top:-30px;margin:0">&times;</a>
        <h1 id="etq_nome_fantasia" style="font-size:12pt;margin-top:0px;font-weight:bolder;text-transform:uppercase"></h1>
        <span id="spn_etq_endereco"></span><span>&nbsp;nº&nbsp;</span><span id="spn_etq_numero"></span><span id="separa_complemento">&nbsp;-&nbsp;</span><span id="spn_etq_complemento"></span><span id="separa_bairro">&nbsp;-&nbsp;</span><span id="spn_etq_bairro"></span>
        <br /><span id="spn_etq_cidade"></span><span>&nbsp;-&nbsp;</span><span id="spn_etq_uf"></span><span id="separa_cep">&nbsp;-&nbsp;</span><span id="spn_etq_cep"></span>
        <br /><span id="spn_label_fone">Fone:&nbsp;(</span><span id="spn_etq_ddd_1"></span><span id="spn_fecha_ddd_1">)&nbsp;</span><span id="spn_etq_tel_1"></span>
        <span id="separa_tel">&nbsp;/&nbsp;</span><span id="spn_abre_ddd_2">(</span><span id="spn_etq_ddd_2"></span><span id="spn_fecha_ddd_2">)&nbsp;</span><span id="spn_etq_tel_2"></span>
        <br /><span id="spn_label_email">Email:&nbsp;</span><span id="spn_etq_email"></span>
    </div>

    <div id="caixa-confirmacao" title="Deseja realmente sair?">
  <span id="msgEtq" style="display:none">Você fez alterações nos dados para etiqueta. Tem certeza que deseja sair sem salvá-las?</span>
</div>

<!-- #include file = "../global/JanelaBuscaCEP.htm"    -->


<!--  CADASTRO DO ORÇAMENTISTA / INDICADOR -->
<br />
<table width="<%=CStr(WIDTH_MAIN_TABLE)%>" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
<%	if operacao_selecionada=OP_INCLUI then
		s = "Cadastro de Novo Orçamentista / Indicador"
	else
		s = "Consulta/Edição de Orçamentista/Indicador Cadastrado"
		end if
%>
	<td align="center" valign="bottom"><span class="PEDIDO"><%=s%></span></td>
</tr>
</table>
<br>


<!--  CAMPOS DO CADASTRO  -->
<form id="fCAD" name="fCAD" method="post" action="OrcamentistaEIndicadorAtualiza.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="operacao_selecionada" id="operacao_selecionada" value='<%=operacao_selecionada%>'>
<input type="hidden" name="tipo_PJ_PF" id="tipo_PJ_PF" value='<%=tipo_PJ_PF%>'>
<input type="hidden" name="pagina_relatorio_originou_edicao" id="pagina_relatorio_originou_edicao" value='<%=pagina_relatorio_originou_edicao%>'>
<input type="hidden" name="ChecadoStatusBloqueado" id="ChecadoStatusBloqueado" value='<%=Cstr(blnChecadoStatusBloqueado)%>'>
<input type="hidden" name="url_origem" id="url_origem" value='<%=url_origem%>' />
<input type="hidden" name="c_FormFieldValues" id="c_FormFieldValues" value="" />

<% if operacao_selecionada=OP_CONSULTA then %>
<INPUT type="hidden" name="c_forma_como_conheceu_codigo_original" id="c_forma_como_conheceu_codigo_original" value='<%=Trim("" & rs("forma_como_conheceu_codigo"))%>' />
<% else %>
<INPUT type="hidden" name="c_forma_como_conheceu_codigo_original" id="c_forma_como_conheceu_codigo_original" value='' />
<% end if%>

<% if operacao_selecionada <> OP_INCLUI then %>
<% if blnVisivelIdMagentoB2B then %>
<input type="hidden" name="c_id_magento_b2b_original" id="c_id_magento_b2b_original" value="<%=Trim("" & rs("id_magento_b2b"))%>" />
<% else %>
<input type="hidden" name="c_id_magento_b2b" id="c_id_magento_b2b" value="<%=Trim("" & rs("id_magento_b2b"))%>" />
<input type="hidden" name="c_id_magento_b2b_original" id="c_id_magento_b2b_original" value="<%=Trim("" & rs("id_magento_b2b"))%>" />
<% end if %>
<% end if %>

<%	dim s_comissao_cartao_status, s_comissao_cartao_cpf, s_comissao_cartao_titular, s_comissao_NFSe_cnpj, s_comissao_NFSe_razao_social
	dim s_opcao_dados_bancarios, s_banco, s_agencia, s_agencia_dv, s_conta_operacao, s_conta, s_conta_dv, s_tipo_conta, s_favorecido, s_favorecido_cnpj_cpf
	dim s_pix_tipo_chave, s_pix_chave, s_pix_favorecido
	s_comissao_cartao_status = ""
	s_comissao_cartao_cpf = ""
	s_comissao_cartao_titular = ""
	s_comissao_NFSe_cnpj = ""
	s_comissao_NFSe_razao_social = ""
	s_opcao_dados_bancarios = ""
	s_banco = ""
	s_agencia = ""
	s_agencia_dv = ""
	s_conta_operacao = ""
	s_conta = ""
	s_conta_dv = ""
	s_tipo_conta = ""
	s_favorecido = ""
	s_favorecido_cnpj_cpf = ""
	s_pix_tipo_chave = ""
	s_pix_chave = ""
	s_pix_favorecido = ""
	if operacao_selecionada = OP_CONSULTA then
		s_comissao_cartao_status = Trim("" & rs("comissao_cartao_status"))
		s_comissao_cartao_cpf = cnpj_cpf_formata(Trim("" & rs("comissao_cartao_cpf")))
		s_comissao_cartao_titular = Trim("" & rs("comissao_cartao_titular"))
		s_comissao_NFSe_cnpj = cnpj_cpf_formata(Trim("" & rs("comissao_NFSe_cnpj")))
		s_comissao_NFSe_razao_social = Trim("" & rs("comissao_NFSe_razao_social"))
		s_opcao_dados_bancarios = Trim("" & rs("opcao_dados_bancarios"))
		s_banco = Trim("" & rs("banco"))
		s_agencia = Trim("" & rs("agencia"))
		s_agencia_dv = Trim("" & rs("agencia_dv"))
		s_conta_operacao = Trim("" & rs("conta_operacao"))
		s_conta = Trim("" & rs("conta"))
		s_conta_dv = Trim("" & rs("conta_dv"))
		s_tipo_conta = Trim("" & rs("tipo_conta"))
		s_favorecido = Trim("" & rs("favorecido"))
		s_favorecido_cnpj_cpf = cnpj_cpf_formata(Trim("" & rs("favorecido_cnpj_cpf")))
		s_pix_tipo_chave = Trim("" & rs("pix_tipo_chave"))
		s_pix_chave = Trim("" & rs("pix_chave"))
		s_pix_favorecido = Trim("" & rs("pix_favorecido"))
		end if
%>

<input type="hidden" name="ckb_comissao_cartao_status_original" id="ckb_comissao_cartao_status_original" value="<%=s_comissao_cartao_status%>" />
<input type="hidden" name="c_comissao_cartao_cpf_original" id="c_comissao_cartao_cpf_original" value="<%=s_comissao_cartao_cpf%>" />
<input type="hidden" name="c_comissao_cartao_titular_original" id="c_comissao_cartao_titular_original" value="<%=s_comissao_cartao_titular%>" />
<input type="hidden" name="c_comissao_NFSe_cnpj_original" id="c_comissao_NFSe_cnpj_original" value="<%=s_comissao_NFSe_cnpj%>" />
<input type="hidden" name="c_comissao_NFSe_razao_social_original" id="c_comissao_NFSe_razao_social_original" value="<%=s_comissao_NFSe_razao_social%>" />
<input type="hidden" name="rb_opcao_dados_bancarios_original" id="rb_opcao_dados_bancarios_original" value="<%=s_opcao_dados_bancarios%>" />
<input type="hidden" name="banco_original" id="banco_original" value="<%=s_banco%>" />
<input type="hidden" name="agencia_original" id="agencia_original" value="<%=s_agencia%>" />
<input type="hidden" name="agencia_dv_original" id="agencia_dv_original" value="<%=s_agencia_dv%>" />
<input type="hidden" name="tipo_operacao_original" id="tipo_operacao_original" value="<%=s_conta_operacao%>" />
<input type="hidden" name="conta_original" id="conta_original" value="<%=s_conta%>" />
<input type="hidden" name="conta_dv_original" id="conta_dv_original" value="<%=s_conta_dv%>" />
<input type="hidden" name="tipo_conta_original" id="tipo_conta_original" value="<%=s_tipo_conta%>" />
<input type="hidden" name="favorecido_original" id="favorecido_original" value="<%=s_favorecido%>" />
<input type="hidden" name="favorecido_cnpjcpf_original" id="favorecido_cnpjcpf_original" value="<%=s_favorecido_cnpj_cpf%>" />
<input type="hidden" name="c_pix_tipo_chave_original" id="c_pix_tipo_chave_original" value="<%=s_pix_tipo_chave%>" />
<input type="hidden" name="c_pix_chave_original" id="c_pix_chave_original" value="<%=s_pix_chave%>" />
<input type="hidden" name="c_pix_favorecido_original" id="c_pix_favorecido_original" value="<%=s_pix_favorecido%>" />


<!-- ************   NOME/RAZÃO SOCIAL   ************ -->
<table width="<%=CStr(WIDTH_MAIN_TABLE)%>" class="Q" cellspacing="0">
	<tr>
		<td class="MD" width="15%" align="left"><p class="R">APELIDO</p><p class="C"><input id="id_selecionado" name="id_selecionado" class="TA" value="<%=id_selecionado%>" readonly size="18" style="text-align:center; color:#0000ff"></p></td>
<%if tipo_PJ_PF=ID_PJ then s_label = "RAZÃO SOCIAL" else s_label="NOME" %>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("razao_social_nome")) else s=""%>
		<td width="85%" align="left"><p class="R"><%=s_label%></p><p class="C"><input id="razao_social_nome" name="razao_social_nome" class="TA" type="text" maxlength="60" size="60" value="<%=s%>" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCAD.c_responsavel_principal.focus(); filtra_nome_identificador();"></p></td>
	</tr>
</table>

<% if blnVisivelIdMagentoB2B then
		if operacao_selecionada=OP_CONSULTA then s = Trim("" & rs("id_magento_b2b")) else s=""
		if s = "0" then s = ""
%>
<!-- ************  ID DO PARCEIRO NO MAGENTO B2B   ************ -->
<table width="<%=CStr(WIDTH_MAIN_TABLE)%>" class="QS" cellSpacing="0">
	<tr>
		<td align="left"><p class="R">ID MAGENTO B2B</p><p class="C"><input id="c_id_magento_b2b" name="c_id_magento_b2b" class="TA" type="text" maxlength="12" size="60" value="<%=s%>" onkeypress="if (digitou_enter(true)) fCAD.c_nome_fantasia.focus();"></p></td>
	</tr>
</table>
<% end if %>

<!-- ************   RESPONSÁVEL PRINCIPAL   ************ -->
<table width="<%=CStr(WIDTH_MAIN_TABLE)%>" class="QS" cellSpacing="0">
	<tr>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("responsavel_principal")) else s=""%>
		<td align="left"><p class="R">PRINCIPAL</p><p class="C"><input id="c_responsavel_principal" name="c_responsavel_principal" class="TA" type="text" maxlength="60" size="60" value="<%=s%>" onkeypress="if (digitou_enter(true)) fCAD.c_nome_fantasia.focus();"></p></td>
	</tr>
</table>

<!-- ************   NOME FANTASIA   ************ -->
<table width="<%=CStr(WIDTH_MAIN_TABLE)%>" class="QS" cellspacing="0">
	<tr>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("nome_fantasia")) else s=""%>
		<td align="left"><p class="R">NOME FANTASIA</p><p class="C"><input id="c_nome_fantasia" name="c_nome_fantasia" class="TA" type="text" maxlength="60" size="60" value="<%=s%>" onkeypress="if (digitou_enter(true)) fCAD.cnpj_cpf.focus(); filtra_nome_identificador();"></p></td>
	</tr>
</table>

<!-- ************   CNPJ/CPF + IE/RG   ************ -->
<table width="<%=CStr(WIDTH_MAIN_TABLE)%>" class="QS" cellspacing="0">
	<tr>
<%if tipo_PJ_PF=ID_PJ then s_label = "CNPJ" else s_label="CPF" %>
<%if operacao_selecionada=OP_CONSULTA then s=cnpj_cpf_formata(Trim("" & rs("cnpj_cpf"))) else s=""%>
	<td class="MD" width="50%" align="left"><p class="R"><%=s_label%></p><p class="C">
		<input id="cnpj_cpf" name="cnpj_cpf" class="TA" value="<%=s%>" maxlength="18" size="24" 
		<% if tipo_PJ_PF = ID_PJ then %>
			onblur="if (!cnpj_ok(this.value)) {alert('CNPJ inválido!'); this.focus();} else this.value=cnpj_formata(this.value);" onkeypress="if (digitou_enter(true) && cnpj_ok(this.value)) fCAD.ie_rg.focus(); filtra_cnpj();"
		<% else %>
			onblur="if (!cpf_ok(this.value)) {alert('CPF inválido!'); this.focus();} else this.value=cpf_formata(this.value);" onkeypress="if (digitou_enter(true) && cpf_ok(this.value)) fCAD.ie_rg.focus(); filtra_cpf();"
		<% end if %>
		></p></td>
<%if tipo_PJ_PF=ID_PJ then s_label = "IE" else s_label="RG" %>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("ie_rg")) else s=""%>
		<td width="50%" align="left"><p class="R"><%=s_label%></p><p class="C"><input id="ie_rg" name="ie_rg" class="TA" type="text" maxlength="20" size="25" value="<%=s%>" onkeypress="if (digitou_enter(true)) fCAD.endereco.focus(); filtra_nome_identificador();"></p></td>
	</tr>
</table>

<!-- ************   ENDEREÇO   ************ -->
<table width="<%=CStr(WIDTH_MAIN_TABLE)%>" class="QS" cellspacing="0">
	<tr>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("endereco")) else s=""%>
		<td width="100%" align="left"><p class="R">ENDEREÇO</p><p class="C"><input id="endereco" name="endereco" class="TA" value="<%=s%>" maxlength="60" style="width:635px;" onkeypress="if (digitou_enter(true)) fCAD.endereco_numero.focus(); filtra_nome_identificador();"></p></td>
	</tr>
</table>

<!-- ************   Nº/COMPLEMENTO   ************ -->
<table width="<%=CStr(WIDTH_MAIN_TABLE)%>" class="QS" cellspacing="0">
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
<table width="<%=CStr(WIDTH_MAIN_TABLE)%>" class="QS" cellspacing="0">
	<tr>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("bairro")) else s=""%>
		<td width="50%" class="MD" align="left"><p class="R">BAIRRO</p><p class="C"><input id="bairro" name="bairro" class="TA" value="<%=s%>" maxlength="72" style="width:310px;" onkeypress="if (digitou_enter(true)) fCAD.cidade.focus(); filtra_nome_identificador();"></p></td>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("cidade")) else s=""%>
		<td width="50%" align="left"><p class="R">CIDADE</p><p class="C"><input id="cidade" name="cidade" class="TA" value="<%=s%>" maxlength="60" style="width:310px;" onkeypress="if (digitou_enter(true)) fCAD.uf.focus(); filtra_nome_identificador();"></p></td>
	</tr>
</table>

<!-- ************   UF/CEP   ************ -->
<table width="<%=CStr(WIDTH_MAIN_TABLE)%>" class="QS" cellspacing="0">
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
<table width="<%=CStr(WIDTH_MAIN_TABLE)%>" class="QS" cellspacing="0">
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
<table width="<%=CStr(WIDTH_MAIN_TABLE)%>" class="QS" cellspacing="0">
	<tr>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("ddd_cel")) else s=""%>
		<td width="15%" class="MD" align="left" nowrap><p class="R">DDD (CEL)</p><p class="C"><input id="ddd_cel" name="ddd_cel" class="TA" value="<%=s%>" maxlength="2" size="3" onkeypress="if (digitou_enter(true) && ddd_ok(this.value)) fCAD.tel_cel.focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inválido!!');this.focus();}"></p></td>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("tel_cel")) else s=""%>
		<td width="25%" class="MD" align="left"><p class="R">TELEFONE (CEL)</p><p class="C"><input id="tel_cel" name="tel_cel" class="TA" value="<%=telefone_formata(s)%>" maxlength="10" size="11" onkeypress="if (digitou_enter(true) && telefone_ok(this.value)) fCAD.contato.focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('Telefone inválido!!');this.focus();} else this.value=telefone_formata(this.value);"></p></td>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("contato")) else s=""%>
		<td align="left"><p class="R">CONTATO</p><p class="C"><input id="contato" name="contato" class="TA" value="<%=s%>" maxlength="40" size="55" onkeypress="if (digitou_enter(true)) fCAD.banco.focus(); filtra_nome_identificador();"></p></td>
	</tr>
</table>

<!-- ************   DADOS BANCÁRIOS   *******************  -->
<table width="<%=CStr(WIDTH_MAIN_TABLE)%>" class="QS" cellspacing="0">
	<tr>
		<td width="100%" style="padding-bottom:10px;" align="left">
			<p class="R" style="padding-bottom:8px;">DADOS BANCÁRIOS</p>
			<table width="607" class="Q" cellspacing="0" style="margin-left:20px;">
			<tr>
				<td width="100%">
					<table width="100%" border="0">
					<tr>
						<%	sDescricao = consulta_descricao_vetor_t_codigo_descricao(vCodDescrOpDadosBancarios, GRUPO_T_CODIGO_DESCRICAO__ORCAMENTISTA_INDICADOR__OP_DADOS_BANCARIOS, COD_ORCAMENTISTA_INDICADOR__OP_DADOS_BANCARIOS__CONTA_BANCARIA, blnDescricaoCadastrada)
							sDescricao = UCase(sDescricao)
							if Not blnDescricaoCadastrada then sDescricao = "DADOS CONTA BANCÁRIA" %>
						<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("opcao_dados_bancarios")) else s=""%>
						<td colspan="2" align="left">
							<input type="radio" id="rb_opcao_dados_bancarios" name="rb_opcao_dados_bancarios" value="<%=CStr(COD_ORCAMENTISTA_INDICADOR__OP_DADOS_BANCARIOS__CONTA_BANCARIA)%>" class="TA"
								<% if s = CStr(COD_ORCAMENTISTA_INDICADOR__OP_DADOS_BANCARIOS__CONTA_BANCARIA) then Response.Write " checked"%>
								/><span id="spn_op_conta_bancaria_status" class="C" style="cursor:default;" onclick="fCAD.rb_opcao_dados_bancarios[0].click();"><%=sDescricao%></span>
						</td>
					</tr>
					<tr>
						<td style="width:10px;">&nbsp;</td>
						<td width="99%" align="left" style="padding-bottom:8px;padding-right:12px;">
							<!-- ************   BANCO/AGÊNCIA/CONTA   ************ -->
							<table class="Q" style="border-bottom:0pt;" width="100%" cellspacing="0">
								<tr>
							<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("banco")) else s=""%>
									<td width="15%" class="MD" nowrap align="left"><p class="R">BANCO</p><p class="C"><input id="banco" name="banco" class="TA" value="<%=s%>" maxlength="4" size="3" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCAD.agencia.focus(); filtra_numerico();" onblur="this.value=trim(this.value);tipoOperacao();"></p></td>
							<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("agencia")) else s=""%>
									<td width="17%" class="MD" align="left"><p class="R">AGÊNCIA</p><p class="C"><input id="agencia" name="agencia" class="TA" value="<%=s%>" maxlength="8" size="5" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCAD.agencia_dv.focus(); filtra_agencia_bancaria();" onblur="this.value=trim(this.value);"></p></td>
							<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("agencia_dv")) else s=""%>
									<td width="5%" class="MD" align="left"><p class="R">DÍG.</p><p class="C"><input id="agencia_dv" name="agencia_dv" class="TA" value="<%=s%>" maxlength="1" size="1" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCAD.conta.focus();" onblur="this.value=trim(this.value);"></p></td>
							<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("conta_operacao")) else s=""%>
									<td width="15%" class="MD" align="left"><p class="R">TIPO OPERAÇÃO</p><p class="C"><input id="tipo_operacao" name="tipo_operacao" class="TA" value="<%=s%>" maxlength="3" size="4" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCAD.tipo_conta.focus();" onblur="this.value=trim(this.value);"></p></td>
							<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("conta")) else s=""%>
									<td width="17%" class="MD" align="left"><p class="R">CONTA</p><p class="C"><input id="conta" name="conta" class="TA" value="<%=s%>" maxlength="12" size="12" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCAD.conta_dv.focus(); filtra_conta_bancaria();" onblur="this.value=trim(this.value);"></p></td>
							<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("conta_dv")) else s=""%>
									<td width="5%" class="MD" align="left"><p class="R">DÍG.</p><p class="C"><input id="conta_dv" name="conta_dv" class="TA" value="<%=s%>" maxlength="2" size="1" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCAD.favorecido.focus();" onblur="this.value=trim(this.value);"></p></td>
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
							<!-- ************   FAVORECIDO / CPF CNPJ FAVORECIDO   *******************  -->
							<table class="Q" width="100%" cellspacing="0">
								<tr>
							<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("favorecido")) else s=""%>
									<td class="MD" width="70%" align="left"><p class="R">FAVORECIDO</p><p class="C"><input id="favorecido" name="favorecido" class="TA" value="<%=s%>" maxlength="40" size="60" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCAD.favorecido_cnpjcpf.focus(); filtra_nome_identificador();" onblur="this.value=trim(this.value);"></p></td>
							<%if operacao_selecionada=OP_CONSULTA then s=cnpj_cpf_formata(Trim("" & rs("favorecido_cnpj_cpf"))) else s="" %>
									<td width="30%" align="left"><p class="R">CPF/CNPJ DO FAVORECIDO</p><p class="C"><input id="favorecido_cnpjcpf" name="favorecido_cnpjcpf" class="TA" type="text" maxlength="18" size="25" value="<%=s%>"
												onblur="if (retorna_so_digitos(this.value).length==14) { this.value=cnpj_formata(this.value);} else if (retorna_so_digitos(this.value).length==11){ this.value=cpf_formata(this.value);} else if (retorna_so_digitos(this.value).length!=0) alert('Formato de CPF/CNPJ inválido!');"></p></td>
								</tr>
							</table>
						</td>
					</tr>
					</table>
				</td>
			</tr>
			<tr>
				<td width="100%" class="MC">
					<table width="100%" border="0">
					<tr>
						<%	sDescricao = consulta_descricao_vetor_t_codigo_descricao(vCodDescrOpDadosBancarios, GRUPO_T_CODIGO_DESCRICAO__ORCAMENTISTA_INDICADOR__OP_DADOS_BANCARIOS, COD_ORCAMENTISTA_INDICADOR__OP_DADOS_BANCARIOS__CHAVE_PIX, blnDescricaoCadastrada)
							sDescricao = UCase(sDescricao)
							if Not blnDescricaoCadastrada then sDescricao = "CHAVE PIX" %>
						<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("opcao_dados_bancarios")) else s=""%>
						<td colspan="2" align="left">
							<input type="radio" id="rb_opcao_dados_bancarios" name="rb_opcao_dados_bancarios" value="<%=CStr(COD_ORCAMENTISTA_INDICADOR__OP_DADOS_BANCARIOS__CHAVE_PIX)%>" class="TA"
								<% if s = CStr(COD_ORCAMENTISTA_INDICADOR__OP_DADOS_BANCARIOS__CHAVE_PIX) then Response.Write " checked"%>
								/><span id="spn_op_chave_pix_status" class="C" style="cursor:default;" onclick="fCAD.rb_opcao_dados_bancarios[1].click();"><%=sDescricao%></span>
						</td>
					</tr>
					<tr>
						<td style="width:10px;">&nbsp;</td>
						<td width="99%" align="left" style="padding-bottom:8px;padding-right:12px;">
							<!-- ************   TIPO CHAVE/CHAVE/FAVORECIDO   ************ -->
							<table class="Q" width="100%" cellspacing="0">
								<tr>
									<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("pix_tipo_chave")) else s=""%>
									<td align="left" class="MD"><p class="R">TIPO DE CHAVE</p><p class="C">
										<select id="c_pix_tipo_chave" name="c_pix_tipo_chave" style="width:130px;">
										<% =tipo_chave_pix_monta_itens_select(s, True) %>
										</select>
									</td>
									<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("pix_chave")) else s=""%>
									<td align="left">
										<p class="R">CHAVE PIX</p>
										<input id="c_pix_chave" name="c_pix_chave" class="TA" value="<%=s%>"
											maxlength="80" size="60"
											style="text-align:left;"
											onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCAD.c_pix_favorecido.focus();"
											onblur="this.value=trim(this.value);"
											/>
									</td>
								</tr>
								<tr>
									<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("pix_favorecido")) else s=""%>
									<td class="MC" colspan="2" align="left">
										<p class="R">NOME DO FAVORECIDO</p>
										<input type="text" id="c_pix_favorecido" name="c_pix_favorecido" class="TA" value="<%=s%>"
											maxlength="80" size="70"
											onblur="this.value=trim(this.value);"
											onkeypress="filtra_nome_identificador();"
											/>
									</td>
								</tr>
							</table>
						</td>
					</tr>
					</table>
				</td>
			</tr>
			</table>
		</td>
	</tr>
</table>

<!-- ************   DADOS P/ PAGTO COMISSÃO: CARTÃO / NFSe   *******************  -->
<table width="<%=CStr(WIDTH_MAIN_TABLE)%>" class="QS" cellspacing="0">
	<tr>
		<td width="100%" style="padding-bottom:10px;" align="left">
			<p class="R" style="padding-bottom:8px;">PAGAMENTO DA COMISSÃO</p>
			<table width="607" class="Q" cellspacing="0" style="margin-left:20px;">
			<tr>
				<td width="100%">
					<p class="R">CARTÃO</p>
					<table width="100%" border="0">
					<tr>
						<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("comissao_cartao_status")) else s=""%>
						<td colspan="2" align="left">
							<input type="checkbox" id="ckb_comissao_cartao_status" name="ckb_comissao_cartao_status" value="ON" class="TA CKB_COM_CAR"
								<% if Not blnUsuarioDeptoFinanceiro then Response.Write " disabled" %>
								<% if s = "1" then Response.Write " checked"%>
								/><span id="spn_comissao_cartao_status" class="C" style="cursor:default;" onclick="fCAD.ckb_comissao_cartao_status.click();">Pagamento Via Cartão</span>
						</td>
					</tr>
					<tr>
						<td style="width:20px;">&nbsp;</td>
						<td width="95%" style="padding-bottom:8px;padding-right:12px;">
							<table class="Q" width="100%" cellspacing="0">
								<tr>
									<%if operacao_selecionada=OP_CONSULTA then s=cnpj_cpf_formata(Trim("" & rs("comissao_cartao_cpf"))) else s=""%>
									<td>
									<p class="R">CPF</p>
									<input type="text" id="c_comissao_cartao_cpf" name="c_comissao_cartao_cpf" class="TA" value="<%=s%>" maxlength="14" size="18" 
										onblur="if (!cpf_ok(this.value)) {alert('CPF inválido!'); this.focus();} else this.value=cpf_formata(this.value);"
										onkeypress="if (digitou_enter(true) && cpf_ok(this.value)) fCAD.c_comissao_cartao_titular.focus(); filtra_cpf();"
										<% if Not blnUsuarioDeptoFinanceiro then Response.Write " readonly tabindex=-1" %>
										/>
									</td>
								</tr>
								<tr>
									<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("comissao_cartao_titular")) else s=""%>
									<td class="MC">
									<p class="R">TITULAR DO CARTÃO</p>
									<input type="text" id="c_comissao_cartao_titular" name="c_comissao_cartao_titular" class="TA" value="<%=s%>" maxlength="60" size="70"
										onblur="this.value=trim(this.value);"
										onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCAD.c_perc_desagio_RA.focus(); filtra_nome_identificador();"
										<% if Not blnUsuarioDeptoFinanceiro then Response.Write " readonly tabindex=-1" %>
										/>
									</td>
								</tr>
							</table>
						</td>
					</tr>
					</table>
				</td>
			</tr>
			<tr>
				<td width="100%" class="MC">
					<p class="R">EMITENTE NFSe</p>
					<table width="100%" border="0">
					<tr>
						<td style="width:20px;">&nbsp;</td>
						<td width="95%" style="padding-bottom:8px;padding-right:12px;">
							<table class="Q" width="100%" cellspacing="0">
								<tr>
									<%if operacao_selecionada=OP_CONSULTA then s=cnpj_cpf_formata(Trim("" & rs("comissao_NFSe_cnpj"))) else s=""%>
									<td>
									<p class="R">CNPJ</p>
									<input type="text" id="c_comissao_NFSe_cnpj" name="c_comissao_NFSe_cnpj" class="TA" value="<%=s%>" maxlength="18" size="24"
										onblur="if (!cnpj_cpf_ok(this.value)) {alert('CNPJ/CPF inválido!'); this.focus();} else this.value=cnpj_cpf_formata(this.value);"
										onkeypress="if (digitou_enter(true) && cnpj_ok(this.value)) fCAD.c_comissao_NFSe_razao_social.focus(); filtra_cnpj_cpf();"
										<% if Not blnUsuarioDeptoFinanceiro then Response.Write " readonly tabindex=-1" %>
										/>
									</td>
								</tr>
								<tr>
									<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("comissao_NFSe_razao_social")) else s=""%>
									<td class="MC">
									<p class="R">RAZÃO SOCIAL DO EMITENTE</p>
									<input type="text" id="c_comissao_NFSe_razao_social" name="c_comissao_NFSe_razao_social" class="TA" value="<%=s%>" maxlength="60" size="70"
										onblur="this.value=trim(this.value);"
										onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCAD.c_perc_desagio_RA.focus(); filtra_nome_identificador();"
										<% if Not blnUsuarioDeptoFinanceiro then Response.Write " readonly tabindex=-1" %>
										/>
									</td>
								</tr>
							</table>
						</td>
					</tr>
					</table>
				</td>
			</tr>
			</table>
		</td>
	</tr>
</table>

<!-- ************   LOJA (DO ORÇAMENTISTA)   ************ -->
<table width="<%=CStr(WIDTH_MAIN_TABLE)%>" class="QS" cellspacing="0">
	<tr>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("loja")) else s=""%>
		<td align="left"><p class="R">LOJA&nbsp;&nbsp;(ORÇAMENTISTAS)</p><p class="C">
			<select id="loja" name="loja" style="width:490px;">
				<% =loja_do_orcamentista_monta_itens_select(s) %>
			</select>
		</p></td>
	</tr>
</table>

<!-- ************   ATENDIDO PELO VENDEDOR (P/ INDICADORES)   ************ -->
<table width="<%=CStr(WIDTH_MAIN_TABLE)%>" class="QS" cellspacing="0">
	<tr>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("vendedor")) else s=""%>
		<td align="left"><p class="R">ATENDIDO POR&nbsp;&nbsp;(INDICADORES)</p><p class="C">
			<select id="vendedor" name="vendedor" style="width:490px;">
			  <% =vendedor_do_indicador_monta_itens_select(s) %>
			</select>
		</p></td>
	</tr>
</table>

<!-- ************   ACESSO AO SISTEMA/STATUS   ************ -->
<table width="<%=CStr(WIDTH_MAIN_TABLE)%>" class="QS" cellspacing="0">
	<tr>
<%if operacao_selecionada=OP_CONSULTA then s_parametro=Cstr(rs("hab_acesso_sistema")) else s_parametro=""%>
		<td width="25%" class="MD" align="left"><p class="R">ACESSO AO SISTEMA</p><p class="C">
			<input type="radio" id="rb_acesso_liberado" name="rb_acesso" value="1" class="TA"<%if s_parametro = "1" then Response.Write(" checked")%>><span onclick="fCAD.rb_acesso[0].click();" style="cursor:default; color:#006600">Liberado</span>
			<br><input type="radio" id="rb_acesso_bloqueado" name="rb_acesso" value="0" class="TA"<%if (s_parametro<>"1") And (s_parametro<>"") then Response.Write(" checked")%>><span onclick="fCAD.rb_acesso[1].click()" style="cursor:default; color:#ff0000">Bloqueado</span>
			</p>
			<input type="hidden" name="rb_acesso_original" id="rb_acesso_original" value="<%=s_parametro%>" />
		</td>
<%if operacao_selecionada=OP_CONSULTA then s_parametro=Trim("" & rs("status")) else s_parametro=""%>
		<td width="25%" class="MD" align="left"><p class="R">STATUS</p><p class="C">
			<input type="radio" id="rb_status_ativo" name="rb_status" value="A" class="TA"<%if s_parametro = "A" then Response.Write(" checked")%>><span onclick="fCAD.rb_status[0].click();" style="cursor:default; color:#006600">Ativo</span>
			<br><input type="radio" id="rb_status_inativo" name="rb_status" value="I" class="TA"<%if (s_parametro<>"A") And (s_parametro<>"") then Response.Write(" checked")%>><span onclick="fCAD.rb_status[1].click();" style="cursor:default; color:#ff0000">Inativo</span>
			</p></td>
<%if operacao_selecionada=OP_CONSULTA then s_parametro=Trim("" & rs("permite_RA_status")) else s_parametro=""%>
		<td width="25%" class="MD" align="left"><p class="R">PERMITE RA</p><p class="C">
			<input type="radio" id="rb_permite_RA_status_sim" name="rb_permite_RA_status" value="1" class="TA"<%if s_parametro = "1" then Response.Write(" checked")%>><span onclick="fCAD.rb_permite_RA_status[0].click();" style="cursor:default; color:#006600">Sim</span>
			<br><input type="radio" id="rb_permite_RA_status_nao" name="rb_permite_RA_status" value="0" class="TA"<%if (s_parametro<>"1") And (s_parametro<>"") then Response.Write(" checked")%>><span onclick="fCAD.rb_permite_RA_status[1].click();" style="cursor:default; color:#ff0000">Não</span>
			</p></td>
<%if operacao_selecionada=OP_CONSULTA then s_parametro=Trim("" & rs("desempenho_nota")) else s_parametro=""%>
		<td width="25%" align="left" valign="top"><p class="R">AVALIAÇÃO DESEMPENHO</p><p class="C">
			<select id="c_desempenho_nota" name="c_desempenho_nota" style="margin-top:4pt; margin-bottom:4pt;width:45px;">
				<% =desempenho_nota_monta_itens_select(s_parametro) %>
			</select>
		</p></td>
	</tr>
</table>

<!-- ************  SENHA / CONFIRMAÇÃO DA SENHA   ************ -->
<table width="<%=CStr(WIDTH_MAIN_TABLE)%>" class="QS" cellspacing="0">
	<tr>
<%
	senha_descripto= ""
	if operacao_selecionada=OP_CONSULTA then
		s = Trim("" & rs("datastamp"))
		chave = gera_chave(FATOR_BD)
		decodifica_dado s, senha_descripto, chave
		end if
%>       
		<td class="MD" width="50%" align="left"><p class="R">SENHA</p><p class="C"><input id="senha" name="senha" class="TA" type="password" maxlength="15" size="18" value="<%=senha_descripto%>" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCAD.senha2.focus();"></p></td>
		<td width="50%" align="left"><p class="R">SENHA (CONFIRMAÇÃO)</p><p class="C"><input id="senha2" name="senha2" class="TA" type="password" maxlength="15" size="18" value="<%=senha_descripto%>" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCAD.c_perc_desagio_RA.focus();"></p></td>
		<input type="hidden" name="senha_original" id="senha_original" value="<%=senha_descripto%>" />
	</tr>
</table>

<!-- ************   LOGIN BLOQUEADO AUTOMATICAMENTE?   ************ -->
<table width="<%=CStr(WIDTH_MAIN_TABLE)%>" class="QS" cellspacing="0">
	<tr>
<%
	s = "&nbsp;"
	s_color = "black"
	if operacao_selecionada=OP_CONSULTA then
		if rs("StLoginBloqueadoAutomatico") <> 0 then
			s = "Bloqueado em " & formata_data_hora_sem_seg(rs("DataHoraBloqueadoAutomatico")) & " (" & Trim("" & rs("QtdeConsecutivaFalhaLogin")) & " tentativas consecutivas com senha errada)"
			s_color = "red"
			end if
		end if
%>
		<td width="100%" align="left">
		<p class="R">LOGIN BLOQUEADO AUTOMATICAMENTE</p>
		<p class="C" id="pMsgStLoginBloqueadoAutomatico" style="color:<%=s_color%>;"><%=s%>
		<% if operacao_selecionada=OP_CONSULTA then
				if rs("StLoginBloqueadoAutomatico") <> 0 then %>
		<input type="checkbox" id="ckb_desbloquear_bloqueio_automatico" name="ckb_desbloquear_bloqueio_automatico" value="ON" class="TA" style="margin-left:15px;" /><span class="C" onclick="fCAD.ckb_desbloquear_bloqueio_automatico.click();" style="cursor:default;">Desbloquear</span>
		<%		end if
			end if %>
		</p>
		</td>
	</tr>
</table>

<!-- ************   PERCENTUAL DE DESÁGIO DO RA / LIMITE MENSAL COMPRAS   ************ -->
<table width="<%=CStr(WIDTH_MAIN_TABLE)%>" class="QS" cellspacing="0">
	<tr>
<%if operacao_selecionada=OP_CONSULTA then s=formata_perc(rs("perc_desagio_RA")) else s=""%>
		<td width="50%" class="MD" align="left"><p class="R">PERCENTUAL DESÁGIO DO RA&nbsp;&nbsp;(INDICADORES)</p><p class="C">
			<input id="c_perc_desagio_RA" name="c_perc_desagio_RA" class="TA" value="<%=s%>" maxlength="5" 
			style="text-align:right;width:60px;"
			onkeypress="if (digitou_enter(true)) fCAD.c_vl_meta.focus(); filtra_percentual();"
			onblur="this.value=formata_numero(this.value,2); if ((converte_numero(this.value)>100)||(converte_numero(this.value)<0)) {alert('Percentual inválido!!');this.focus();}"><span style="margin-left:2px;">%</span>
		</p></td>

<%if operacao_selecionada=OP_CONSULTA then s=formata_moeda(rs("vl_limite_mensal")) else s=""%>
<input type="hidden" name="c_vl_limite_mensal" id="c_vl_limite_mensal" value="<%=s%>">

<%if operacao_selecionada=OP_CONSULTA then s=formata_moeda(rs("vl_meta")) else s=""%>
		<td width="50%" align="left"><p class="R">VL META&nbsp;&nbsp;(<%=SIMBOLO_MONETARIO%>)</p><p class="C">
			<input id="c_vl_meta" name="c_vl_meta" class="TA" value="<%=s%>" maxlength="18" 
			style="text-align:left;width:180px;"
			onkeypress="if (digitou_enter(true)) fCAD.c_email.focus(); filtra_moeda();"
			onblur="this.value=formata_moeda(this.value);">
		</p></td>
	</tr>
</table>

<!-- ************   E-MAILS   ************ -->
<table width="<%=CStr(WIDTH_MAIN_TABLE)%>" class="QS" cellspacing="0">
	<tr>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("email")) else s=""%>
		<td align="left"><p class="R">E-MAIL (1)</p><p class="C">
			<input id="c_email" name="c_email" class="TA" value="<%=s%>" maxlength="60" 
			style="text-align:left;" size="74"
			onkeypress="if (digitou_enter(true)) fCAD.c_email2.focus(); filtra_email();"
			onblur="this.value=trim(this.value);">
		</p></td>
	</tr>
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
			onkeypress="filtra_email();"
			onblur="this.value=trim(this.value);">
		</p></td>
	</tr>
</table>

<!-- ************   TIPO DE ESTABELECIMENTO   ************ -->
<table width="<%=CStr(WIDTH_MAIN_TABLE)%>" class="QS" cellspacing="0">
	<tr>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("tipo_estabelecimento")) else s=""%>
		<td width="100%" style="padding-bottom:4px;" align="left">
		<p class="R">ESTABELECIMENTO</p>
		<input type="radio" id="rb_estabelecimento_casa" name="rb_estabelecimento" value="<%=COD_PARCEIRO_TIPO_ESTABELECIMENTO__CASA%>" class="TA"<%if s = COD_PARCEIRO_TIPO_ESTABELECIMENTO__CASA then Response.Write(" checked")%>><span id="lbl_estabelecimento" onclick="fCAD.rb_estabelecimento[0].click()" style="cursor:default;" class="C">Casa</span>
		<br><input type="radio" id="rb_estabelecimento_escritorio" name="rb_estabelecimento" value="<%=COD_PARCEIRO_TIPO_ESTABELECIMENTO__ESCRITORIO%>" class="TA"<%if s = COD_PARCEIRO_TIPO_ESTABELECIMENTO__ESCRITORIO then Response.Write(" checked")%>><span id="lbl_estabelecimento" onclick="fCAD.rb_estabelecimento[1].click()" style="cursor:default;" class="C">Escritório</span>
		<br><input type="radio" id="rb_estabelecimento_loja" name="rb_estabelecimento" value="<%=COD_PARCEIRO_TIPO_ESTABELECIMENTO__LOJA%>" class="TA"<%if s = COD_PARCEIRO_TIPO_ESTABELECIMENTO__LOJA then Response.Write(" checked")%>><span id="lbl_estabelecimento" onclick="fCAD.rb_estabelecimento[2].click()" style="cursor:default;" class="C">Loja</span>
		<br><input type="radio" id="rb_estabelecimento_oficina" name="rb_estabelecimento" value="<%=COD_PARCEIRO_TIPO_ESTABELECIMENTO__OFICINA%>" class="TA"<%if s = COD_PARCEIRO_TIPO_ESTABELECIMENTO__OFICINA then Response.Write(" checked")%>><span id="lbl_estabelecimento" onclick="fCAD.rb_estabelecimento[3].click()" style="cursor:default;" class="C">Oficina</span>
		</td>
	</tr>
</table>

<!-- ************   CAPTADOR   ************ -->
<table width="<%=CStr(WIDTH_MAIN_TABLE)%>" class="QS" cellspacing="0">
	<tr>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("captador")) else s=""%>
		<td align="left"><p class="R">CAPTADOR</p><p class="C">
			<select id="c_captador" name="c_captador" style="margin-top:4pt; margin-bottom:4pt;">
				<%=captadores_monta_itens_select(s)%>
			</select>
		</p></td>
	</tr>
</table>

<!-- ************   FORMA COMO CONHECEU A DIS   ************ -->
<table width="<%=CStr(WIDTH_MAIN_TABLE)%>" class="QS" cellspacing="0">
	<tr>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("forma_como_conheceu_codigo")) else s=""%>
		<td align="left"><p class="R">FORMA COMO CONHECEU A DIS</p><p class="C">
			<select id="c_forma_como_conheceu_codigo" name="c_forma_como_conheceu_codigo" style="margin-top:4pt; margin-bottom:4pt;width:490px;">
				<%=codigo_descricao_monta_itens_select(GRUPO_T_CODIGO_DESCRICAO__CAD_ORCAMENTISTA_E_INDICADOR__FORMA_COMO_CONHECEU, s)%>
			</select>
		</p></td>
	</tr>
</table>

<!-- ************   VENDEDORES   **************** -->

<% set rs2 = cn.Execute("SELECT * FROM t_ORCAMENTISTA_E_INDICADOR_CONTATOS WHERE (indicador='" & id_selecionado & "') ORDER BY dt_cadastro DESC") %>
<table width="<%=CStr(WIDTH_MAIN_TABLE)%>" class="QS" cellspacing="0" style="padding-bottom:6px;">
	<tr>
		<td align="left" class="MB" colspan="2"><p class="R">VENDEDORES</p></td>
	</tr>
    <tr>
        <td align="left"><p class="R" style="margin-bottom:3px;margin-top:3px">NOME</p></td>
        <td align="left"><p class="R" style="margin-bottom:3px;margin-top:3px;margin-right:5px">CADASTRO</p></td>
    </tr>

<% i = 1
    do while Not rs2.Eof
    i = i + 1
%>
    <tr>
        <td align="left" width="40%">
            <input id="c_indicador_contato_<%=i%>" name="c_indicador_contato" class="TA" value='<%=Trim("" & rs2("nome"))%>' style="text-align: left;margin-left: 5px;border:1px solid #c0c0c0;" size="40" />
            <input type="hidden" name="contato_id" id="contato_id_<%=i%>" value="<%=rs2("id")%>" />

        </td>
        <td align="left">
            <input id="c_indicador_contato_data_<%=i%>" name="c_indicador_contato_data" class="TA" value='<%=formata_data(Trim("" & rs2("dt_cadastro")))%>' style="text-align: left;margin-left: 5px;" size="20" readonly tabindex=-1 />
        </td>
    </tr>
<% rs2.MoveNext
loop %>
<% for cont = i to CADASTRO_INDICADOR_QTDE_MAX_VENDEDORES %>
    <tr>
        <td align="left" width="40%">
            <input id="c_indicador_contato_<%=i%>" name="c_indicador_contato" class="TA" value="" style="text-align: left;margin-left: 5px;border:1px solid #c0c0c0;" maxlength="60" size="40" />
            <input type="hidden" name="contato_id" id="contato_id_<%=i%>" value="" />

        </td>
        <td align="left">
            <input id="c_indicador_contato_data_<%=i%>" name="c_indicador_contato_data" class="TA" value="" style="text-align: left;margin-left: 5px;" size="20" readonly tabindex=-1 />
        </td>
    </tr>
<% next %>
</table>

<!-- ************   OBS   ************ -->
<table width="<%=CStr(WIDTH_MAIN_TABLE)%>" class="QS" cellspacing="0">
	<tr>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("obs")) else s=""%>
		<td align="left"><p class="R">OBSERVAÇÕES</p><p class="C">
			<textarea name="c_obs" id="c_obs" class="PLLe" rows="<%=Cstr(MAX_LINHAS_OBS_ORCAMENTISTA_INDICADOR)%>" 
				style="width:635px;margin-left:1pt;" onkeypress="limita_tamanho(this,MAX_TAM_OBS);" onblur="this.value=trim(this.value);"
				><%=s%></textarea>
		</p></td>
	</tr>
</table>

<!-- ************   CHECADO / PARCEIRO DESDE   ************ -->
<table width="<%=CStr(WIDTH_MAIN_TABLE)%>" class="QS" cellspacing="0">
	<tr>
<%if operacao_selecionada=OP_CONSULTA then s_parametro=Cstr(rs("checado_status")) else s_parametro=""%>
		<td width="50%" class="MD" align="left" valign="top"><p class="R">CHECADO</p>
			<% if blnChecadoStatusBloqueado then %>
				<%if s_parametro = "1" then %>
					<span class="C" style="color:#006600;">SIM (checado) por <%=Trim("" & rs("checado_usuario")) & " - " & formata_data_hora(rs("checado_data"))%></span>
				<% else %>
					<span class="C" style="color:#ff0000;">NÃO (não-checado)</span>
				<% end if %>
			<% else %>
				<p class="C">
				<input type="radio" id="rb_checado_sim" name="rb_checado" value="1" class="TA"<%if s_parametro = "1" then Response.Write(" checked")%>><span onclick="fCAD.rb_checado[0].click();" style="cursor:default; color:#006600">Checado</span>
				<br><input type="radio" id="rb_checado_nao" name="rb_checado" value="0" class="TA"<%if (s_parametro<>"1") And (s_parametro<>"") then Response.Write(" checked")%>><span onclick="fCAD.rb_checado[1].click()" style="cursor:default; color:#ff0000">Não-checado</span>
				</p>
			<% end if %>
			</td>
		<td width="50%" align="left" valign="top"><p class="R">PARCEIRO DESDE</p>
			<span class="C"><%=formata_data(rs("dt_cadastro"))%></span>
		</td>
	</tr>
</table>

<!-- ************   RESTRIÇÃO FORMA DE PAGAMENTO   ************ -->
<table width="<%=CStr(WIDTH_MAIN_TABLE)%>" class="QS" cellspacing="0">
	<tr>
	<td width="100%" align="left" valign="top"><p class="R">RESTRIÇÕES NA FORMA DE PAGAMENTO</p>
		<table width="100%" cellpadding="0" cellspacing="0">
		<tr>
			<td width="50%" align="left" valign="top">
				<p class="R TitTipoCli">Pessoa Física</p>
			<%	s = monta_sql_restricao_forma_pagto(id_selecionado, ID_PF)
				set t = cn.Execute(s)
				do while Not t.Eof
					s_ckb_id = "ckb_" & ID_PF & "_" & Trim("" & t("id_forma_pagto"))
					s_span_id = "spn_" & ID_PF & "_" & Trim("" & t("id_forma_pagto"))
					s_ckb_value = ID_PF & "_" & Trim("" & t("id_forma_pagto"))
					if CLng(t("st_restricao_ativa")) <> 0 then
						s_checked = " checked"
						s_color = "red"
					else
						s_checked = ""
						s_color="darkgreen"
						end if
				%>
					<p class="C"><input type="checkbox" id="<%=s_ckb_id%>" name="<%=s_ckb_id%>" value="<%=s_ckb_value%>" class="TA CKB_PF"<%=s_checked%>><span id="<%=s_span_id%>" style="cursor:default;color:<%=s_color%>;" onclick="fCAD.<%=s_ckb_id%>.click();"><%=Trim("" & t("descricao"))%></span>&nbsp;</p>
				<%
					t.MoveNext
					loop
				%>
			</td>
			<td width="50%" align="left" valign="top">
				<p class="R TitTipoCli">Pessoa Jurídica</p>
			<%	s = monta_sql_restricao_forma_pagto(id_selecionado, ID_PJ)
				set t = cn.Execute(s)
				s_lista_id_forma_pagto = ""
				do while Not t.Eof
					if s_lista_id_forma_pagto <> "" then s_lista_id_forma_pagto = s_lista_id_forma_pagto & "|"
					s_lista_id_forma_pagto = s_lista_id_forma_pagto & Trim("" & t("id_forma_pagto"))
					s_ckb_id = "ckb_" & ID_PJ & "_" & Trim("" & t("id_forma_pagto"))
					s_span_id = "spn_" & ID_PJ & "_" & Trim("" & t("id_forma_pagto"))
					s_ckb_value = ID_PJ & "_" & Trim("" & t("id_forma_pagto"))
					if CLng(t("st_restricao_ativa")) <> 0 then
						s_checked = " checked"
						s_color = "red"
					else
						s_checked = ""
						s_color="darkgreen"
						end if
				%>
					<p class="C"><input type="checkbox" id="<%=s_ckb_id%>" name="<%=s_ckb_id%>" value="<%=s_ckb_value%>" class="TA CKB_PJ"<%=s_checked%>><span id="<%=s_span_id%>" style="cursor:default;color:<%=s_color%>;" onclick="fCAD.<%=s_ckb_id%>.click();"><%=Trim("" & t("descricao"))%></span>&nbsp;</p>
				<%
					t.MoveNext
					loop
				%>
			<input type="hidden" name="c_lista_id_forma_pagto" id="c_lista_id_forma_pagto" value="<%=s_lista_id_forma_pagto%>" />
			</td>
		</tr>
		</table>
	</td>
	</tr>
</table>


<!-- ************   DADOS PARA ETIQUETA   **************** -->
<br />
<table width="<%=CStr(WIDTH_MAIN_TABLE)%>" class="QS" cellSpacing="0">
	<tr>
		<td align="center" class="MC" style="width: 10px"><a href="javascript:mostraOcultaDadosEtiqueta()" title="Mostrar dados para gerar etiqueta"><img id="imgEtiqueta" src="../imagem/plus.gif" border="0" /></a></td>
		<td align="left" class="MC" valign="middle"><a href="javascript:mostraOcultaDadosEtiqueta()" title="Mostrar dados para gerar etiqueta"><p class="R">DADOS PARA ETIQUETA</p></a></td>
		</tr>
</table>


<table id="Etq1" width="<%=CStr(WIDTH_MAIN_TABLE)%>" class="QS" cellSpacing="0">
	<tr>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("etq_endereco")) else s=""%>
		<td width="100%" align="left"><p class="R">ENDEREÇO</p><p class="C"><input id="etq_endereco" name="etq_endereco" class="TA" value="<%=s%>" maxlength="60" style="width:635px;" onkeypress="if (digitou_enter(true)) fCAD.etq_endereco_numero.focus(); filtra_nome_identificador();"></p></td>
	</tr>
</table>


<table id="Etq2" width="<%=CStr(WIDTH_MAIN_TABLE)%>" class="QS" cellSpacing="0">
	<tr>
	<td class="MD" width="50%" align="left"><p class="R">Nº</p><p class="C">
		<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("etq_endereco_numero")) else s=""%>
		<input id="etq_endereco_numero" name="etq_endereco_numero" class="TA" value="<%=s%>" maxlength="20" style="width:310px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCAD.etq_endereco_complemento.focus(); filtra_nome_identificador();"></p></td>
	<td align="left"><p class="R">COMPLEMENTO</p><p class="C">
		<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("etq_endereco_complemento")) else s=""%>
		<input id="etq_endereco_complemento" name="etq_endereco_complemento" class="TA" value="<%=s%>" maxlength="60" style="width:310px;" onkeypress="if (digitou_enter(true)) fCAD.etq_bairro.focus(); filtra_nome_identificador();"></p></td>
	</tr>
</table>


<table id="Etq3" width="<%=CStr(WIDTH_MAIN_TABLE)%>" class="QS" cellSpacing="0">
	<tr>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("etq_bairro")) else s=""%>
		<td width="50%" class="MD" align="left"><p class="R">BAIRRO</p><p class="C"><input id="etq_bairro" name="etq_bairro" class="TA" value="<%=s%>" maxlength="72" style="width:310px;" onkeypress="if (digitou_enter(true)) fCAD.etq_cidade.focus(); filtra_nome_identificador();"></p></td>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("etq_cidade")) else s=""%>
		<td width="50%" align="left"><p class="R">CIDADE</p><p class="C"><input id="etq_cidade" name="etq_cidade" class="TA" value="<%=s%>" maxlength="60" style="width:310px;" onkeypress="if (digitou_enter(true)) fCAD.etq_uf.focus(); filtra_nome_identificador();"></p></td>
	</tr>
</table>


<table id="Etq4" width="<%=CStr(WIDTH_MAIN_TABLE)%>" class="QS" cellSpacing="0">
	<tr>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("etq_uf")) else s=""%>
		<td class="MD"  width="50%" align="left"><p class="R">UF</p><p class="C"><input id="etq_uf" name="etq_uf" class="TA" value="<%=s%>" maxlength="2" size="3" onkeypress="if (digitou_enter(true) && uf_ok(this.value)) fCAD.etq_cep.focus();" onblur="this.value=trim(this.value); if (!uf_ok(this.value)) {alert('UF inválida!!');this.focus();} else this.value=ucase(this.value);"></p></td>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("etq_cep")) else s=""%>
		<td width="50%" align="left"><p class="R">CEP</p><p class="C"><input id="etq_cep" name="etq_cep" class="TA" value="<%=cep_formata(s)%>" maxlength="9" size="11" onkeypress="if (digitou_enter(true) && cep_ok(this.value)) fCAD.etq_ddd_1.focus(); filtra_cep();" onblur="if (!cep_ok(this.value)) {alert('CEP inválido!!');this.focus();} else this.value=cep_formata(this.value);"></p></td>
		
	</tr>
</table>


<table id="Etq5" width="<%=CStr(WIDTH_MAIN_TABLE)%>" class="QS" cellSpacing="0">
	<tr>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("etq_ddd_1")) else s=""%>
		<td width="15%" class="MD" align="left"><p class="R">DDD</p><p class="C"><input id="etq_ddd_1" name="etq_ddd_1" class="TA" value="<%=s%>" maxlength="4" size="5" onkeypress="if (digitou_enter(true) && ddd_ok(this.value)) fCAD.etq_tel_1.focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inválido!!');this.focus();}"></p></td>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("etq_tel_1")) else s=""%>
		<td width="35%" class="MD" align="left"><p class="R">TELEFONE</p><p class="C"><input id="etq_tel_1" name="etq_tel_1" class="TA" value="<%=telefone_formata(s)%>" maxlength="11" size="12" onkeypress="if (digitou_enter(true) && telefone_ok(this.value)) fCAD.etq_ddd_2.focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('Telefone inválido!!');this.focus();} else this.value=telefone_formata(this.value);"></p></td>
		
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("etq_ddd_2")) else s=""%>
		<td width="15%" class="MD" align="left" nowrap><p class="R">DDD</p><p class="C"><input id="etq_ddd_2" name="etq_ddd_2" class="TA" value="<%=s%>" maxlength="2" size="3" onkeypress="if (digitou_enter(true) && ddd_ok(this.value)) fCAD.etq_tel_2.focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inválido!!');this.focus();}"></p></td>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("etq_tel_2")) else s=""%>
		<td width="35%" align="left"><p class="R">TELEFONE</p><p class="C"><input id="etq_tel_2" name="etq_tel_2" class="TA" value="<%=telefone_formata(s)%>" maxlength="10" size="11" onkeypress="if (digitou_enter(true) && telefone_ok(this.value)) fCAD.etq_email.focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('Telefone inválido!!');this.focus();} else this.value=telefone_formata(this.value);"></p></td>
	</tr>
</table>


<table id="Etq6" width="<%=CStr(WIDTH_MAIN_TABLE)%>" class="QS" cellSpacing="0">
	<tr>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("etq_email")) else s=""%>
        <td width="90%" align="left"><p class="R">E-MAIL</p><p class="C">
			<input id="etq_email" name="etq_email" class="TA" value="<%=s%>" maxlength="60" 
			style="text-align:left;" size="50"
			onkeypress="if (digitou_enter(true)) fCAD.etq_gera.focus(); filtra_email();"
			onblur="this.value=trim(this.value);">
		</p></td>
        <td width="5%" align="center"><a href="javascript:CopiaEndereco()"><img src="../imagem/copia_20x20.png" title="Usar mesmo endereço do cadastro" border="0"></a></td>
        <td width="5%" align="center"><a href="javascript:AbreJanelaEtiqueta()"><img id="etq_gera" src="../imagem/lupa_20x20.png" style="width:20px;height:20px" title="Gerar etiqueta" border="0"></a></td>
	</tr>
</table>

<!-- ************   TABELA DE DESCONTOS   **************** -->

<% set rs2 = cn.Execute("SELECT * FROM t_ORCAMENTISTA_E_INDICADOR_DESCONTO WHERE (apelido='" & id_selecionado & "') ORDER BY ordenacao") %>
<% inc = 1 
   s = ""
   sid="-1"
    %>
<% if (operacao_permitida(OP_CEN_CAD_TAB_DESCONTOS_ORCAMENTISTAS_E_INDICADORES, s_lista_operacoes_permitidas)) then
    s_disabled=""
   else
    s_disabled = " disabled"
   end if %>
<br />
<table width="<%=CStr(WIDTH_MAIN_TABLE)%>" class="QS" cellSpacing="0">
	<tr>
		<td class="MC" align="center" style="width: 10px"><a href="javascript:mostraOcultaDescontos()" title="Mostrar tabela de descontos"><img id="imgDescontos" src="../imagem/plus.gif" border="0" /></a></td>
		<td class="MC" align="left" valign="middle"><a href="javascript:mostraOcultaDescontos()" title="Mostrar tabela de descontos"><p class="R">TABELA DE DESCONTOS</p></a></td>
		</tr>
</table>

<table id="tblDesc" width="<%=CStr(WIDTH_MAIN_TABLE)%>" class="QS" cellSpacing="0">
	<tr>

		<td width="490px" align="left"><p class="R" style="margin-bottom:3px;margin-top:3px">DESCRIÇÃO</p></td>
        <td width="129px" align="right"><p class="R" style="margin-bottom:3px;margin-top:3px;margin-right:5px">VALOR</p></td>
        <td width="20px" align="left">&nbsp;</td>
    </tr>
    
    <% do while Not rs2.Eof %>
    <tr>
        <td>
            <p class="C"><input id="desc_descricao<%=inc%>" name="desc_descricao" class="TA" value="<%=rs2("descricao")%>" maxlength="100" style="width:490px;border:1px solid #c0c0c0" onkeypress="if (digitou_enter(true)) fCAD.desc_valor<%=inc%>.focus(); filtra_nome_identificador(); this.style.backgroundColor='#fff'; this.style.borderColor='#c0c0c0'" <%=s_disabled%>></p>
		</td>
        <td>
            <p class="C"><%=SIMBOLO_MONETARIO%>&nbsp;<input id="desc_valor<%=inc%>" name="desc_valor" class="TA" value="<%=formata_moeda(rs2("valor"))%>" maxlength="10" style="width:97px;border:1px solid #c0c0c0;text-align:right" onkeypress="if (digitou_enter(true)) fCAD.desc_descricao<%=inc+1%>.focus(); filtra_nome_identificador(); this.style.backgroundColor='#fff'; this.style.borderColor='#c0c0c0'" onblur="this.value=formata_moeda(this.value);calcTotal();" <%=s_disabled%>></p>
            <input type="hidden" name="id_desc" id="id_desc_<%=inc%>" value="<%=rs2("id")%>" />
        </td>
        <td>
            <% if (operacao_permitida(OP_CEN_CAD_TAB_DESCONTOS_ORCAMENTISTAS_E_INDICADORES, s_lista_operacoes_permitidas)) then %>
            <img src="../imagem/error_14x14.png" title="Limpar registro" onclick="limpaRegistro(<%=inc%>);calcTotal();" style="cursor:pointer" />
            <% end if %>
        </td>
    </tr>
    <% inc = inc + 1
         rs2.MoveNext
         loop %>

    <% for cont = inc to 10 %>
            <tr>
        <td width="500px" align="left">
            <p class="C"><input id="desc_descricao<%=cont%>" name="desc_descricao" class="TA" value="<%=s%>" maxlength="100" style="width:490px;border:1px solid #c0c0c0" onkeypress="if (digitou_enter(true)) fCAD.desc_valor<%=cont%>.focus(); filtra_nome_identificador(); this.style.backgroundColor='#fff'; this.style.borderColor='#c0c0c0'"<%=s_disabled%>></p>
		</td>
        <td width="149px" align="left">
            <p class="C"><%=SIMBOLO_MONETARIO%>&nbsp;<input id="desc_valor<%=cont%>" name="desc_valor" class="TA" value="<%=s%>" maxlength="10" style="width:97px;border:1px solid #c0c0c0;text-align:right" onkeypress="if (digitou_enter(true)) fCAD.desc_descricao<%=cont+1%>.focus(); filtra_nome_identificador(); this.style.backgroundColor='#fff'; this.style.borderColor='#c0c0c0'" onblur="this.value=formata_moeda(this.value);calcTotal();"<%=s_disabled%>></p>
            <input type="hidden" name="id_desc" id="id_desc_<%=cont%>" value="<%=sid%>" />
        </td>
        <td>
           <img src="../imagem/error_14x14.png" title="Limpar registro" onclick="limpaRegistro(<%=cont%>);calcTotal();" style="cursor:pointer" />
        </td>
	</tr>

    <% next %>

    <tr>
        <td align="right"><span class="C">TOTAL:</span></td>
        <td align="right"><span id="spn_total" class="C"></span></td>
        <td>&nbsp;</td>
    </tr>
 
</table>

<!-- ************   SEPARADOR   ************ -->
<table width="<%=CStr(WIDTH_MAIN_TABLE)%>" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc" align="left">&nbsp;</td></tr>
</table>
<br>
<input type="hidden" id="hdd_endereco" name="hdd_endereco" value="" />
<input type="hidden" id="hdd_numero" name="hdd_numero" value="" />
<input type="hidden" id="hdd_complemento" name="hdd_complemento" value="" />
<input type="hidden" id="hdd_bairro" name="hdd_bairro" value="" />
<input type="hidden" id="hdd_cidade" name="hdd_cidade" value="" />
<input type="hidden" id="hdd_uf" name="hdd_uf" value="" />
<input type="hidden" id="hdd_cep" name="hdd_cep" value="" />
<input type="hidden" id="hdd_email" name="hdd_email" value="" />
<input type="hidden" id="hdd_ddd_1" name="hdd_ddd_1" value="" />
<input type="hidden" id="hdd_ddd_2" name="hdd_ddd_2" value="" />
<input type="hidden" id="hdd_tel_1" name="hdd_tel_1" value="" />
<input type="hidden" id="hdd_tel_2" name="hdd_tel_2" value="" />

<script type="text/javascript">

    function SalvouEtiqueta(f) {

        if (f.etq_endereco.value == f.hdd_endereco.value && f.etq_endereco_numero.value == f.hdd_numero.value && f.etq_endereco_complemento.value == f.hdd_complemento.value && f.etq_bairro.value == f.hdd_bairro.value && f.etq_cidade.value == f.hdd_cidade.value && f.etq_uf.value == f.hdd_uf.value && f.etq_cep.value == f.hdd_cep.value && f.etq_email.value == f.hdd_email.value && f.etq_ddd_1.value == f.hdd_ddd_1.value && f.etq_ddd_2.value == f.hdd_ddd_2.value && f.etq_tel_1.value == f.hdd_tel_1.value && f.etq_tel_2.value == f.hdd_tel_2.value) {

            history.back();
        }
        else {
            $("#msgEtq").css('display', 'block');
            $("#caixa-confirmacao").dialog({
                resizable: false,
                height: 175,
                width: 500,
                scroll: false,

                modal: true,

                buttons: {
                    "Sim": function () {
                        $(this).dialog("close");
                        history.back();
                    },
                    "Não": function () {
                        $(this).dialog("close");
                        $("#Etq1").show();
                        $("#Etq2").show();
                        $("#Etq3").show();
                        $("#Etq4").show();
                        $("#Etq5").show();
                        $("#Etq6").show();
                        $("#imgEtiqueta").attr({ src: '../imagem/minus.gif' });
                        $("#msgEtq").css('display', 'none');
                        return;
                    }
                }
            });
        }
    }
    </script>


<table class="notPrint" width="<%=CStr(WIDTH_MAIN_TABLE)%>" cellSpacing="0">
<tr>
	<td align="left"><a href="javascript:SalvouEtiqueta(fCAD);" title="cancela as alterações no cadastro">
		<img src="../botao/cancelar.gif" width="176" height="55" border="0"></a></td>

	<%
	s = ""
	if operacao_selecionada=OP_CONSULTA then
		s = "<td align='center'><div name='dREMOVE' id='dREMOVE'><a href='javascript:RemoveItem(fCAD)' "
		s =s + "title='remove o orçamentista cadastrado'><img src='../botao/remover.gif' width=176 height=55 border=0></a></div>"
		end if
	%><%=s%>
	<td align="right"><div name="dATUALIZA" id="dATUALIZA">
		<a name="bATUALIZA" id="bATUALIZA" href="javascript:AtualizaItem(fCAD)" title="atualiza o cadastro do orçamentista">
		<img src="../botao/confirmar.gif" width="176" height="55" border="0"></a></div>
	</td>
</tr>
</table>
</form>

</center>
</body>

<% end if %>

</html>


<%
'	FECHA CONEXAO COM O BANCO DE DADOS
	rs.Close
	set rs = nothing

    rs2.Close
	set rs2 = nothing
	
	cn.Close
	set cn = nothing
%>