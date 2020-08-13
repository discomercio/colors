<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/Global.asp"    -->
<%
'     =====================================
'	  C L I E N T E E D I T A . A S P
'     =====================================
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
	dim intCounter
	dim s, s_aux, usuario, loja, cnpj_cpf_selecionado, operacao_selecionada, pagina_retorno
	dim s_readonly, s_disabled, s_style_visibility, s_dest
	usuario = trim(Session("usuario_atual"))
	loja = Trim(Session("loja_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 
	If (loja = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 
	
'	CLIENTE A EDITAR
	operacao_selecionada = trim(request("operacao_selecionada"))
	cnpj_cpf_selecionado = retorna_so_digitos(trim(request("cnpj_cpf_selecionado")))
	
	if (operacao_selecionada<>OP_INCLUI) And (operacao_selecionada<>OP_CONSULTA) then Response.Redirect("aviso.asp?id=" & ERR_OPERACAO_NAO_ESPECIFICADA)
	if (operacao_selecionada=OP_INCLUI) And (cnpj_cpf_selecionado="") then Response.Redirect("aviso.asp?id=" & ERR_CLIENTE_NAO_ESPECIFICADO) 
	
'	ESTÁ DEFINIDA A PÁGINA QUE DEVE SER EXIBIDA APÓS A ATUALIZAÇÃO NO CADASTRO?
	pagina_retorno = trim(request("pagina_retorno"))

'	EDIÇÃO BLOQUEADA?
	dim edicao_bloqueada, blnEdicaoBloqueada
	edicao_bloqueada = ucase(trim(request("edicao_bloqueada")))
	blnEdicaoBloqueada = False
	if edicao_bloqueada = "S" then blnEdicaoBloqueada = True


'	CONECTA COM O BANCO DE DADOS
	dim cn,rs,tRefBancaria,tRefComercial,tRefProfissional
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	dim blnLojaHabilitadaProdCompostoECommerce
	blnLojaHabilitadaProdCompostoECommerce = isLojaHabilitadaProdCompostoECommerce(loja)

	dim blnUsarMemorizacaoCompletaEnderecos
	blnUsarMemorizacaoCompletaEnderecos = isActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos

	dim intIdx
	Dim id_cliente, msg_erro
	if operacao_selecionada=OP_INCLUI then
		if Not gera_nsu(NSU_CADASTRO_CLIENTES, id_cliente, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_GERAR_NSU)
	else
		id_cliente = trim(request("cliente_selecionado"))
		if id_cliente = "" then Response.Redirect("aviso.asp?id=" & ERR_ID_INVALIDO)
		end if

	s="SELECT * FROM t_CLIENTE WHERE (cnpj_cpf='" & cnpj_cpf_selecionado & "') Or (id='" & id_cliente & "')"
	set rs = cn.Execute(s)
	if Err <> 0 then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
		
	if operacao_selecionada=OP_INCLUI then
		if Not rs.EOF then Response.Redirect("aviso.asp?id=" & ERR_CLIENTE_JA_CADASTRADO)
	elseif operacao_selecionada=OP_CONSULTA then
		if rs.EOF then Response.Redirect("ClientePesquisa.asp?cnpj_cpf_selecionado=" & cnpj_cpf_selecionado)
		end if

	dim eh_cpf
	if (operacao_selecionada=OP_INCLUI) then
		s=cnpj_cpf_selecionado
	else
		s=Trim("" & rs("cnpj_cpf"))
		end if
		
	if len(s)=11 then eh_cpf=True else eh_cpf=False

'	REF BANCÁRIA
	dim blnCadRefBancaria
	dim int_MAX_REF_BANCARIA_CLIENTE
	dim strRefBancariaBanco, strRefBancariaAgencia, strRefBancariaConta
	dim strRefBancariaDdd, strRefBancariaTelefone, strRefBancariaContato
'	O cadastro de Referência Bancária será exibido p/ PF e PJ
	blnCadRefBancaria = True
	if eh_cpf then 
		int_MAX_REF_BANCARIA_CLIENTE = MAX_REF_BANCARIA_CLIENTE_PF
	else
		int_MAX_REF_BANCARIA_CLIENTE = MAX_REF_BANCARIA_CLIENTE_PJ
		end if

'	PJ: REF COMERCIAL
	dim blnCadRefComercial
	dim int_MAX_REF_COMERCIAL_CLIENTE
	dim strRefComercialNomeEmpresa, strRefComercialContato, strRefComercialDdd, strRefComercialTelefone
	if (Not eh_cpf) then blnCadRefComercial = True else blnCadRefComercial = False
	int_MAX_REF_COMERCIAL_CLIENTE = MAX_REF_COMERCIAL_CLIENTE_PJ

'	PF: REF PROFISSIONAL
	dim blnCadRefProfissional
	dim int_MAX_REF_PROFISSIONAL_CLIENTE
	dim strRefProfNomeEmpresa, strRefProfCargo, strRefProfDdd, strRefProfTelefone
	dim strRefProfPeriodoRegistro, strRefProfRendimentos, strRefProfCnpj
	if (eh_cpf) then blnCadRefProfissional = True else blnCadRefProfissional = False
	int_MAX_REF_PROFISSIONAL_CLIENTE = MAX_REF_PROFISSIONAL_CLIENTE_PF
	
'	PJ: DADOS DO SÓCIO MAJORITÁRIO
	dim blnCadSocioMaj
	if (Not eh_cpf) then blnCadSocioMaj = True else blnCadSocioMaj = False

%>


<%=DOCTYPE_LEGADO%>

<html>


<head>
	<title><%=TITULO_JANELA_MODULO_ORCAMENTO%></title>
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
var ja_carregou=false;
var conteudo_original;
var fCepPopup;

$(function () {

    var f;
	if ((typeof (fORC) !== "undefined") && (fORC !== null)) {
		f = fORC;

        if (!f.rb_end_entrega[1].checked) {
            Disabled_change(f, true);
        }
	}

	// Trata o problema em que os campos do formulário são limpos após retornar à esta página c/ o history.back() pela 2ª vez quando ocorre erro de consistência
	if (trim(fCAD.c_FormFieldValues.value) != "") {
		stringToForm(fCAD.c_FormFieldValues.value, $('#fCAD'));
	}

	if ((typeof (fORC) !== "undefined") && (fORC !== null)) {
		if (trim(fORC.c_FormFieldValues.value) != "") {
			stringToForm(fORC.c_FormFieldValues.value, $('#fORC'));
		}
    }

    trataProdutorRuralEndEtg_PF(null);
    trocarEndEtgTipoPessoa(null);

});
function Disabled_True(f) {
    Disabled_change(f, true);
}
function Disabled_False(f) {
    Disabled_change(f, false);
}

function Disabled_change(f, value) {

    if(f.EndEtg_nome) f.EndEtg_nome.disabled = value;
    f.EndEtg_endereco.disabled = value;
    f.EndEtg_endereco_numero.disabled = value;
    f.EndEtg_bairro.disabled = value;
    f.EndEtg_cidade.disabled = value;
    f.EndEtg_obs.disabled = value;
    f.EndEtg_uf.disabled = value;
    f.EndEtg_cep.disabled = value;
    f.bPesqCepEndEtgNovo.disabled = value;
    f.EndEtg_endereco_complemento.disabled = value;

    var lista = $(".Habilitar_EndEtg_outroendereco input");
    for (var i = 0; i < lista.length; i++) {
        lista[i].disabled = value;
    }
    trocarEndEtgTipoPessoa(null);
}

function ProcessaSelecaoCEP(){};

function OrcamentoAbrePesquisaCep(){ AbrePesquisaCepComum(TrataCepEnderecoOrcamento); }
function AbrePesquisaCep(){ AbrePesquisaCepComum(TrataCepEnderecoCadastro); }

function AbrePesquisaCepComum(TrataCepEnderecoRotina){
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
	ProcessaSelecaoCEP=TrataCepEnderecoRotina;
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

function TrataCepEnderecoOrcamento(strCep, strUF, strLocalidade, strBairro, strLogradouro, strEnderecoNumero, strEnderecoComplemento) {
var f;
	f=fORC;
	f.orcamento_endereco_cep.value=cep_formata(strCep);
	f.orcamento_endereco_uf.value=strUF;
	f.orcamento_endereco_cidade.value=strLocalidade;
	f.orcamento_endereco_bairro.value=strBairro;
	f.orcamento_endereco_logradouro.value=strLogradouro;
	f.orcamento_endereco_numero.value=strEnderecoNumero;
	f.orcamento_endereco_complemento.value=strEnderecoComplemento;
	f.orcamento_endereco_logradouro.focus();
	window.status="Concluído";
}

function AbrePesquisaCepEndEtg(){
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
	f=fORC;
	ProcessaSelecaoCEP=TrataCepEnderecoEntrega;
	strUrl="../Global/AjaxCepPesqPopup.asp";
	if (trim(f.EndEtg_cep.value)!="") strUrl=strUrl+"?CepDefault="+trim(f.EndEtg_cep.value);
	fCepPopup=window.open(strUrl, "AjaxCepPesqPopup", "status=1,toolbar=0,location=0,menubar=0,directories=0,resizable=1,scrollbars=1,width=980,height=650,left=0,top=0");
	fCepPopup.focus();
}

function TrataCepEnderecoEntrega(strCep, strUF, strLocalidade, strBairro, strLogradouro, strEnderecoNumero, strEnderecoComplemento) {
var f;
	f=fORC;
	f.EndEtg_cep.value=cep_formata(strCep);
	f.EndEtg_uf.value=strUF;
	f.EndEtg_cidade.value=strLocalidade;
	f.EndEtg_bairro.value=strBairro;
	f.EndEtg_endereco.value=strLogradouro;
	f.EndEtg_endereco_numero.value=strEnderecoNumero;
	f.EndEtg_endereco_complemento.value=strEnderecoComplemento;
	f.EndEtg_endereco.focus();
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

function fORCConcluir( f ){
var s;
var eh_cpf;
	if (!ja_carregou) return;
	
	s = retorna_so_digitos(fCAD.cnpj_cpf_selecionado.value);
	eh_cpf = false;
	if (s.length == 11) eh_cpf = true;

	s = retorna_dados_formulario(fCAD);
	if (s!=conteudo_original) {
		if (!confirm("As alterações feitas serão perdidas!!\nContinua mesmo assim?")) return;
		}

    ValidarDadosCadastrais();
    if (!ValidarDadosCadastraisOK)
        return;

	if ((!f.rb_end_entrega[0].checked)&&(!f.rb_end_entrega[1].checked)) {
		alert('Informe se o endereço de entrega será o mesmo endereço do cadastro ou não!!');
		return;
		}

	if (f.rb_end_entrega[1].checked) {
		if (trim(f.EndEtg_endereco.value)=="") {
			alert('Preencha o endereço de entrega!!');
			f.EndEtg_endereco.focus();
			return;
			}

		if (trim(f.EndEtg_endereco_numero.value)=="") {
			alert('Preencha o número do endereço de entrega!!');
			f.EndEtg_endereco_numero.focus();
			return;
			}

		if (trim(f.EndEtg_bairro.value)=="") {
			alert('Preencha o bairro do endereço de entrega!!');
			f.EndEtg_bairro.focus();
			return;
			}

		if (trim(f.EndEtg_cidade.value)=="") {
			alert('Preencha a cidade do endereço de entrega!!');
			f.EndEtg_cidade.focus();
			return;
			}
		if (trim(f.EndEtg_obs.value) == "") {
		    alert('Selecione a justificativa do endereço de entrega!!');
		    f.EndEtg_cep.focus();
		    return;
		}
		s=trim(f.EndEtg_uf.value);
		if ((s=="")||(!uf_ok(s))) {
			alert('UF inválida no endereço de entrega!!');
			f.EndEtg_uf.focus();
			return;
			}
			
		if (trim(f.EndEtg_cep.value)=="") {
			alert('Informe o CEP do endereço de entrega!!');
			f.EndEtg_cep.focus();
			return;
			}
			
		if (!cep_ok(f.EndEtg_cep.value)) {
			alert('CEP inválido no endereço de entrega!!');
			f.EndEtg_cep.focus();
			return;
			}

<%if blnUsarMemorizacaoCompletaEnderecos then%>
<%if Not eh_cpf then%>
            var EndEtg_tipo_pessoa = $('input[name="EndEtg_tipo_pessoa"]:checked').val();
            if (!EndEtg_tipo_pessoa)
                EndEtg_tipo_pessoa = "";
            if (EndEtg_tipo_pessoa != "PJ" && EndEtg_tipo_pessoa != "PF") {
                alert('Necessário escolher Pessoa Jurídica ou Pessoa Física no Endereço de entrega!!');
                f.EndEtg_tipo_pessoa.focus();
                return;
            }

            if (EndEtg_tipo_pessoa == "PJ") {
                //Campos PJ: 

                if (f.EndEtg_cnpj_cpf_PJ.value == "" || !cnpj_ok(f.EndEtg_cnpj_cpf_PJ.value)) {
                    alert('Endereço de entrega: CNPJ inválido!!');
                    f.EndEtg_cnpj_cpf_PJ.focus();
                    return;
                }

                if ($('input[name="EndEtg_contribuinte_icms_status_PJ"]:checked').length == 0) {
                    alert('Endereço de entrega: informe se o cliente é contribuinte do ICMS, não contribuinte ou isento!!');
                    f.EndEtg_contribuinte_icms_status_PJ.focus();
                    return;
                }

                if ((f.EndEtg_contribuinte_icms_status_PJ[1].checked) && (trim(f.EndEtg_ie_PJ.value) == "")) {
                    alert('Endereço de entrega: se o cliente é contribuinte do ICMS a inscrição estadual deve ser preenchida!!');
                    f.EndEtg_ie_PJ.focus();
                    return;
                }
                if ((f.EndEtg_contribuinte_icms_status_PJ[0].checked) && (f.EndEtg_ie_PJ.value.toUpperCase().indexOf('ISEN') >= 0)) {
                    alert('Endereço de entrega: se cliente é não contribuinte do ICMS, não pode ter o valor ISENTO no campo de Inscrição Estadual!!');
                    f.EndEtg_ie_PJ.focus();
                    return;
                }
                if ((f.EndEtg_contribuinte_icms_status_PJ[1].checked) && (f.EndEtg_ie_PJ.value.toUpperCase().indexOf('ISEN') >= 0)) {
                    alert('Endereço de entrega: se cliente é contribuinte do ICMS, não pode ter o valor ISENTO no campo de Inscrição Estadual!!');
                    f.EndEtg_ie_PJ.focus();
                    return;
                }
                if (f.EndEtg_contribuinte_icms_status_PJ[2].checked) {
                    if (f.EndEtg_ie_PJ.value != "") {
                        alert("Endereço de entrega: se o Contribuinte ICMS é isento, o campo IE deve ser vazio!");
                        f.EndEtg_ie_PF.focus();
                        return;
                    }
                }

                if (trim(f.EndEtg_nome.value) == "") {
                    alert('Preencha a razão social no endereço de entrega!!');
                    f.EndEtg_nome.focus();
                    return;
                }

                /*
                telefones PJ:
                EndEtg_ddd_com
                EndEtg_tel_com
                EndEtg_ramal_com
                EndEtg_ddd_com_2
                EndEtg_tel_com_2
                EndEtg_ramal_com_2
    */

                if (!ddd_ok(f.EndEtg_ddd_com.value)) {
                    alert('Endereço de entrega: DDD inválido!!');
                    f.EndEtg_ddd_com.focus();
                    return;
                }
                if (!telefone_ok(f.EndEtg_tel_com.value)) {
                    alert('Endereço de entrega: telefone inválido!!');
                    f.EndEtg_tel_com.focus();
                    return;
                }
                if ((f.EndEtg_ddd_com.value == "") && (f.EndEtg_tel_com.value != "")) {
                    alert('Endereço de entrega: preencha o DDD do telefone.');
                    f.EndEtg_ddd_com.focus();
                    return;
                }
                if ((f.EndEtg_tel_com.value == "") && (f.EndEtg_ddd_com.value != "")) {
                    alert('Endereço de entrega: preencha o telefone.');
                    f.EndEtg_tel_com.focus();
                    return;
                }
                if (trim(f.EndEtg_ddd_com.value) == "" && trim(f.EndEtg_ramal_com.value) != "") {
                    alert('Endereço de entrega: DDD comercial inválido!!');
                    f.EndEtg_ddd_com.focus();
                    return;
                }


                if (!ddd_ok(f.EndEtg_ddd_com_2.value)) {
                    alert('Endereço de entrega: DDD inválido!!');
                    f.EndEtg_ddd_com_2.focus();
                    return;
                }
                if (!telefone_ok(f.EndEtg_tel_com_2.value)) {
                    alert('Endereço de entrega: telefone inválido!!');
                    f.EndEtg_tel_com_2.focus();
                    return;
                }
                if ((f.EndEtg_ddd_com_2.value == "") && (f.EndEtg_tel_com_2.value != "")) {
                    alert('Endereço de entrega: preencha o DDD do telefone.');
                    f.EndEtg_ddd_com_2.focus();
                    return;
                }
                if ((f.EndEtg_tel_com_2.value == "") && (f.EndEtg_ddd_com_2.value != "")) {
                    alert('Endereço de entrega: preencha o telefone.');
                    f.EndEtg_tel_com_2.focus();
                    return;
                }
                if (trim(f.EndEtg_ddd_com_2.value) == "" && trim(f.EndEtg_ramal_com_2.value) != "") {
                    alert('Endereço de entrega: DDD comercial 2 inválido!!');
                    f.EndEtg_ddd_com_2.focus();
                    return;
                }

            }
            else {
                //campos PF

                if (f.EndEtg_cnpj_cpf_PF.value == "" || !cpf_ok(f.EndEtg_cnpj_cpf_PF.value)) {
                    alert('Endereço de entrega: CPF inválido!!');
                    f.EndEtg_cnpj_cpf_PF.focus();
                    return;
                }

                if ((!f.EndEtg_produtor_rural_status_PF[0].checked) && (!f.EndEtg_produtor_rural_status_PF[1].checked)) {
                    alert('Endereço de entrega: informe se o cliente é produtor rural ou não!!');
                    return;
                }
                if (!f.EndEtg_produtor_rural_status_PF[0].checked) {
                    if (!f.EndEtg_contribuinte_icms_status_PF[1].checked) {
                        alert('Endereço de entrega: para ser cadastrado como Produtor Rural, é necessário ser contribuinte do ICMS e possuir nº de IE!!');
                        return;
                    }
                    if ((!f.EndEtg_contribuinte_icms_status_PF[0].checked) && (!f.EndEtg_contribuinte_icms_status_PF[1].checked) && (!f.EndEtg_contribuinte_icms_status_PF[2].checked)) {
                        alert('Endereço de entrega: informe se o cliente é contribuinte do ICMS, não contribuinte ou isento!!');
                        return;
                    }
                    if ((f.EndEtg_contribuinte_icms_status_PF[1].checked) && (trim(f.EndEtg_ie_PF.value) == "")) {
                        alert('Endereço de entrega: se o cliente é contribuinte do ICMS a inscrição estadual deve ser preenchida!!');
                        f.EndEtg_ie_PF.focus();
                        return;
                    }
                    if ((f.EndEtg_contribuinte_icms_status_PF[0].checked) && (f.EndEtg_ie_PF.value.toUpperCase().indexOf('ISEN') >= 0)) {
                        alert('Endereço de entrega: se cliente é não contribuinte do ICMS, não pode ter o valor ISENTO no campo de Inscrição Estadual!!');
                        f.EndEtg_ie_PF.focus();
                        return;
                    }
                    if ((f.EndEtg_contribuinte_icms_status_PF[1].checked) && (f.EndEtg_ie_PF.value.toUpperCase().indexOf('ISEN') >= 0)) {
                        alert('Endereço de entrega: se cliente é contribuinte do ICMS, não pode ter o valor ISENTO no campo de Inscrição Estadual!!');
                        f.EndEtg_ie_PF.focus();
                        return;
                    }

                    if (f.EndEtg_contribuinte_icms_status_PF[2].checked) {
                        if (f.EndEtg_ie_PF.value != "") {
                            alert("Endereço de entrega: se o Contribuinte ICMS é isento, o campo IE deve ser vazio!");
                            f.EndEtg_ie_PF.focus();
                            return;
                        }
                    }
                }
            

                if (trim(f.EndEtg_nome.value) == "") {
                    alert('Preencha o nome no endereço de entrega!!');
                    f.EndEtg_nome.focus();
                    return;
                }

                /*
                telefones PF:
                EndEtg_ddd_res
                EndEtg_tel_res
                EndEtg_ddd_cel
                EndEtg_tel_cel
                */
                if (!ddd_ok(f.EndEtg_ddd_res.value)) {
                    alert('Endereço de entrega: DDD inválido!!');
                    f.EndEtg_ddd_res.focus();
                    return;
                }
                if (!telefone_ok(f.EndEtg_tel_res.value)) {
                    alert('Endereço de entrega: telefone inválido!!');
                    f.EndEtg_tel_res.focus();
                    return;
                }
                if ((trim(f.EndEtg_ddd_res.value) != "") || (trim(f.EndEtg_tel_res.value) != "")) {
                    if (trim(f.EndEtg_ddd_res.value) == "") {
                        alert('Endereço de entrega: preencha o DDD!!');
                        f.EndEtg_ddd_res.focus();
                        return;
                    }
                    if (trim(f.EndEtg_tel_res.value) == "") {
                        alert('Endereço de entrega: preencha o telefone!!');
                        f.EndEtg_tel_res.focus();
                        return;
                    }
                }

                if (!ddd_ok(f.EndEtg_ddd_cel.value)) {
                    alert('Endereço de entrega: DDD inválido!!');
                    f.EndEtg_ddd_cel.focus();
                    return;
                }
                if (!telefone_ok(f.EndEtg_tel_cel.value)) {
                    alert('Endereço de entrega: telefone inválido!!');
                    f.EndEtg_tel_cel.focus();
                    return;
                }
                if ((f.EndEtg_ddd_cel.value == "") && (f.EndEtg_tel_cel.value != "")) {
                    alert('Endereço de entrega: preencha o DDD do celular.');
                    f.EndEtg_tel_cel.focus();
                    return;
                }
                if ((f.EndEtg_tel_cel.value == "") && (f.EndEtg_ddd_cel.value != "")) {
                    alert('Endereço de entrega: preencha o número do celular.');
                    f.EndEtg_tel_cel.focus();
                    return;
                }


            }


<%end if%>
<%end if%>
		}


	//trecho comentado por Luiz para evitar bloqueio de alteração de clientes sem CEP
	//if (trim(fCAD.cep.value)=="") {
	//	alert('É necessário preencher o CEP no cadastro do cliente!!');
	//	return;
	//	}
		
	if (eh_cpf) {
		if ((trim(fCAD.produtor_rural_cadastrado.value) == "0") ||
			((fCAD.rb_produtor_rural[0].checked) && (fCAD.produtor_rural_cadastrado.value != fCAD.rb_produtor_rural[0].value)) ||
			((fCAD.rb_produtor_rural[1].checked) && (fCAD.produtor_rural_cadastrado.value != fCAD.rb_produtor_rural[1].value))) {
			alert('É necessário gravar os dados do cadastro do cliente para que a opção Produtor Rural seja gravada!');
			return;
			}
		if ((!fCAD.rb_produtor_rural[0].checked) && (!fCAD.rb_produtor_rural[1].checked)) {
			alert('Informe se o cliente é produtor rural ou não!!');
			return;
			}
		if (fCAD.rb_produtor_rural[1].checked) {
			if (!fCAD.rb_contribuinte_icms[1].checked) {
				alert('Para ser cadastrado como Produtor Rural, é necessário ser contribuinte do ICMS e possuir nº de IE!!');
				return;
			}
			if ((!fCAD.rb_contribuinte_icms[0].checked) && (!fCAD.rb_contribuinte_icms[1].checked) && (!fCAD.rb_contribuinte_icms[2].checked)) {
				alert('Informe se o cliente é contribuinte do ICMS, não contribuinte ou isento!!');
				return;
				}
			if ((trim(fCAD.contribuinte_icms_cadastrado.value) == "0") ||
				((fCAD.rb_contribuinte_icms[0].checked) && (fCAD.contribuinte_icms_cadastrado.value != fCAD.rb_contribuinte_icms[0].value)) ||
				((fCAD.rb_contribuinte_icms[1].checked) && (fCAD.contribuinte_icms_cadastrado.value != fCAD.rb_contribuinte_icms[1].value)) ||
				((fCAD.rb_contribuinte_icms[2].checked) && (fCAD.contribuinte_icms_cadastrado.value != fCAD.rb_contribuinte_icms[2].value))) {
				alert('É necessário gravar os dados do cadastro do cliente para que a opção Contribuinte ICMS seja gravada!');
				return;
				}
			}
		}
	else {
		if ((!fCAD.rb_contribuinte_icms[0].checked) && (!fCAD.rb_contribuinte_icms[1].checked) && (!fCAD.rb_contribuinte_icms[2].checked)) {
			alert('Informe se o cliente é contribuinte do ICMS, não contribuinte ou isento!!');
			return;
		}
		if ((trim(fCAD.contribuinte_icms_cadastrado.value) == "0") ||
			((fCAD.rb_contribuinte_icms[0].checked) && (fCAD.contribuinte_icms_cadastrado.value != fCAD.rb_contribuinte_icms[0].value)) ||
			((fCAD.rb_contribuinte_icms[1].checked) && (fCAD.contribuinte_icms_cadastrado.value != fCAD.rb_contribuinte_icms[1].value)) ||
			((fCAD.rb_contribuinte_icms[2].checked) && (fCAD.contribuinte_icms_cadastrado.value != fCAD.rb_contribuinte_icms[2].value))) {
			alert('É necessário gravar os dados do cadastro do cliente para que a opção Contribuinte ICMS seja gravada!');
			return;
			}
		}
		// Verifica se o campo IE está vazio quando contribuinte ICMS = isento
		if (eh_cpf) {
			if (!fCAD.rb_produtor_rural[0].checked) {
				if ((fCAD.rb_contribuinte_icms[0].checked) && (fCAD.ie.value.toUpperCase().indexOf('ISEN') >= 0)) {
					alert('Se cliente é não contribuinte do ICMS, não pode ter o valor ISENTO no campo de Inscrição Estadual!!');
					fCAD.ie.focus();
					return;
				}
				if ((fCAD.rb_contribuinte_icms[1].checked) && (fCAD.ie.value.toUpperCase().indexOf('ISEN') >= 0)) {
					alert('Se cliente é contribuinte do ICMS, não pode ter o valor ISENTO no campo de Inscrição Estadual!!');
					fCAD.ie.focus();
					return;
				}
		        if (fCAD.rb_contribuinte_icms[2].checked) {
		            if (fCAD.ie.value != "") {
		                alert("Se o Contribuinte ICMS é isento, o campo IE deve ser vazio!");
		                fCAD.ie.focus();
		                return;
		            }
		        }
		    }
		}
		else {
			if ((fCAD.rb_contribuinte_icms[0].checked) && (fCAD.ie.value.toUpperCase().indexOf('ISEN') >= 0)) {
				alert('Se cliente é não contribuinte do ICMS, não pode ter o valor ISENTO no campo de Inscrição Estadual!!');
				fCAD.ie.focus();
				return;
			}
			if ((fCAD.rb_contribuinte_icms[1].checked) && (fCAD.ie.value.toUpperCase().indexOf('ISEN') >= 0)) {
				alert('Se cliente é contribuinte do ICMS, não pode ter o valor ISENTO no campo de Inscrição Estadual!!');
				fCAD.ie.focus();
				return;
			}
		    if (fCAD.rb_contribuinte_icms[2].checked) {
		        if (fCAD.ie.value != "") {
		            alert("Se o Contribuinte ICMS é isento, o campo IE deve ser vazio!");
		            fCAD.ie.focus();
		            return;
		        }
		    }
		}

    // PARA CLIENTE PJ, É OBRIGATÓRIO O PREENCHIMENTO DO E-MAIL
    if (!eh_cpf) {
        if ((trim(fCAD.email_original.value) == "") && (trim(fCAD.email_xml_original.value) == "")) {
            alert("É obrigatório que o cliente tenha um endereço de e-mail cadastrado!");
            fCAD.email.focus();
            return;
        }
    }

    //campos do endereço de entrega que precisam de transformacao
    transferirCamposEndEtg(fORC);

	fORC.c_FormFieldValues.value = formToString($("#fORC"));

	dORCAMENTO.style.visibility="hidden";
	window.status = "Aguarde ...";
	f.submit(); 
}

function AtualizaCliente( f ) {
var s, eh_cpf, i, blnConsistir, blnConsistirDadosBancarios, blnOk;
var blnCadRefBancaria, blnCadSocioMaj, blnCadRefComercial, blnCadRefProfissional;

	if (!ja_carregou) return;

	s=retorna_so_digitos(f.cnpj_cpf_selecionado.value);
	eh_cpf=false;
	if (s.length==11) eh_cpf=true;
	
	if ((s=="")||(!cnpj_cpf_ok(s))) {
		alert('CNPJ/CPF inválido!!');
		return;
		}
		
	if (eh_cpf) {
		s=trim(f.sexo.value);
		if ((s=="")||(!sexo_ok(s))) {
			alert('Indique qual o sexo!!');
			f.sexo.focus();
			return;
			}
		if (!isDate(f.dt_nasc)) {
			alert('Data inválida!!');
			f.dt_nasc.focus();
			return;
			}
		if ((!f.rb_produtor_rural[0].checked) && (!f.rb_produtor_rural[1].checked)) {
			alert('Informe se o cliente é produtor rural ou não!!');
			return;
			}
		if (!f.rb_produtor_rural[0].checked) {
			if (!fCAD.rb_contribuinte_icms[1].checked) {
				alert('Para ser cadastrado como Produtor Rural, é necessário ser contribuinte do ICMS e possuir nº de IE!!');
				return;
			}
			if ((!f.rb_contribuinte_icms[0].checked) && (!f.rb_contribuinte_icms[1].checked) && (!f.rb_contribuinte_icms[2].checked)) {
				alert('Informe se o cliente é contribuinte do ICMS, não contribuinte ou isento!!');
				return;
				}
			if ((f.rb_contribuinte_icms[1].checked) && (trim(f.ie.value) == "")) {
				alert('Se o cliente é contribuinte do ICMS a inscrição estadual deve ser preenchida!!');
				f.ie.focus();
				return;
			}
			if ((f.rb_contribuinte_icms[0].checked) && (f.ie.value.toUpperCase().indexOf('ISEN') >= 0)) {
				alert('Se cliente é não contribuinte do ICMS, não pode ter o valor ISENTO no campo de Inscrição Estadual!!');
				f.ie.focus();
				return;
			}
			if ((f.rb_contribuinte_icms[1].checked) && (f.ie.value.toUpperCase().indexOf('ISEN') >= 0)) {
				alert('Se cliente é contribuinte do ICMS, não pode ter o valor ISENTO no campo de Inscrição Estadual!!');
				f.ie.focus();
				return;
				}
			}
		}
	else {
		//deixar de exigir preenchimento se cliente não é contribuinte?
		//s=trim(f.ie.value);
		//if (s=="") {
		//	alert('Preencha a Inscrição Estadual!!');
		//	f.ie.focus();
		//	return;
		//	}
		s=trim(f.contato.value);
		if (s=="") {
			alert('Informe o nome da pessoa para contato!!');
			f.contato.focus();
			return;
			}
		if ((!f.rb_contribuinte_icms[0].checked) && (!f.rb_contribuinte_icms[1].checked) && (!f.rb_contribuinte_icms[2].checked)) {
			alert('Informe se o cliente é contribuinte do ICMS, não contribuinte ou isento!!');
			return;
			}
		if ((f.rb_contribuinte_icms[1].checked) && (trim(f.ie.value) == "")) {
			alert('Se o cliente é contribuinte do ICMS a inscrição estadual deve ser preenchida!!');
			f.ie.focus();
			return;
		}
		if ((f.rb_contribuinte_icms[0].checked) && (f.ie.value.toUpperCase().indexOf('ISEN') >= 0)) {
			alert('Se cliente é não contribuinte do ICMS, não pode ter o valor ISENTO no campo de Inscrição Estadual!!');
			f.ie.focus();
			return;
		}
		if ((f.rb_contribuinte_icms[1].checked) && (f.ie.value.toUpperCase().indexOf('ISEN') >= 0)) {
			alert('Se cliente é contribuinte do ICMS, não pode ter o valor ISENTO no campo de Inscrição Estadual!!');
			f.ie.focus();
			return;
			}
		}

		// Verifica se o campo IE está vazio quando contribuinte ICMS = isento
		if (eh_cpf) {
			if (!fCAD.rb_produtor_rural[0].checked) {
				if (fCAD.rb_contribuinte_icms[2].checked) {
					if (fCAD.ie.value != "") {
						alert("Se o Contribuinte ICMS é isento, o campo IE deve ser vazio!");
						fCAD.ie.focus();
						return;
					}
				}
			}
		}
		else {
			if (fCAD.rb_contribuinte_icms[2].checked) {
				if (fCAD.ie.value != "") {
					alert("Se o Contribuinte ICMS é isento, o campo IE deve ser vazio!");
					fCAD.ie.focus();
					return;
				}
			}
		}

	if (trim(f.nome.value)=="") {
		alert('Preencha o nome!!');
		f.nome.focus();
		return;
		}

	if (trim(f.endereco.value)=="") {
		alert('Preencha o endereço!!');
		f.endereco.focus();
		return;
		}

	if (trim(f.endereco_numero.value)=="") {
		alert('Preencha o número do endereço!!');
		f.endereco_numero.focus();
		return;
		}

	if (trim(f.bairro.value)=="") {
		alert('Preencha o bairro!!');
		f.bairro.focus();
		return;
		}

	if (trim(f.cidade.value)=="") {
		alert('Preencha a cidade!!');
		f.cidade.focus();
		return;
		}

	s=trim(f.uf.value);
	if ((s=="")||(!uf_ok(s))) {
		alert('UF inválida!!');
		f.uf.focus();
		return;
		}
		
	if (trim(f.cep.value)=="") {
		alert('Informe o CEP!!');
		return;
		}
		
	if (!cep_ok(f.cep.value)) {
		alert('CEP inválido!!');
		f.cep.focus();
		return;
		}

	if (eh_cpf) {
		if (!ddd_ok(f.ddd_res.value)) {
			alert('DDD inválido!!');
			f._res.focus();
			return;
			}
		if (!telefone_ok(f.tel_res.value)) {
			alert('Telefone inválido!!');
			f.tel_res.focus();
			return;
			}
		if ((trim(f.ddd_res.value)!="")||(trim(f.tel_res.value)!="")) {
			if (trim(f.ddd_res.value)=="") {
				alert('Preencha o DDD!!');
				f.ddd_res.focus();
				return;
				}
			if (trim(f.tel_res.value)=="") {
				alert('Preencha o telefone!!');
				f.tel_res.focus();
				return;
				}
			}
		}
		
	if (!ddd_ok(f.ddd_com.value)) {
		alert('DDD inválido!!');
		f.ddd_com.focus();
		return;
		}

	if (!telefone_ok(f.tel_com.value)) {
		alert('Telefone comercial inválido!!');
		f.tel_com.focus();
		return;
		}

	if ((trim(f.ddd_com.value)!="")||(trim(f.tel_com.value)!="")) {
		if (trim(f.ddd_com.value)=="") {
			alert('Preencha o DDD!!');
			f.ddd_com.focus();
			return;
			}
		if (trim(f.tel_com.value)=="") {
			alert('Preencha o telefone!!');
			f.tel_com.focus();
			return;
			}
		}
    if (trim(f.ddd_com.value) == "" && trim(f.ramal_com.value) != "") {
        alert('DDD comercial inválido!!');
        f.ddd_com.focus();
        return;
    }
	
	if (eh_cpf) {
		if ((trim(f.tel_res.value)=="")&&(trim(f.tel_com.value)=="")&&(trim(f.tel_cel.value)=="")) {
			alert('Preencha pelo menos um telefone!!');
			return;
			}
		}
		else {
		    if (trim(f.tel_com_2.value) == "") {
		        if (trim(f.ddd_com.value) == "") {
		            alert('Preencha o DDD!!');
		            f.ddd_com.focus();
		            return;
		        }
		        if (trim(f.tel_com.value) == "") {
		            alert('Preencha o telefone!!');
		            f.tel_com.focus();
		            return;
		        }
		    }
			if (trim(f.ddd_com_2.value) == "" && trim(f.ramal_com_2.value) != "") {
				alert('DDD comercial 2 inválido!!');
                f.ddd_com_2.focus();
				return;
			}

		}
if (eh_cpf) {
    if (!ddd_ok(f.ddd_cel.value)) {
        alert('DDD inválido!!');
        f.ddd_cel.focus();
        return;
    }
    if (!telefone_ok(f.tel_cel.value)) {
        alert('Telefone inválido!!');
        f.tel_res.focus();
        return;
    }
    if ((f.ddd_cel.value == "") && (f.tel_cel.value != "")) {
        alert('Preencha o DDD do celular.');
        f.ddd_cel.focus();
        return;
    }
    if ((f.tel_cel.value == "") && (f.ddd_cel.value != "")) {
        alert('Preencha o número do celular.');
        f.tel_cel.focus();
        return;
    }
}
if (!eh_cpf) {
    if (!ddd_ok(f.ddd_com_2.value)) {
        alert('DDD inválido!!');
        f.ddd_com_2.focus();
        return;
    }
    if (!telefone_ok(f.tel_com_2.value)) {
        alert('Telefone inválido!!');
        f.tel_com_2.focus();
        return;
    }
    if ((f.ddd_com_2.value == "") && (f.tel_com_2.value != "")) {
        alert('Preencha o DDD do telefone.');
        f.ddd_com_2.focus();
        return;
    }
    if ((f.tel_com_2.value == "") && (f.ddd_com_2.value != "")) {
        alert('Preencha o telefone.');
        f.tel_com_2.focus();
        return;
    }

}
	
	if ( (trim(f.email.value)!="") && (!email_ok(f.email.value)) ) {
		alert('E-mail inválido!!');
		f.email.focus();
		return;
		}

	if ( (trim(f.email_xml.value)!="") && (!email_ok(f.email_xml.value)) ) {
		alert('E-mail (XML) inválido!!');
		f.email_xml.focus();
		return;
	}

    // PARA CLIENTE PJ, É OBRIGATÓRIO O PREENCHIMENTO DO E-MAIL
    if (!eh_cpf) {
        if ((trim(fCAD.email.value) == "") && (trim(fCAD.email_xml.value) == "")) {
            alert("É obrigatório informar um endereço de e-mail");
            fCAD.email.focus();
            return;
        }
    }

/*	if (trim(f.midia.options[f.midia.selectedIndex].value)=="") {
		alert('Indique a forma pela qual conheceu a Bonshop!!');
		return;
		}
*/

//  Ref Bancaria
		//  O cadastro de Referência Bancária será feito p/ PJ
		if (!eh_cpf) {
		    blnCadRefBancaria = true;
		    if (blnCadRefBancaria) {
		        for (i = 1; i < f.c_RefBancariaBanco.length; i++) {
		            blnConsistir = false;
		            if (trim(f.c_RefBancariaBanco[i].value) != "") blnConsistir = true;
		            if (trim(f.c_RefBancariaAgencia[i].value) != "") blnConsistir = true;
		            if (trim(f.c_RefBancariaConta[i].value) != "") blnConsistir = true;
		            if (trim(f.c_RefBancariaDdd[i].value) != "") blnConsistir = true;
		            if (trim(f.c_RefBancariaTelefone[i].value) != "") blnConsistir = true;
		            if (trim(f.c_RefBancariaContato[i].value) != "") blnConsistir = true;
		            if (blnConsistir) {
		                if (trim(f.c_RefBancariaBanco[i].value) == "") {
		                    alert('Informe o banco no cadastro de Referência Bancária!!');
		                    f.c_RefBancariaBanco[i].focus();
		                    return;
		                }
		                if (trim(f.c_RefBancariaAgencia[i].value) == "") {
		                    alert('Informe a agência no cadastro de Referência Bancária!!');
		                    f.c_RefBancariaAgencia[i].focus();
		                    return;
		                }
		                if (trim(f.c_RefBancariaConta[i].value) == "") {
		                    alert('Informe o número da conta no cadastro de Referência Bancária!!');
		                    f.c_RefBancariaConta[i].focus();
		                    return;
		                }
		            }
		        }
		    }
		}

//  Ref Profissional
//  O cadastro de Referência Profissional será feito apenas p/ PF
/*	if (eh_cpf) blnCadRefProfissional=true; else blnCadRefProfissional=false;
	if (blnCadRefProfissional) {
		for (i=1; i<f.c_RefProfNomeEmpresa.length; i++) {
			blnConsistir=false;
			if (trim(f.c_RefProfNomeEmpresa[i].value)!="") blnConsistir=true;
			if (trim(f.c_RefProfCargo[i].value)!="") blnConsistir=true;
			if (trim(f.c_RefProfDdd[i].value)!="") blnConsistir=true;
			if (trim(f.c_RefProfTelefone[i].value)!="") blnConsistir=true;
			if (trim(f.c_RefProfPeriodoRegistro[i].value)!="") blnConsistir=true;
			if (trim(f.c_RefProfRendimentos[i].value)!="") blnConsistir=true;
			if (trim(f.c_RefProfCnpj[i].value)!="") blnConsistir=true;
			if (blnConsistir) {
				if (trim(f.c_RefProfNomeEmpresa[i].value)=="") {
					alert('Informe o nome da empresa no cadastro de Referência Profissional!!');
					f.c_RefProfNomeEmpresa[i].focus();
					return;
					}
				if (trim(f.c_RefProfCargo[i].value)=="") {
					alert('Informe o cargo no cadastro de Referência Profissional!!');
					f.c_RefProfCargo[i].focus();
					return;
					}
				if (trim(f.c_RefProfCnpj[i].value)!="") {
					if (!cnpj_ok(f.c_RefProfCnpj[i].value)) {
						alert('CNPJ inválido!!');
						f.c_RefProfCnpj[i].focus();
						return;
						}
					}
				}
			}
		}
*/

//  Ref Comercial
//  O cadastro de Referência Comercial será feito apenas p/ PJ
	if (!eh_cpf) blnCadRefComercial=true; else blnCadRefComercial=false;
	if (blnCadRefComercial) {
		for (i=1; i<f.c_RefComercialNomeEmpresa.length; i++) {
			blnConsistir=false;
			if (trim(f.c_RefComercialNomeEmpresa[i].value)!="") blnConsistir=true;
			if (trim(f.c_RefComercialContato[i].value)!="") blnConsistir=true;
			if (trim(f.c_RefComercialDdd[i].value)!="") blnConsistir=true;
			if (trim(f.c_RefComercialTelefone[i].value)!="") blnConsistir=true;
			if (blnConsistir) {
				if (trim(f.c_RefComercialNomeEmpresa[i].value)=="") {
					alert('Informe o nome da empresa no cadastro de Referência Comercial!!');
					f.c_RefComercialNomeEmpresa[i].focus();
					return;
					}
				}
			}
		}

//  Dados do Sócio Majoritário
/*	if (!eh_cpf) blnCadSocioMaj=true; else blnCadSocioMaj=false;
	if (blnCadSocioMaj) {
		blnConsistir=false;
		blnConsistirDadosBancarios=false;
		if (trim(f.c_SocioMajNome.value)!="") blnConsistir=true;
		if (trim(f.c_SocioMajCpf.value)!="") blnConsistir=true;
		if (trim(f.c_SocioMajBanco.value)!="") {
			blnConsistir=true;
			blnConsistirDadosBancarios=true;
			}
		if (trim(f.c_SocioMajAgencia.value)!="") {
			blnConsistir=true;
			blnConsistirDadosBancarios=true;
			}
		if (trim(f.c_SocioMajConta.value)!="") {
			blnConsistir=true;
			blnConsistirDadosBancarios=true;
			}
		if (trim(f.c_SocioMajDdd.value)!="") blnConsistir=true;
		if (trim(f.c_SocioMajTelefone.value)!="") blnConsistir=true;
		if (trim(f.c_SocioMajContato.value)!="") blnConsistir=true;
		if (blnConsistir) {
			if (trim(f.c_SocioMajNome.value)=="") {
				alert('Informe o nome do sócio majoritário!!');
				f.c_SocioMajNome.focus();
				return;
				}
			}
		if (blnConsistirDadosBancarios) {
			if (trim(f.c_SocioMajBanco.value)=="") {
				alert('Informe o banco nos dados bancários do sócio majoritário!!');
				f.c_SocioMajBanco.focus();
				return;
				}
			if (trim(f.c_SocioMajAgencia.value)=="") {
				alert('Informe a agência nos dados bancários do sócio majoritário!!');
				f.c_SocioMajAgencia.focus();
				return;
				}
			if (trim(f.c_SocioMajConta.value)=="") {
				alert('Informe o número da conta nos dados bancários do sócio majoritário!!');
				f.c_SocioMajConta.focus();
				return;
				}
			}
		}
*/

	fCAD.c_FormFieldValues.value = formToString($("#fCAD"));

	dATUALIZA.style.visibility="hidden";
	window.status = "Aguarde ...";
	f.submit();
}

//retornamos através de uma variavel global. Fizemos para deixar a estrutura da rotina fica igual às outras.
var ValidarDadosCadastraisOK = false;
function ValidarDadosCadastrais() {
    var eh_cpf, s;
    var f = fORC;
    ValidarDadosCadastraisOK = false;

	s = retorna_so_digitos(fCAD.cnpj_cpf_selecionado.value);
	eh_cpf=false;
	if (s.length==11) eh_cpf=true;
	
	if (!eh_cpf) {
		s=trim(f.orcamento_endereco_contato.value);
		if (s=="") {
			alert('Dados cadastrais: informe o nome da pessoa para contato!!');
			f.orcamento_endereco_contato.focus();
			return;
			}
		if ((!f.orcamento_endereco_contribuinte_icms_status[0].checked) && (!f.orcamento_endereco_contribuinte_icms_status[1].checked) && (!f.orcamento_endereco_contribuinte_icms_status[2].checked)) {
			alert('Dados cadastrais: informe se o cliente é contribuinte do ICMS, não contribuinte ou isento!!');
			return;
			}
		if ((f.orcamento_endereco_contribuinte_icms_status[1].checked) && (trim(f.orcamento_endereco_ie.value) == "")) {
			alert('Dados cadastrais: se o cliente é contribuinte do ICMS a inscrição estadual deve ser preenchida!!');
			f.orcamento_endereco_ie.focus();
			return;
		}
		if ((f.orcamento_endereco_contribuinte_icms_status[0].checked) && (f.orcamento_endereco_ie.value.toUpperCase().indexOf('ISEN') >= 0)) {
			alert('Dados cadastrais: se cliente é não contribuinte do ICMS, não pode ter o valor ISENTO no campo de Inscrição Estadual!!');
			f.orcamento_endereco_ie.focus();
			return;
		}
		if ((f.orcamento_endereco_contribuinte_icms_status[1].checked) && (f.orcamento_endereco_ie.value.toUpperCase().indexOf('ISEN') >= 0)) {
			alert('Dados cadastrais: se cliente é contribuinte do ICMS, não pode ter o valor ISENTO no campo de Inscrição Estadual!!');
			f.orcamento_endereco_ie.focus();
			return;
			}
		}

		// Verifica se o campo IE está vazio quando contribuinte ICMS = isento
		if (!eh_cpf) {
			if (f.orcamento_endereco_contribuinte_icms_status[2].checked) {
				if (f.orcamento_endereco_ie.value != "") {
					alert("Dados cadastrais: se o Contribuinte ICMS é isento, o campo IE deve ser vazio!");
					f.orcamento_endereco_ie.focus();
					return;
				}
			}
		}

	if (!eh_cpf) {
	    if (trim(f.orcamento_endereco_nome.value)=="") {
		    alert('Dados cadastrais: preencha o nome!!');
		    f.orcamento_endereco_nome.focus();
		    return;
		    }
		}

	if (trim(f.orcamento_endereco_logradouro.value)=="") {
		alert('Dados cadastrais: preencha o endereço!!');
		f.orcamento_endereco_logradouro.focus();
		return;
		}

	if (trim(f.orcamento_endereco_numero.value)=="") {
		alert('Dados cadastrais: preencha o número do endereço!!');
		f.orcamento_endereco_numero.focus();
		return;
		}

	if (trim(f.orcamento_endereco_bairro.value)=="") {
		alert('Dados cadastrais: preencha o bairro!!');
		f.orcamento_endereco_bairro.focus();
		return;
		}

	if (trim(f.orcamento_endereco_cidade.value)=="") {
		alert('Dados cadastrais: preencha a cidade!!');
		f.orcamento_endereco_cidade.focus();
		return;
		}

	s=trim(f.orcamento_endereco_uf.value);
	if ((s=="")||(!uf_ok(s))) {
		alert('Dados cadastrais: UF inválida!!');
		f.orcamento_endereco_uf.focus();
		return;
		}
		
	if (trim(f.orcamento_endereco_cep.value)=="") {
		alert('Dados cadastrais: informe o CEP!!');
		return;
		}
		
	if (!cep_ok(f.orcamento_endereco_cep.value)) {
		alert('Dados cadastrais: CEP inválido!!');
		f.orcamento_endereco_cep.focus();
		return;
		}

	if (eh_cpf) {
		if (!ddd_ok(f.orcamento_endereco_ddd_res.value)) {
			alert('Dados cadastrais: DDD inválido!!');
			f.orcamento_endereco_ddd_res.focus();
			return;
			}
		if (!telefone_ok(f.orcamento_endereco_tel_res.value)) {
			alert('Dados cadastrais: telefone inválido!!');
			f.orcamento_endereco_tel_res.focus();
			return;
			}
		if ((trim(f.orcamento_endereco_ddd_res.value)!="")||(trim(f.orcamento_endereco_tel_res.value)!="")) {
			if (trim(f.orcamento_endereco_ddd_res.value)=="") {
				alert('Dados cadastrais: preencha o DDD!!');
				f.orcamento_endereco_ddd_res.focus();
				return;
				}
			if (trim(f.orcamento_endereco_tel_res.value)=="") {
				alert('Dados cadastrais: preencha o telefone!!');
				f.orcamento_endereco_tel_res.focus();
				return;
				}
			}
		}
		
	if (!ddd_ok(f.orcamento_endereco_ddd_com.value)) {
		alert('Dados cadastrais: DDD inválido!!');
		f.orcamento_endereco_ddd_com.focus();
		return;
		}

	if (!telefone_ok(f.orcamento_endereco_tel_com.value)) {
		alert('Dados cadastrais: telefone comercial inválido!!');
		f.orcamento_endereco_tel_com.focus();
		return;
		}

	if ((trim(f.orcamento_endereco_ddd_com.value)!="")||(trim(f.orcamento_endereco_tel_com.value)!="")) {
		if (trim(f.orcamento_endereco_ddd_com.value)=="") {
			alert('Dados cadastrais: preencha o DDD!!');
			f.orcamento_endereco_ddd_com.focus();
			return;
			}
		if (trim(f.orcamento_endereco_tel_com.value)=="") {
			alert('Dados cadastrais: preencha o telefone!!');
			f.orcamento_endereco_tel_com.focus();
			return;
			}
		}
    if (trim(f.orcamento_endereco_ddd_com.value) == "" && trim(f.orcamento_endereco_ramal_com.value) != "") {
        alert('Dados cadastrais: DDD comercial inválido!!');
        f.orcamento_endereco_ddd_com.focus();
        return;
    }
	
	if (eh_cpf) {
		if ((trim(f.orcamento_endereco_tel_res.value)=="")&&(trim(f.orcamento_endereco_tel_com.value)=="")&&(trim(f.orcamento_endereco_tel_cel.value)=="")) {
			alert('Dados cadastrais: preencha pelo menos um telefone!!');
			return;
			}
		}
		else {
		    if (trim(f.orcamento_endereco_tel_com_2.value) == "") {
		        if (trim(f.orcamento_endereco_ddd_com.value) == "") {
		            alert('Dados cadastrais: preencha o DDD!!');
		            f.orcamento_endereco_ddd_com.focus();
		            return;
		        }
		        if (trim(f.orcamento_endereco_tel_com.value) == "") {
		            alert('Dados cadastrais: preencha o telefone!!');
		            f.orcamento_endereco_tel_com.focus();
		            return;
		        }
		    }
			if (trim(f.orcamento_endereco_ddd_com_2.value) == "" && trim(f.orcamento_endereco_ramal_com_2.value) != "") {
				alert('Dados cadastrais: DDD comercial 2 inválido!!');
                f.orcamento_endereco_ddd_com_2.focus();
				return;
			}

}
if (eh_cpf) {
    if (!ddd_ok(f.orcamento_endereco_ddd_cel.value)) {
        alert('Dados cadastrais: DDD inválido!!');
        f.orcamento_endereco_ddd_cel.focus();
        return;
    }
    if (!telefone_ok(f.orcamento_endereco_tel_cel.value)) {
        alert('Dados cadastrais: telefone inválido!!');
        f.orcamento_endereco_tel_cel.focus();
        return;
    }
    if ((f.orcamento_endereco_ddd_cel.value == "") && (f.orcamento_endereco_tel_cel.value != "")) {
        alert('Dados cadastrais: preencha o DDD do celular.');
        f.orcamento_endereco_ddd_cel.focus();
        return;
    }
    if ((f.orcamento_endereco_tel_cel.value == "") && (f.orcamento_endereco_ddd_cel.value != "")) {
        alert('Dados cadastrais: preencha o número do celular.');
        f.orcamento_endereco_tel_cel.focus();
        return;
    }
}
if (!eh_cpf) {
    if (!ddd_ok(f.orcamento_endereco_ddd_com_2.value)) {
        alert('Dados cadastrais: DDD inválido!!');
        f.orcamento_endereco_ddd_com_2.focus();
        return;
    }
    if (!telefone_ok(f.orcamento_endereco_tel_com_2.value)) {
        alert('Dados cadastrais: telefone inválido!!');
        f.orcamento_endereco_tel_com_2.focus();
        return;
    }
    if ((f.orcamento_endereco_ddd_com_2.value == "") && (f.orcamento_endereco_tel_com_2.value != "")) {
        alert('Dados cadastrais: preencha o DDD do telefone.');
        f.orcamento_endereco_ddd_com_2.focus();
        return;
    }
    if ((f.orcamento_endereco_tel_com_2.value == "") && (f.orcamento_endereco_ddd_com_2.value != "")) {
        alert('Dados cadastrais: preencha o telefone.');
        f.orcamento_endereco_tel_com_2.focus();
        return;
    }

}
	
	if ( (trim(f.orcamento_endereco_email.value)!="") && (!email_ok(f.orcamento_endereco_email.value)) ) {
		alert('Dados cadastrais: e-mail inválido!!');
		f.orcamento_endereco_email.focus();
		return;
		}

	if ( (trim(f.orcamento_endereco_email_xml.value)!="") && (!email_ok(f.orcamento_endereco_email_xml.value)) ) {
		alert('Dados cadastrais: e-mail (XML) inválido!!');
		f.orcamento_endereco_email_xml.focus();
		return;
	}

    // PARA CLIENTE PJ, É OBRIGATÓRIO O PREENCHIMENTO DO E-MAIL
    if (!eh_cpf) {
        if ((trim(f.orcamento_endereco_email.value) == "") && (trim(f.orcamento_endereco_email.value) == "")) {
            alert("Dados cadastrais: é obrigatório informar um endereço de e-mail");
            f.orcamento_endereco_email.focus();
            return;
        }
    }

    ValidarDadosCadastraisOK = true;
}


function AtualizaClienteContrib(f) {
	var s, eh_cpf, i;

	if (!ja_carregou) return;

	s = retorna_so_digitos(f.cnpj_cpf_selecionado.value);
	eh_cpf = false;
	if (s.length == 11) eh_cpf = true;

	if (eh_cpf) {
		if ((!f.rb_produtor_rural[0].checked) && (!f.rb_produtor_rural[1].checked)) {
			alert('Informe se o cliente é produtor rural ou não!!');
			return;
		}
		if (!f.rb_produtor_rural[0].checked) {
			if (!fCAD.rb_contribuinte_icms[1].checked) {
				alert('Para ser cadastrado como Produtor Rural, é necessário ser contribuinte do ICMS e possuir nº de IE!!');
				return;
			}
			if ((!f.rb_contribuinte_icms[0].checked) && (!f.rb_contribuinte_icms[1].checked) && (!f.rb_contribuinte_icms[2].checked)) {
				alert('Informe se o cliente é contribuinte do ICMS, não contribuinte ou isento!!');
				return;
			}
			if ((f.rb_contribuinte_icms[1].checked) && (trim(f.ie.value) == "")) {
				alert('Se o cliente é contribuinte do ICMS a inscrição estadual deve ser preenchida!!');
				f.ie.focus();
				return;
			}
			if ((f.rb_contribuinte_icms[0].checked) && (f.ie.value.toUpperCase().indexOf('ISEN') >= 0)) {
				alert('Se cliente é não contribuinte do ICMS, não pode ter o valor ISENTO no campo de Inscrição Estadual!!');
				f.ie.focus();
				return;
			}
			if ((f.rb_contribuinte_icms[1].checked) && (f.ie.value.toUpperCase().indexOf('ISEN') >= 0)) {
				alert('Se cliente é contribuinte do ICMS, não pode ter o valor ISENTO no campo de Inscrição Estadual!!');
				f.ie.focus();
				return;
			}
		}
	}
	else {
		if ((!f.rb_contribuinte_icms[0].checked) && (!f.rb_contribuinte_icms[1].checked) && (!f.rb_contribuinte_icms[2].checked)) {
			alert('Informe se o cliente é contribuinte do ICMS, não contribuinte ou isento!!');
			return;
		}
		if ((f.rb_contribuinte_icms[1].checked) && (trim(f.ie.value) == "")) {
			alert('Se o cliente é contribuinte do ICMS a inscrição estadual deve ser preenchida!!');
			f.ie.focus();
			return;
		}
		if ((f.rb_contribuinte_icms[0].checked) && (f.ie.value.toUpperCase().indexOf('ISEN') >= 0)) {
			alert('Se cliente é não contribuinte do ICMS, não pode ter o valor ISENTO no campo de Inscrição Estadual!!');
			f.ie.focus();
			return;
		}
		if ((f.rb_contribuinte_icms[1].checked) && (f.ie.value.toUpperCase().indexOf('ISEN') >= 0)) {
			alert('Se cliente é contribuinte do ICMS, não pode ter o valor ISENTO no campo de Inscrição Estadual!!');
			f.ie.focus();
			return;
		}
}

// Verifica se o campo IE está vazio quando contribuinte ICMS = isento
if (eh_cpf) {
    if (!fCAD.rb_produtor_rural[0].checked) {
        if (fCAD.rb_contribuinte_icms[2].checked) {
            if (fCAD.ie.value != "") {
                alert("Se o Contribuinte ICMS é isento, o campo IE deve ser vazio!");
                fCAD.ie.focus();
                return;
            }
        }
    }
}
else {
    if (fCAD.rb_contribuinte_icms[2].checked) {
        if (fCAD.ie.value != "") {
            alert("Se o Contribuinte ICMS é isento, o campo IE deve ser vazio!");
            fCAD.ie.focus();
            return;
        }
    }
}

    // PARA CLIENTE PJ, É OBRIGATÓRIO O PREENCHIMENTO DO E-MAIL
    if (!eh_cpf) {
        if ((trim(fCAD.email.value) == "") && (trim(fCAD.email_xml.value) == "")) {
            alert("É obrigatório informar um endereço de e-mail");
            fCAD.email.focus();
            return;
        }
    }

	fCAD.c_FormFieldValues.value = formToString($("#fCAD"));

	dATUALIZACONTRIB.style.visibility = "hidden";
	window.status = "Aguarde ...";
	f.submit();
}

function transferirCamposEndEtg(formulario) {
<%if blnUsarMemorizacaoCompletaEnderecos then %>
<%if Not eh_cpf then %>
    //Transferimos os dados do endereço de entrega dos campos certos. 
    //Temos dois conjuntos de campos (para PF e PJ) porque o layout é muito diferente.
    var pj = $('input[name="EndEtg_tipo_pessoa"]:checked').val() == "PJ";
    if (pj) {
        formulario.EndEtg_cnpj_cpf.value = formulario.EndEtg_cnpj_cpf_PJ.value;
        formulario.EndEtg_ie.value = formulario.EndEtg_ie_PJ.value;
        formulario.EndEtg_contribuinte_icms_status.value = $('input[name="EndEtg_contribuinte_icms_status_PJ"]:checked').val();
        if (!$('input[name="EndEtg_contribuinte_icms_status_PJ"]:checked').val())
            formulario.EndEtg_contribuinte_icms_status.value = "";
    }
    else {
        formulario.EndEtg_cnpj_cpf.value = formulario.EndEtg_cnpj_cpf_PF.value;
        formulario.EndEtg_ie.value = formulario.EndEtg_ie_PF.value;
        formulario.EndEtg_contribuinte_icms_status.value = $('input[name="EndEtg_contribuinte_icms_status_PF"]:checked').val();
        if (!$('input[name="EndEtg_contribuinte_icms_status_PF"]:checked').val())
            formulario.EndEtg_contribuinte_icms_status.value = "";
        formulario.EndEtg_produtor_rural_status.value = $('input[name="EndEtg_produtor_rural_status_PF"]:checked').val();
        if (!$('input[name="EndEtg_produtor_rural_status_PF"]:checked').val())
            formulario.EndEtg_produtor_rural_status.value = "";
    }

    //os campos a mais são enviados junto. Deixamos enviar...
<%end if%>
<%end if%>
}

//para mudar o tipo do endereço de entrega
function trocarEndEtgTipoPessoa(novoTipo) {
<%if blnUsarMemorizacaoCompletaEnderecos then%>
    if (novoTipo && $('input[name="EndEtg_tipo_pessoa"]:disabled').length == 0)
        setarValorRadio($('input[name="EndEtg_tipo_pessoa"]'), novoTipo);

    var pj = $('input[name="EndEtg_tipo_pessoa"]:checked').val() == "PJ";

    if (pj) {
        $(".Mostrar_EndEtg_pf").css("display", "none");
        $(".Mostrar_EndEtg_pj").css("display", "");
        $("#Label_EndEtg_nome").text("RAZÃO SOCIAL");
    }
    else {
        //display block prejudica as tabelas
        $(".Mostrar_EndEtg_pf").css("display", "");
        $(".Mostrar_EndEtg_pj").css("display", "none");
        $("#Label_EndEtg_nome").text("NOME");
    }
<%else%>
    //oculta todos
    $(".Mostrar_EndEtg_pf").css("display", "none");
    $(".Mostrar_EndEtg_pj").css("display", "none");
    $(".Habilitar_EndEtg_outroendereco").css("display", "none");
<%end if%>
}

function trataContribuinteIcmsEndEtg_PJ(novoTipo)
{
    if (novoTipo && $('input[name="EndEtg_contribuinte_icms_status_PJ"]:disabled').length == 0)
        setarValorRadio($('input[name="EndEtg_contribuinte_icms_status_PJ"]'),novoTipo);
}
function trataContribuinteIcmsEndEtg_PF(novoTipo)
{
    if (novoTipo && $('input[name="EndEtg_contribuinte_icms_status_PF"]:disabled').length == 0)
        setarValorRadio($('input[name="EndEtg_contribuinte_icms_status_PF"]'),novoTipo);
}

function trataProdutorRuralEndEtg_PF(novoTipo) {
    //ao clicar na opção Produtor Rural, exibir/ocultar os campos apropriados (endereço de entrega)
    if (novoTipo && $('input[name="EndEtg_produtor_rural_status_PF"]:disabled').length == 0)
        setarValorRadio($('input[name="EndEtg_produtor_rural_status_PF"]'), novoTipo);

    var sim = $('input[name="EndEtg_produtor_rural_status_PF"]:checked').val() == "<%=COD_ST_CLIENTE_PRODUTOR_RURAL_SIM%>";

    //contribuinte ICMS sempre aparece para PJ
    if(sim) {
        $(".Mostrar_EndEtg_contribuinte_icms_PF").css("display", "");
    }
    else {
        $(".Mostrar_EndEtg_contribuinte_icms_PF").css("display", "none");
    }
}

function trataProdutorRuralEndEtg_PJ(novoTipo) {
    if (novoTipo && $('input[name="EndEtg_produtor_rural_status_PJ"]:disabled').length == 0)
        setarValorRadio($('input[name="EndEtg_produtor_rural_status_PJ"]'), novoTipo);
}

//definir um valor como ativo em um radio 
function setarValorRadio(array, valor)
{
    for (var i = 0; i < array.length; i++)
    {
        var este = array[i];
        if (este.value == valor)
            este.checked = true;
    }
}


</script>

<script type="text/javascript">
	function exibeJanelaCEP_Cli() {
		$.mostraJanelaCEP("cep", "uf", "cidade", "bairro", "endereco", "endereco_numero", "endereco_complemento");
	}

	function exibeJanelaCEP_Etg() {
		$.mostraJanelaCEP("EndEtg_cep", "EndEtg_uf", "EndEtg_cidade", "EndEtg_bairro", "EndEtg_endereco", "EndEtg_endereco_numero", "EndEtg_endereco_complemento");
	}

	function exibeJanelaCEP_Orc() {
		$.mostraJanelaCEP("orcamento_endereco_cep", "orcamento_endereco_uf", "orcamento_endereco_cidade", "orcamento_endereco_bairro", "orcamento_endereco_logradouro", "orcamento_endereco_numero", "orcamento_endereco_complemento");
    }

	function trataProdutorRural() {
		//ao clicar na opção Produtor Rural, exibir/ocultar os campos apropriados
		if ((typeof (fCAD.rb_produtor_rural) !== "undefined") && (fCAD.rb_produtor_rural !== null)) {
			if (!fCAD.rb_produtor_rural[1].checked) {
				$("#t_contribuinte_icms").css("display", "none");
			}
			else {
				$("#t_contribuinte_icms").css("display", "block");
			}
		}
	}

<%if blnUsarMemorizacaoCompletaEnderecos then%>
    function copiarDadosCadastrais() {
        <%if not eh_cpf then %>
            fORC.orcamento_endereco_nome.value = fCAD.nome.value;
            fORC.orcamento_endereco_contribuinte_icms_status[0].checked = fCAD.rb_contribuinte_icms[0].checked;
            fORC.orcamento_endereco_contribuinte_icms_status[1].checked = fCAD.rb_contribuinte_icms[1].checked;
            fORC.orcamento_endereco_contribuinte_icms_status[2].checked = fCAD.rb_contribuinte_icms[2].checked;
            fORC.orcamento_endereco_ie.value = fCAD.ie.value;
        <% end if%>

        fORC.orcamento_endereco_logradouro.value = fCAD.endereco.value;
        fORC.orcamento_endereco_numero.value = fCAD.endereco_numero.value;
        fORC.orcamento_endereco_complemento.value = fCAD.endereco_complemento.value;
        fORC.orcamento_endereco_bairro.value = fCAD.bairro.value;
        fORC.orcamento_endereco_cidade.value = fCAD.cidade.value;
        fORC.orcamento_endereco_uf.value = fCAD.uf.value;
        fORC.orcamento_endereco_cep.value = fCAD.cep.value;

        <%if eh_cpf then %>
            fORC.orcamento_endereco_ddd_res.value = fCAD.ddd_res.value;
            fORC.orcamento_endereco_tel_res.value = fCAD.tel_res.value;
            fORC.orcamento_endereco_ddd_cel.value = fCAD.ddd_cel.value;
            fORC.orcamento_endereco_tel_cel.value = fCAD.tel_cel.value;
            fORC.orcamento_endereco_ddd_com.value = fCAD.ddd_com.value;
            fORC.orcamento_endereco_tel_com.value = fCAD.tel_com.value;
            fORC.orcamento_endereco_ramal_com.value = fCAD.ramal_com.value;
        <%else %>

            fORC.orcamento_endereco_ddd_com.value = fCAD.ddd_com.value;
            fORC.orcamento_endereco_tel_com.value = fCAD.tel_com.value;
            fORC.orcamento_endereco_ramal_com.value = fCAD.ramal_com.value;
            fORC.orcamento_endereco_ddd_com_2.value = fCAD.ddd_com_2.value;
            fORC.orcamento_endereco_tel_com_2.value = fCAD.tel_com_2.value;
            fORC.orcamento_endereco_ramal_com_2.value = fCAD.ramal_com_2.value;
            fORC.orcamento_endereco_contato.value = fCAD.contato.value;
        <% end if %>

        fORC.orcamento_endereco_email.value = fCAD.email.value;
        fORC.orcamento_endereco_email_xml.value = fCAD.email_xml.value;
    }
<%end if%>

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

<%	if operacao_selecionada=OP_INCLUI then
		if eh_cpf then
			s = "fCAD.rg.focus();"
		else
			s = "fCAD.ie.focus();"
			end if
	else
		s = "focus();"
		end if
%>
<body id="corpoPagina" onload="<%=s%>conteudo_original=retorna_dados_formulario(fCAD);ja_carregou=true;trataProdutorRural();">

<center>

<!-- #include file = "../global/JanelaBuscaCEP.htm"    -->


<!--  CADASTRO DO CLIENTE -->

<table width="698" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
<%	if operacao_selecionada=OP_INCLUI then
		s = "Cadastro de Novo Cliente"
		s_readonly = ""
		s_disabled = ""
		s_style_visibility = ""
	else
		s = "Consulta de Cliente Cadastrado"
		s_readonly = " readonly tabindex=-1"
		s_disabled = " disabled tabindex=-1"
		s_style_visibility = "visibility:hidden;"
		end if
%>
	<td align="center" valign="bottom"><p class="PEDIDO"><%=s%><br><span class="C">&nbsp;</span></p></td>
</tr>
</table>
<br>

<!-- ************   EXIBE OBSERVAÇÕES CREDITÍCIAS?  ************ -->
<%	if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("obs_crediticias")) else s=""
	if s <> "" then %>
		<span class="Lbl" style="display:none">OBSERVAÇÕES CREDITÍCIAS</span>
		<div class='MtAviso' style="width:649px;FONT-WEIGHT:bold;border:1pt solid black;display:none;" align="CENTER"><P style='margin:5px 2px 5px 2px;'><%=s%></p></div>
		<br>
	<% end if %>


<!-- ************  CAMPOS DO CADASTRO  ************ -->
<form id="fCAD" name="fCAD" method="post" action="ClienteAtualiza.asp">
<input type="hidden" name="operacao_selecionada" id="operacao_selecionada" value='<%=operacao_selecionada%>'>
<input type="hidden" name="cliente_selecionado" id="cliente_selecionado" value='<%=id_cliente%>'>
<input type="hidden" name="pagina_retorno" id="pagina_retorno" value='<%=pagina_retorno%>'>

<%if operacao_selecionada=OP_CONSULTA then%>
<INPUT type="hidden" name='contribuinte_icms_cadastrado' id="contribuinte_icms_cadastrado" value='<%=Trim("" & rs("contribuinte_icms_status"))%>'>
<INPUT type="hidden" name='produtor_rural_cadastrado' id="produtor_rural_cadastrado" value='<%=Trim("" & rs("produtor_rural_status"))%>'>
<%else%>
<INPUT type="hidden" name='contribuinte_icms_cadastrado' id="contribuinte_icms_cadastrado" value=''>
<INPUT type="hidden" name='produtor_rural_cadastrado' id="produtor_rural_cadastrado" value=''>
<%end if%>

<input type="hidden" name="c_FormFieldValues" id="c_FormFieldValues" value="" />



<!-- ************   CNPJ/IE OU CPF/RG/NASCIMENTO/SEXO  ************ -->
<table width="649" class="Q" cellspacing="0">
	<tr>
		<td width="210" align="left">
	<%if eh_cpf then s="CPF" else s="CNPJ"%>
	<p class="R"><%=s%></p><p class="C">
	<%	if operacao_selecionada=OP_CONSULTA then
			s=Trim("" & rs("cnpj_cpf"))
			s=cnpj_cpf_formata(s)
		else
			s=cnpj_cpf_formata(cnpj_cpf_selecionado)
			end if
	%>
	<input id="cnpj_cpf_selecionado" name="cnpj_cpf_selecionado" class="TA" value="<%=s%>" readonly tabindex=-1 size="22" style="text-align:center; color:#0000ff"></p></td>

<%if eh_cpf then%>
	<td class="MDE" width="210" align="left"><p class="R">RG</p><p class="C">
		<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("rg")) else s=""%>
		<input id="rg" name="rg" <%=s_readonly%> class="TA" type="text" maxlength="20" size="22" value="<%=s%>" onkeypress="if (digitou_enter(true)) fCAD.dt_nasc.focus(); filtra_nome_identificador();"></p></td>
	<td class="MD" align="left"><p class="R">NASCIMENTO</p><p class="C">
		<%if operacao_selecionada=OP_CONSULTA then s=formata_data(rs("dt_nasc")) else s=""%>
		<input id="dt_nasc" name="dt_nasc" <%=s_readonly%> class="TA" type="text" maxlength="10" size="14" value="<%=s%>" onkeypress="if (digitou_enter(true) && isDate(this)) fCAD.sexo.focus(); filtra_data();" onblur="if (tem_info(this.value)) if (!isDate(this)) {alert('Data inválida!!');this.focus();}"></p></td>
	<td align="left"><p class="R">SEXO</p><p class="C">
		<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("sexo")) else s=""%>
		<input id="sexo" name="sexo" <%=s_readonly%> class="TA" type="text" maxlength="1" size="2" value="<%=s%>" onkeypress="if (digitou_enter(true)) if (!tem_info(this.value)) fCAD.nome.focus(); else if (sexo_ok(this.value)) fCAD.nome.focus(); filtra_sexo();" onkeyup="this.value=ucase(this.value);"></p></td>

<%else%>
	<td class="MDE" width="215" align="left"><p class="R">IE</p><p class="C">
		<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("ie")) else s=""%>
		<input id="ie" name="ie" class="TA" type="text" maxlength="20" size="25" value="<%=s%>" onkeypress="if (digitou_enter(true)) fCAD.nome.focus(); filtra_nome_identificador();"></p></td>
	<td align="left"><p class="R">CONTRIBUINTE ICMS</p><p class="C">
		<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("contribuinte_icms_status")) else s=""%>
		<%if s=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO then s_aux="checked" else s_aux=""%>
		<% intIdx = 0 %>
		<input type="radio" id="rb_contribuinte_icms_nao" name="rb_contribuinte_icms" value="<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO%>" <%=s_aux%>><span class="C" style="cursor:default" onclick="fCAD.rb_contribuinte_icms[<%=Cstr(intIdx)%>].click();">Não</span>
		<%if s=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM then s_aux="checked" else s_aux=""%>
		<% intIdx = intIdx + 1 %>
		<input type="radio" id="rb_contribuinte_icms_sim" name="rb_contribuinte_icms" value="<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM%>" <%=s_aux%>><span class="C" style="cursor:default" onclick="fCAD.rb_contribuinte_icms[<%=Cstr(intIdx)%>].click();">Sim</span>
		<%if s=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_ISENTO then s_aux="checked" else s_aux=""%>
		<% intIdx = intIdx + 1 %>
		<input type="radio" id="rb_contribuinte_icms_isento" name="rb_contribuinte_icms" value="<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_ISENTO%>" <%=s_aux%>><span class="C" style="cursor:default" onclick="fCAD.rb_contribuinte_icms[<%=Cstr(intIdx)%>].click();">Isento</span></p></td>
<%end if%>
	</tr>
</table>

<!-- ************   PRODUTOR RURAL / CONTRIBUINTE ICMS / IE ************ -->
<%if eh_cpf then%>
<table width="649" class="QS" cellspacing="0">
	<tr>
	<td align="left"><p class="R">PRODUTOR RURAL</p><p class="C">
		<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("produtor_rural_status")) else s=""%>
		<%if s=COD_ST_CLIENTE_PRODUTOR_RURAL_NAO then s_aux="checked" else s_aux=""%>
		<% intIdx = 0 %>
		<input type="radio" id="rb_produtor_rural_nao" name="rb_produtor_rural" value="<%=COD_ST_CLIENTE_PRODUTOR_RURAL_NAO%>" <%=s_aux%> onclick="trataProdutorRural();"><span class="C" style="cursor:default" onclick="fCAD.rb_produtor_rural[<%=Cstr(intIdx)%>].click();">Não</span>
		<%if s=COD_ST_CLIENTE_PRODUTOR_RURAL_SIM then s_aux="checked" else s_aux=""%>
		<% intIdx = intIdx + 1 %>
		<input type="radio" id="rb_produtor_rural_sim" name="rb_produtor_rural" value="<%=COD_ST_CLIENTE_PRODUTOR_RURAL_SIM%>" <%=s_aux%> onclick="trataProdutorRural();"><span class="C" style="cursor:default" onclick="fCAD.rb_produtor_rural[<%=Cstr(intIdx)%>].click();">Sim</span></p></td>
	</tr>
</table>

<table width="649" class="QS" cellspacing="0" id="t_contribuinte_icms" onload="trataProdutorRural();">
	<tr>
	<td width="210" class="MD" align="left"><p class="R">IE</p><p class="C">
		<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("ie")) else s=""%>
		<input id="ie" name="ie" class="TA" type="text" maxlength="20" size="25" value="<%=s%>" onkeypress="if (digitou_enter(true)) fCAD.nome.focus(); filtra_nome_identificador();"></p></td>

	<td align="left"><p class="R">CONTRIBUINTE ICMS</p><p class="C">
		<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("contribuinte_icms_status")) else s=""%>
		<%if s=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO then s_aux="checked" else s_aux=""%>
		<% intIdx = 0 %>
		<input type="radio" id="rb_contribuinte_icms_nao" name="rb_contribuinte_icms" value="<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO%>" <%=s_aux%>><span class="C" style="cursor:default" onclick="fCAD.rb_contribuinte_icms[<%=Cstr(intIdx)%>].click();">Não</span>
		<%if s=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM then s_aux="checked" else s_aux=""%>
		<% intIdx = intIdx + 1 %>
		<input type="radio" id="rb_contribuinte_icms_sim" name="rb_contribuinte_icms" value="<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM%>" <%=s_aux%>><span class="C" style="cursor:default" onclick="fCAD.rb_contribuinte_icms[<%=Cstr(intIdx)%>].click();">Sim</span>
		<%if s=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_ISENTO then s_aux="checked" else s_aux=""%>
		<% intIdx = intIdx + 1 %>
		<input type="radio" id="rb_contribuinte_icms_isento" name="rb_contribuinte_icms" value="<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_ISENTO%>" <%=s_aux%>><span class="C" style="cursor:default" onclick="fCAD.rb_contribuinte_icms[<%=Cstr(intIdx)%>].click();">Isento</span></p></td>
	</tr>
</table>
<% end if %>

<!-- ************   NOME  ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
	<%if eh_cpf then s="NOME" else s="RAZÃO SOCIAL"%>
	<td width="100%" align="left"><p class="R"><%=s%></p><p class="C">
		<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("nome")) else s=""%>
		<input id="nome" name="nome" <%=s_readonly%> class="TA" value="<%=s%>" maxlength="60" size="85" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCAD.endereco.focus(); filtra_nome_identificador();"></p></td>
	</tr>
</table>

<!-- ************   ENDEREÇO   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
	<td width="100%" align="left"><p class="R">ENDEREÇO</p><p class="C">
		<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("endereco")) else s=""%>
		<input id="endereco" name="endereco" <%=s_readonly%> class="TA" value="<%=s%>" maxlength="60" style="width:635px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCAD.endereco_numero.focus(); filtra_nome_identificador();"></p></td>
	</tr>
</table>

<!-- ************   Nº/COMPLEMENTO   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
	<td class="MD" width="50%" align="left"><p class="R">Nº</p><p class="C">
		<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("endereco_numero")) else s=""%>
		<input id="endereco_numero" name="endereco_numero" <%=s_readonly%> class="TA" value="<%=s%>" maxlength="20" style="width:310px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCAD.endereco_complemento.focus(); filtra_nome_identificador();"></p></td>
	<td align="left"><p class="R">COMPLEMENTO</p><p class="C">
		<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("endereco_complemento")) else s=""%>
		<input id="endereco_complemento" name="endereco_complemento" <%=s_readonly%> class="TA" value="<%=s%>" maxlength="60" style="width:310px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCAD.bairro.focus(); filtra_nome_identificador();"></p></td>
	</tr>
</table>

<!-- ************   BAIRRO/CIDADE   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
	<td class="MD" width="50%" align="left"><p class="R">BAIRRO</p><p class="C">
		<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("bairro")) else s=""%>
		<input id="bairro" name="bairro" <%=s_readonly%> class="TA" value="<%=s%>" maxlength="72" style="width:310px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCAD.cidade.focus(); filtra_nome_identificador();"></p></td>
	<td align="left"><p class="R">CIDADE</p><p class="C">
		<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("cidade")) else s=""%>
		<input id="cidade" name="cidade" <%=s_readonly%> class="TA" value="<%=s%>" maxlength="60" style="width:310px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCAD.uf.focus(); filtra_nome_identificador();"></p></td>
	</tr>
</table>

<!-- ************   UF/CEP   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
	<td class="MD" width="50%" align="left"><p class="R">UF</p><p class="C">
		<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("uf")) else s=""%>
		<input id="uf" name="uf" <%=s_readonly%> class="TA" value="<%=s%>" maxlength="2" size="3" onkeypress="if (digitou_enter(true) && tem_info(this.value) && uf_ok(this.value)) 
			<%if eh_cpf then Response.write "fCAD.ddd_res.focus();" else Response.Write "fCAD.ddd_com.focus();"%>" onblur="this.value=trim(this.value); if (!uf_ok(this.value)) {alert('UF inválida!!');this.focus();} else this.value=ucase(this.value);"></p></td>
	<td width="50%" align="left">
		<table width="100%" cellpadding="0" cellspacing="0">
		<tr>
			<td width="50%" align="left"><p class="R">CEP</p><p class="C">
				<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("cep")) else s=""%>
				<input id="cep" name="cep" readonly tabindex=-1 class="TA" value="<%=cep_formata(s)%>" maxlength="9" size="11" onkeypress="if (digitou_enter(true) && cep_ok(this.value)) 
					<%if eh_cpf then Response.write "fCAD.ddd_res.focus();" else Response.Write "fCAD.ddd_com.focus();"%> 
					filtra_cep();" onblur="if (!cep_ok(this.value)) {alert('CEP inválido!!');this.focus();} else this.value=cep_formata(this.value);"></p></td>
			<td align="center" width="50%">
				<% if blnPesquisaCEPAntiga then %>
				<button type="button" name="bPesqCep" id="bPesqCep" style="width:130px;font-size:10pt;<%=s_style_visibility%>" class="Botao" onclick="AbrePesquisaCep();">Pesquisar CEP</button>
				<% end if %>
				<% if blnPesquisaCEPAntiga and blnPesquisaCEPNova then Response.Write "&nbsp;" %>
				<% if blnPesquisaCEPNova then %>
				<button type="button" name="bPesqCepNovo" id="bPesqCepNovo" style="width:130px;font-size:10pt;<%=s_style_visibility%>" class="Botao" onclick="exibeJanelaCEP_Cli();">&nbsp;Busca de CEP&nbsp;</button>
				<% end if %>
			</td>
		</tr>
		</table>
	</td>
	</tr>
</table>

<!-- ************   TELEFONE RESIDENCIAL   ************ -->
<% if eh_cpf then %>
<table width="649" class="QS" cellspacing="0">
	<tr>
	<td class="MD" width="20%" align="left"><p class="R">DDD</p><p class="C">
		<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("ddd_res")) else s=""%>
		<input id="ddd_res" name="ddd_res" <%=s_readonly%> class="TA" value="<%=s%>" maxlength="4" size="5" onkeypress="if (digitou_enter(true) && ddd_ok(this.value)) fCAD.tel_res.focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inválido!!');this.focus();}"></p></td>
	<td align="left"><p class="R">TELEFONE RESIDENCIAL</p><p class="C">
		<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("tel_res")) else s=""%>
		<input id="tel_res" name="tel_res" <%=s_readonly%> class="TA" value="<%=telefone_formata(s)%>" maxlength="11" size="12" onkeypress="if (digitou_enter(true) && telefone_ok(this.value)) fCAD.ddd_cel.focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('Telefone inválido!!');this.focus();} else this.value=telefone_formata(this.value);"></p></td>
	</tr>
	<tr>
	<td class="MD MC" width="20%" align="left"><p class="R">DDD</p><p class="C">
		<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("ddd_cel")) else s=""%>
		<input id="ddd_cel" name="ddd_cel" <%=s_readonly%> class="TA" value="<%=s%>" maxlength="4" size="5" onkeypress="if (digitou_enter(true) && ddd_ok(this.value)) fCAD.tel_cel.focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inválido!!');this.focus();}"></p></td>
	<td align="left" class="MC"><p class="R">CELULAR</p><p class="C">
		<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("tel_cel")) else s=""%>
		<input id="tel_cel" name="tel_cel" <%=s_readonly%> class="TA" value="<%=telefone_formata(s)%>" maxlength="9" size="12" onkeypress="if (digitou_enter(true) && telefone_ok(this.value)) fCAD.ddd_com.focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('Número de celular inválido!!');this.focus();} else this.value=telefone_formata(this.value);"></p></td>
	</tr>
</table>
<% end if %>
	
<!-- ************   TELEFONE COMERCIAL   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
	<td class="MD" width="20%" align="left"><p class="R">DDD</p><p class="C">
		<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("ddd_com")) else s=""%>
		<input id="ddd_com" name="ddd_com" <%=s_readonly%> class="TA" value="<%=s%>" maxlength="4" size="5" onkeypress="if (digitou_enter(true) && ddd_ok(this.value)) fCAD.tel_com.focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inválido!!');this.focus();}"></p></td>
	<%if eh_cpf then s=" COMERCIAL" else s=""%>
	<td class="MD" align="left"><p class="R">TELEFONE<%=s%></p><p class="C">
		<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("tel_com")) else s=""%>
		<input id="tel_com" name="tel_com" <%=s_readonly%> class="TA" value="<%=telefone_formata(s)%>" maxlength="11" size="12" onkeypress="if (digitou_enter(true) && telefone_ok(this.value)) fCAD.ramal_com.focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('Telefone inválido!!');this.focus();} else this.value=telefone_formata(this.value);"></p></td>
	<td align="left"><p class="R">RAMAL</p><p class="C">
		<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("ramal_com")) else s=""%>
		<input id="ramal_com" name="ramal_com" <%=s_readonly%> class="TA" value="<%=s%>" maxlength="4" size="6" onkeypress="if (digitou_enter(true))
			<%if Not eh_cpf then Response.Write "fCAD.ddd_com_2.focus();" else Response.Write "filiacao.focus();" %> filtra_numerico();"></p></td>
	</tr>
	<% if Not eh_cpf then %>
	<tr>
	    <td class="MD MC" width="20%" align="left"><p class="R">DDD</p><p class="C">
	    <% if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("ddd_com_2")) else s="" %>
	    <input id="ddd_com_2" name="ddd_com_2" <%=s_readonly%> class="TA" value="<%=s%>" maxlength="4" size="5" onkeypress="if (digitou_enter(true) && ddd_ok(this.value)) fCAD.tel_com_2.focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inválido!!!');this.focus();}" /></p>  
	    </td>
	    <td class="MD MC" align="left"><p class="R">TELEFONE</p><p class="C">
	    <% if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("tel_com_2")) else s=""%>
	    <input id="tel_com_2" name="tel_com_2" <%=s_readonly%> class="TA" value="<%=telefone_formata(s)%>" maxlength="9" size="12" onkeypress="if (digitou_enter(true) && telefone_ok(this.value)) fCAD.ramal_com_2.focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('Telefone inválido!!');this.focus();} else this.value=telefone_formata(this.value);"></p>
	    </td>
	    <td align="left" class="MC"><p class="R">RAMAL</p><p class="C">
	    <% if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("ramal_com_2")) else s=""%>
	    <input id="ramal_com_2" name="ramal_com_2" <%=s_readonly%> class="TA" value="<%=s%>" maxlength="4" size="6" onkeypress="if (digitou_enter(true)) <%if eh_cpf then Response.Write "fCAD.filiacao.focus();" else Response.Write "fCAD.contato.focus();"%> filtra_numerico();" /></p>
	    </td>
	</tr>
	<% end if %>
</table>

<% if eh_cpf then %>
<!-- ************   OBSERVAÇÃO (ANTIGO CAMPO FILIAÇÃO)   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
	<td width="100%" align="left"><p class="R">OBSERVAÇÃO</p><p class="C">
		<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("filiacao")) else s=""%>
		<input id="filiacao" name="filiacao" <%=s_readonly%> class="TA" value="<%=s%>" maxlength="60" size="74" onkeypress="if (digitou_enter(true)) fCAD.email.focus(); filtra_nome_identificador();"></p></td>
	</tr>
</table>
<% else %>
<!-- ************   CONTATO   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
	<td width="100%" align="left"><p class="R">NOME DA PESSOA PARA CONTATO NA EMPRESA</p><p class="C">
		<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("contato")) else s=""%>
		<input id="contato" name="contato" <%=s_readonly%> class="TA" value="<%=s%>" maxlength="30" size="45" onkeypress="if (digitou_enter(true)) fCAD.email.focus(); filtra_nome_identificador();"></p></td>
	</tr>
</table>
<% end if %>

<!-- ************   E-MAIL   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
	<td width="100%" align="left"><p class="R">E-MAIL</p><p class="C">
		<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("email")) else s=""%>
		<input id="email" name="email" <%=s_readonly%> class="TA" value="<%=s%>" maxlength="60" size="74" onkeypress="if (digitou_enter(true)) fCAD.email_xml.focus(); filtra_email();"></p></td>
	    <input type="hidden" name="email_original" id="email_original" value="<%=s%>" />
    </tr>
</table>

<!-- ************   E-MAIL (XML)  ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
	<td width="100%" align="left"><p class="R">E-MAIL (XML)</p><p class="C">
		<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("email_xml")) else s=""%>
		<input id="email_xml" name="email_xml" <%=s_readonly%> class="TA" value="<%=s%>" maxlength="60" size="74" onkeypress="if (digitou_enter(true)) fCAD.obs_crediticias.focus(); filtra_email();"></p></td>
        <input type="hidden" name="email_xml_original" id="email_xml_original" value="<%=s%>" />
	</tr>
</table>

<!-- ************   OBS CREDITÍCIAS (INATIVO)  ************ -->
<table width="649" class="QS" cellspacing="0" style="display:none">
	<tr>
	<td width="100%" align="left"><p class="R">OBSERVAÇÕES CREDITÍCIAS</p><p class="C">
		<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("obs_crediticias")) else s=""%>
		<input id="obs_crediticias" name="obs_crediticias" <%=s_readonly%> class="TA" value="<%=s%>" maxlength="50" size="65" onkeypress="if (digitou_enter(true)) fCAD.midia.focus(); filtra_nome_identificador();"></p></td>
	</tr>
</table>

<!-- ************   MÍDIA (INATIVO)  ************ -->
<table width="649" class="QS" cellspacing="0" style="display:none">
	<tr>
	<td width="100%" align="left"><p class="R">FORMA PELA QUAL CONHECEU A BONSHOP</p><p class="C">
		<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("midia")) else s=""%>
		<select id="midia" name="midia" <%=s_disabled%> style="margin-top:4pt; margin-bottom:4pt;">
			<%=midia_monta_itens_select(s)%>
		</select>
	</tr>
</table>

<!-- ************   REF BANCÁRIA   ************ -->
<%if blnCadRefBancaria then%>
<!--  ASSEGURA CRIAÇÃO DE UM ARRAY DE CAMPOS, MESMO QUANDO HOUVER SOMENTE 1 LINHA!! -->
<input type="hidden" name="c_RefBancariaBanco" id="c_RefBancariaBanco" value="">
<input type="hidden" name="c_RefBancariaAgencia" id="c_RefBancariaAgencia" value="">
<input type="hidden" name="c_RefBancariaConta" id="c_RefBancariaConta" value="">
<input type="hidden" name="c_RefBancariaDdd" id="c_RefBancariaDdd" value="">
<input type="hidden" name="c_RefBancariaTelefone" id="c_RefBancariaTelefone" value="">
<input type="hidden" name="c_RefBancariaContato" id="c_RefBancariaContato" value="">
	<% 
	if operacao_selecionada=OP_CONSULTA then
		s="SELECT * FROM t_CLIENTE_REF_BANCARIA WHERE (id_cliente='" & Trim("" & rs("id")) & "') ORDER BY ordem"
		set tRefBancaria = cn.Execute(s)
		end if
	%>
	
	<% for intCounter=1 to int_MAX_REF_BANCARIA_CLIENTE %>
		<%
		strRefBancariaBanco=""
		strRefBancariaAgencia=""
		strRefBancariaConta=""
		strRefBancariaDdd=""
		strRefBancariaTelefone=""
		strRefBancariaContato=""
		if (operacao_selecionada=OP_CONSULTA) then
			if Not tRefBancaria.Eof then 
				strRefBancariaBanco=Trim("" & tRefBancaria("banco"))
				strRefBancariaAgencia=Trim("" & tRefBancaria("agencia"))
				strRefBancariaConta=Trim("" & tRefBancaria("conta"))
				strRefBancariaDdd=Trim("" & tRefBancaria("ddd"))
				strRefBancariaTelefone=Trim("" & tRefBancaria("telefone"))
				strRefBancariaContato=Trim("" & tRefBancaria("contato"))
				end if
			end if 
		%>
<% if Not eh_cpf then %>
<br>
<table width="649" cellpadding="0" cellspacing="0">
	<tr>
		<td width="100%" align="left">
			<table width="100%" cellspacing="0">
				<tr>
					<td width="100%" align="left">
						<p class="R">REFERÊNCIA BANCÁRIA<%if int_MAX_REF_BANCARIA_CLIENTE > 1 then Response.Write " (" & CStr(intCounter) & ")"%></p>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td width="100%" align="left">
			<table width="100%" class="QS" cellspacing="0">
				<tr>
					<td width="100%" class="MC" align="left">
						<p class="R">BANCO</p>
						<p class="C">
							<select name="c_RefBancariaBanco" id="c_RefBancariaBanco" <%=s_disabled%> style="margin:1px 10px 6px 10px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
							<%=banco_monta_itens_select(strRefBancariaBanco) %>
							</select>
						</p>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td width="100%" align="left">
			<table width="100%" class="QS" cellspacing="0">
				<tr>
					<td class="MD" align="left">
						<p class="R">AGÊNCIA</p>
						<p class="C">
							<input name="c_RefBancariaAgencia" id="c_RefBancariaAgencia" <%=s_readonly%> class="TA" maxlength="8" size="12" value="<%=strRefBancariaAgencia%>" onkeypress="if (digitou_enter(true)) fCAD.c_RefBancariaConta[<%=CStr(intCounter)%>].focus(); filtra_agencia_bancaria();">
						</p>
					</td>
					<td class="MD" align="left">
						<p class="R">CONTA</p>
						<p class="C">
							<input name="c_RefBancariaConta" id="c_RefBancariaConta" <%=s_readonly%> class="TA" maxlength="12" value="<%=strRefBancariaConta%>" onkeypress="if (digitou_enter(true)) fCAD.c_RefBancariaDdd[<%=CStr(intCounter)%>].focus(); filtra_conta_bancaria();">
						</p>
					</td>
					<td class="MD" align="left">
						<p class="R">DDD</p>
						<p class="C">
							<input name="c_RefBancariaDdd" id="c_RefBancariaDdd" <%=s_readonly%> class="TA" maxlength="2" size="4" value="<%=strRefBancariaDdd%>" onkeypress="if (digitou_enter(true)) fCAD.c_RefBancariaTelefone[<%=CStr(intCounter)%>].focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inválido!!');this.focus();}">
						</p>
					</td>
					<td align="left">
						<p class="R">TELEFONE</p>
						<p class="C">
							<input name="c_RefBancariaTelefone" id="c_RefBancariaTelefone" <%=s_readonly%> class="TA" maxlength="9" value="<%=telefone_formata(strRefBancariaTelefone)%>" onkeypress="if (digitou_enter(true)) fCAD.c_RefBancariaContato[<%=CStr(intCounter)%>].focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('Telefone inválido!!');this.focus();} else this.value=telefone_formata(this.value);">
						</p>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td width="100%" align="left">
			<table width="100%" class="QS" cellspacing="0">
				<tr>
					<td width="100%" align="left">
						<p class="R">CONTATO</p>
						<p class="C">
							<input name="c_RefBancariaContato" id="c_RefBancariaContato" <%=s_readonly%> class="TA" maxlength="40"  style="width:600px;" value="<%=strRefBancariaContato%>" onkeypress="if (digitou_enter(true)) {if (<%=CStr(intCounter+1)%>==fCAD.c_RefBancariaAgencia.length) this.focus(); else fCAD.c_RefBancariaAgencia[<%=CStr(intCounter+1)%>].focus();} filtra_nome_identificador();">
						</p>
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
<% end if %>
		<% 
		if (operacao_selecionada=OP_CONSULTA) then
			if Not tRefBancaria.Eof then tRefBancaria.MoveNext
			end if
		%>
		
	<% next %>
	
	<% 
	if operacao_selecionada=OP_CONSULTA then
		tRefBancaria.Close
		end if
	%>
<%end if%>


<!-- ************   REF PROFISSIONAL (INATIVO)   ************ -->
<%if blnCadRefProfissional then%>
<!--  ASSEGURA CRIAÇÃO DE UM ARRAY DE CAMPOS, MESMO QUANDO HOUVER SOMENTE 1 LINHA!! -->
<input type="hidden" name="c_RefProfNomeEmpresa" id="c_RefProfNomeEmpresa" value="">
<input type="hidden" name="c_RefProfCargo" id="c_RefProfCargo" value="">
<input type="hidden" name="c_RefProfDdd" id="c_RefProfDdd" value="">
<input type="hidden" name="c_RefProfTelefone" id="c_RefProfTelefone" value="">
<input type="hidden" name="c_RefProfPeriodoRegistro" id="c_RefProfPeriodoRegistro" value="">
<input type="hidden" name="c_RefProfRendimentos" id="c_RefProfRendimentos" value="">
<input type="hidden" name="c_RefProfCnpj" id="c_RefProfCnpj" value="">
	<% 
	if operacao_selecionada=OP_CONSULTA then
		s="SELECT * FROM t_CLIENTE_REF_PROFISSIONAL WHERE (id_cliente='" & Trim("" & rs("id")) & "') ORDER BY ordem"
		set tRefProfissional = cn.Execute(s)
		end if
	%>
	
	<% for intCounter=1 to int_MAX_REF_PROFISSIONAL_CLIENTE %>
		<%
		strRefProfNomeEmpresa=""
		strRefProfCargo=""
		strRefProfDdd=""
		strRefProfTelefone=""
		strRefProfPeriodoRegistro=""
		strRefProfRendimentos=""
		strRefProfCnpj=""
		if (operacao_selecionada=OP_CONSULTA) then
			if Not tRefProfissional.Eof then 
				strRefProfNomeEmpresa=Trim("" & tRefProfissional("nome_empresa"))
				strRefProfCargo=Trim("" & tRefProfissional("cargo"))
				strRefProfDdd=Trim("" & tRefProfissional("ddd"))
				strRefProfTelefone=Trim("" & tRefProfissional("telefone"))
				strRefProfPeriodoRegistro=formata_data(tRefProfissional("periodo_registro"))
				strRefProfRendimentos=formata_moeda(tRefProfissional("rendimentos"))
				strRefProfCnpj=cnpj_cpf_formata(Trim("" & tRefProfissional("cnpj")))
				end if
			end if 
		%>

<table width="649" cellpadding="0" cellspacing="0" style="display:none">
	<tr>
		<td width="100%" align="left">
			<table width="100%" cellspacing="0">
				<tr>
					<td width="100%" align="left">
						<p class="R">REFERÊNCIA PROFISSIONAL<%if int_MAX_REF_PROFISSIONAL_CLIENTE > 1 then Response.Write " (" & CStr(intCounter) & ")"%></p>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td width="100%" align="left">
			<table width="100%" class="QS" cellspacing="0">
				<tr>
					<td class="MC MD" align="left">
						<p class="R">NOME DA EMPRESA</p>
						<p class="C">
							<input name="c_RefProfNomeEmpresa" id="c_RefProfNomeEmpresa" <%=s_readonly%> class="TA" maxlength="60"  style="width:450px;" value="<%=strRefProfNomeEmpresa%>" onkeypress="if (digitou_enter(true)) fCAD.c_RefProfCnpj[<%=CStr(intCounter)%>].focus(); filtra_nome_identificador();">
						</p>
					</td>
					<td class="MC" align="left">
						<p class="R">CNPJ</p>
						<p class="C">
							<input name="c_RefProfCnpj" id="c_RefProfCnpj" <%=s_readonly%> class="TA" maxlength="18"  size="24" value="<%=strRefProfCnpj%>" onkeypress="if (digitou_enter(true) && cnpj_ok(this.value)) fCAD.c_RefProfCargo[<%=CStr(intCounter)%>].focus(); filtra_cnpj();" onblur="if (!cnpj_ok(this.value)) {alert('CNPJ inválido'); this.focus();} else this.value=cnpj_formata(this.value);">
						</p>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td width="100%" align="left">
			<table width="100%" class="QS" cellspacing="0">
				<tr>
					<td class="MD" align="left">
						<p class="R">CARGO</p>
						<p class="C">
							<input name="c_RefProfCargo" id="c_RefProfCargo" <%=s_readonly%> class="TA" maxlength="40" style="width:350px;" value="<%=strRefProfCargo%>" onkeypress="if (digitou_enter(true)) fCAD.c_RefProfDdd[<%=CStr(intCounter)%>].focus(); filtra_nome_identificador();">
						</p>
					</td>
					<td class="MD" align="left">
						<p class="R">DDD</p>
						<p class="C">
							<input name="c_RefProfDdd" id="c_RefProfDdd" <%=s_readonly%> class="TA" maxlength="2" size="4" value="<%=strRefProfDdd%>" onkeypress="if (digitou_enter(true)) fCAD.c_RefProfTelefone[<%=CStr(intCounter)%>].focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inválido!!');this.focus();}">
						</p>
					</td>
					<td align="left">
						<p class="R">TELEFONE</p>
						<p class="C">
							<input name="c_RefProfTelefone" id="c_RefProfTelefone" <%=s_readonly%> class="TA" maxlength="9" value="<%=telefone_formata(strRefProfTelefone)%>" onkeypress="if (digitou_enter(true)) fCAD.c_RefProfPeriodoRegistro[<%=CStr(intCounter)%>].focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('Telefone inválido!!');this.focus();} else this.value=telefone_formata(this.value);">
						</p>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td width="100%" align="left">
			<table width="100%" class="QS" cellspacing="0">
				<tr>
					<td class="MD" width="50%" align="left">
						<p class="R">REGISTRADO DESDE (DD/MM/AAAA)</p>
						<p class="C">
							<input name="c_RefProfPeriodoRegistro" id="c_RefProfPeriodoRegistro" <%=s_readonly%> class="TA" maxlength="10" value="<%=strRefProfPeriodoRegistro%>" onkeypress="if (digitou_enter(true)) fCAD.c_RefProfRendimentos[<%=CStr(intCounter)%>].focus(); filtra_data();" onblur="if (!isDate(this)) {alert('Data inválida!!');this.focus();}">
						</p>
					</td>
					<td width="50%" align="left">
						<p class="R">RENDIMENTOS (<%=SIMBOLO_MONETARIO%>)</p>
						<p class="C">
							<input name="c_RefProfRendimentos" id="c_RefProfRendimentos" <%=s_readonly%> class="TA" maxlength="18" value="<%=strRefProfRendimentos%>" onkeypress="if (digitou_enter(true)) {if (<%=CStr(intCounter+1)%>==fCAD.c_RefProfNomeEmpresa.length) this.focus(); else fCAD.c_RefProfNomeEmpresa[<%=CStr(intCounter+1)%>].focus();} filtra_moeda_positivo();" onblur="this.value=formata_moeda(this.value);">
						</p>
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
		<% 
		if (operacao_selecionada=OP_CONSULTA) then
			if Not tRefProfissional.Eof then tRefProfissional.MoveNext
			end if
		%>
		
	<% next %>
	
	<% 
	if operacao_selecionada=OP_CONSULTA then
		tRefProfissional.Close
		end if
	%>
<%end if%>


<!-- ************   REF COMERCIAL   ************ -->
<%if blnCadRefComercial then%>
<!--  ASSEGURA CRIAÇÃO DE UM ARRAY DE CAMPOS, MESMO QUANDO HOUVER SOMENTE 1 LINHA!! -->
<input type="hidden" name="c_RefComercialNomeEmpresa" id="c_RefComercialNomeEmpresa" value="">
<input type="hidden" name="c_RefComercialContato" id="c_RefComercialContato" value="">
<input type="hidden" name="c_RefComercialDdd" id="c_RefComercialDdd" value="">
<input type="hidden" name="c_RefComercialTelefone" id="c_RefComercialTelefone" value="">
	<% 
	if operacao_selecionada=OP_CONSULTA then
		s="SELECT * FROM t_CLIENTE_REF_COMERCIAL WHERE (id_cliente='" & Trim("" & rs("id")) & "') ORDER BY ordem"
		set tRefComercial = cn.Execute(s)
		end if
	%>
	
	<% for intCounter=1 to int_MAX_REF_COMERCIAL_CLIENTE %>
		<%
		strRefComercialNomeEmpresa=""
		strRefComercialContato=""
		strRefComercialDdd=""
		strRefComercialTelefone=""
		if (operacao_selecionada=OP_CONSULTA) then
			if Not tRefComercial.Eof then 
				strRefComercialNomeEmpresa=Trim("" & tRefComercial("nome_empresa"))
				strRefComercialContato=Trim("" & tRefComercial("contato"))
				strRefComercialDdd=Trim("" & tRefComercial("ddd"))
				strRefComercialTelefone=Trim("" & tRefComercial("telefone"))
				end if
			end if 
		%>
<br>
<table width="649" cellpadding="0" cellspacing="0">
	<tr>
		<td width="100%" align="left">
			<table width="100%" cellspacing="0">
				<tr>
					<td width="100%" align="left">
						<p class="R">REFERÊNCIA COMERCIAL<%if int_MAX_REF_COMERCIAL_CLIENTE > 1 then Response.Write " (" & CStr(intCounter) & ")"%></p>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td width="100%" align="left">
			<table width="100%" class="QS" cellspacing="0">
				<tr>
					<td width="100%" class="MC" align="left">
						<p class="R">NOME DA EMPRESA</p>
						<p class="C">
							<input name="c_RefComercialNomeEmpresa" id="c_RefComercialNomeEmpresa" <%=s_readonly%> class="TA" maxlength="60"  style="width:600px;" value="<%=strRefComercialNomeEmpresa%>" onkeypress="if (digitou_enter(true)) fCAD.c_RefComercialContato[<%=CStr(intCounter)%>].focus(); filtra_nome_identificador();">
						</p>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td width="100%" align="left">
			<table width="100%" class="QS" cellspacing="0">
				<tr>
					<td class="MD" align="left">
						<p class="R">CONTATO</p>
						<p class="C">
							<input name="c_RefComercialContato" id="c_RefComercialContato" <%=s_readonly%> class="TA" maxlength="40" value="<%=strRefComercialContato%>" onkeypress="if (digitou_enter(true)) fCAD.c_RefComercialDdd[<%=CStr(intCounter)%>].focus(); filtra_nome_identificador();">
						</p>
					</td>
					<td class="MD" align="left">
						<p class="R">DDD</p>
						<p class="C">
							<input name="c_RefComercialDdd" id="c_RefComercialDdd" <%=s_readonly%> class="TA" maxlength="2" size="4" value="<%=strRefComercialDdd%>" onkeypress="if (digitou_enter(true)) fCAD.c_RefComercialTelefone[<%=CStr(intCounter)%>].focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inválido!!');this.focus();}">
						</p>
					</td>
					<td align="left">
						<p class="R">TELEFONE</p>
						<p class="C">
							<input name="c_RefComercialTelefone" id="c_RefComercialTelefone" <%=s_readonly%> class="TA" maxlength="9" value="<%=telefone_formata(strRefComercialTelefone)%>" onkeypress="if (digitou_enter(true)) {if (<%=CStr(intCounter+1)%>==fCAD.c_RefComercialNomeEmpresa.length) this.focus(); else fCAD.c_RefComercialNomeEmpresa[<%=CStr(intCounter+1)%>].focus();} filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('Telefone inválido!!');this.focus();} else this.value=telefone_formata(this.value);">
						</p>
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
		<% 
		if (operacao_selecionada=OP_CONSULTA) then
			if Not tRefComercial.Eof then tRefComercial.MoveNext
			end if
		%>
		
	<% next %>
	
	<% 
	if operacao_selecionada=OP_CONSULTA then
		tRefComercial.Close
		end if
	%>
<%end if%>


<!-- ************   PJ: DADOS DO SÓCIO MAJORITÁRIO (INATIVO)  ************ -->
<%if blnCadSocioMaj then%>
<table width="649" cellpadding="0" cellspacing="0" style="display:none">
	<tr>
		<td width="100%" align="left">
			<table width="100%" cellspacing="0">
				<tr>
					<td width="100%" align="left">
						<p class="R">DADOS DO SÓCIO MAJORITÁRIO</p>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td width="100%" align="left">
			<table width="100%" class="QS" cellspacing="0">
				<tr>
				<td class="MC MD" width="85%" align="left"><p class="R">NOME</p><p class="C">
					<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("SocMaj_Nome")) else s=""%>
					<input id="c_SocioMajNome" name="c_SocioMajNome" <%=s_readonly%> class="TA" value="<%=s%>" maxlength="60" size="61" onkeypress="if (digitou_enter(true)) fCAD.c_SocioMajCpf.focus(); filtra_nome_identificador();" onblur="this.value=trim(this.value);"></p></td>
				<td class="MC" align="left"><p class="R">CPF</p><p class="C">
					<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("SocMaj_CPF")) else s=""%>
					<input id="c_SocioMajCpf" name="c_SocioMajCpf" <%=s_readonly%> class="TA" value="<%=cnpj_cpf_formata(s)%>" maxlength="14" size="15" onkeypress="if (digitou_enter(true) && cpf_ok(this.value)) fCAD.c_SocioMajBanco.focus(); filtra_numerico();" onblur="if (!cpf_ok(this.value)) {alert('CPF inválido!!');this.focus();} else this.value=cnpj_cpf_formata(this.value);"></p></td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td width="100%" align="left">
			<table width="100%" class="QS" cellspacing="0">
				<tr>
					<td width="100%" align="left">
						<p class="R">BANCO</p>
						<p class="C">
							<select name="c_SocioMajBanco" id="c_SocioMajBanco" <%=s_disabled%> style="margin:1px 10px 6px 10px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
							<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("SocMaj_banco")) else s=""%>
							<%=banco_monta_itens_select(s) %>
							</select>
						</p>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td width="100%" align="left">
			<table width="100%" class="QS" cellspacing="0">
				<tr>
					<td class="MD" align="left">
						<p class="R">AGÊNCIA</p>
						<p class="C">
							<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("SocMaj_agencia")) else s=""%>
							<input name="c_SocioMajAgencia" id="c_SocioMajAgencia" <%=s_readonly%> class="TA" maxlength="8" value="<%=s%>" onkeypress="if (digitou_enter(true)) fCAD.c_SocioMajConta.focus(); filtra_agencia_bancaria();">
						</p>
					</td>
					<td class="MD" align="left">
						<p class="R">CONTA</p>
						<p class="C">
							<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("SocMaj_conta")) else s=""%>
							<input name="c_SocioMajConta" id="c_SocioMajConta" <%=s_readonly%> class="TA" maxlength="12" value="<%=s%>" onkeypress="if (digitou_enter(true)) fCAD.c_SocioMajDdd.focus(); filtra_conta_bancaria();">
						</p>
					</td>
					<td class="MD" align="left">
						<p class="R">DDD</p>
						<p class="C">
							<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("SocMaj_ddd")) else s=""%>
							<input name="c_SocioMajDdd" id="c_SocioMajDdd" <%=s_readonly%> class="TA" maxlength="2" size="4" value="<%=s%>" onkeypress="if (digitou_enter(true)) fCAD.c_SocioMajTelefone.focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inválido!!');this.focus();}">
						</p>
					</td>
					<td align="left">
						<p class="R">TELEFONE</p>
						<p class="C">
							<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("SocMaj_telefone")) else s=""%>
							<input name="c_SocioMajTelefone" id="c_SocioMajTelefone" <%=s_readonly%> class="TA" maxlength="9" value="<%=telefone_formata(s)%>" onkeypress="if (digitou_enter(true)) fCAD.c_SocioMajContato.focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('Telefone inválido!!');this.focus();} else this.value=telefone_formata(this.value);">
						</p>
					</td>
				</tr>
			</table>
		</td>
	</tr>
	<tr>
		<td width="100%" align="left">
			<table width="100%" class="QS" cellspacing="0">
				<tr>
					<td width="100%" align="left">
						<p class="R">CONTATO</p>
						<p class="C">
							<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("SocMaj_contato")) else s=""%>
							<input name="c_SocioMajContato" id="c_SocioMajContato" <%=s_readonly%> class="TA" maxlength="40"  style="width:600px;" value="<%=s%>" onkeypress="filtra_nome_identificador();">
						</p>
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
<%end if%>

</form>


<!-- ************   FORM PARA OPÇÃO DE CADASTRAR NOVO ORÇAMENTO?  ************ -->
<% if operacao_selecionada = OP_CONSULTA then %>
	<% if blnLojaHabilitadaProdCompostoECommerce then
			s_dest = "OrcamentoNovoProdCompostoMask.asp"
		else
			s_dest = "OrcamentoNovo.asp"
		end if %>
	<form action="<%=s_dest%>" method="post" id="fORC" name="fORC" onsubmit="if (!fORCConcluir(fORC)) return false">
	<input type="hidden" name="cliente_selecionado" id="cliente_selecionado" value='<%=id_cliente%>'>
	<input type="hidden" name="operacao_selecionada" id="operacao_selecionada" value='<%=OP_INCLUI%>'>
	<input type="hidden" name="c_FormFieldValues" id="c_FormFieldValues" value="" />


<!-- ************   DADOS CADASTRAIS   ************ -->
<%if blnUsarMemorizacaoCompletaEnderecos then%>
    <br>
    <table width="649" class="Q" cellspacing="0">
	    <tr>
		    <td align="left">
		        <p class="R">DADOS CADASTRAIS</p>
		    </td>
		    <td style="width:40px;text-align:right;vertical-align:top;">
    			<a href="javascript:copiarDadosCadastrais();"><img src="../IMAGEM/copia_20x20.png" name="btnCopiarDadosCadastrais" id="btnCopiarDadosCadastrais" title="Copia os dados já existentes para o bloco de dados cadastrais" /></a>
		    </td>
	    </tr>
    </table>
    <%if eh_cpf then %>
        <input type="hidden" name="orcamento_endereco_nome" id="orcamento_endereco_nome" value="<%=rs("nome") %>" />
        <input type="hidden" name="orcamento_endereco_contribuinte_icms_status" id="orcamento_endereco_contribuinte_icms_status" value="<%=rs("contribuinte_icms_status") %>" />
        <input type="hidden" name="orcamento_endereco_ie" id="orcamento_endereco_ie" value="<%=rs("ie") %>" />

    <%else %>
        <!-- ************   DADOS CADASTRAIS PESSOA JURÍDICA   ************ -->

        <!-- ************   NOME  ************ -->
        <table width="649" class="QS" cellspacing="0">
	        <tr>
	        <td width="100%" align="left"><p class="R">RAZÃO SOCIAL</p><p class="C">
		        <input id="orcamento_endereco_nome" name="orcamento_endereco_nome" class="TA" maxlength="60" size="85" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fORC.orcamento_endereco_contribuinte_icms_status_nao.focus(); filtra_nome_identificador();"></p></td>
	        </tr>
        </table>

        <!-- ************   CONTRIBUINTE ICMS / IE ************ -->
        <table width="649" class="QS" cellspacing="0">
	        <tr>
	        <td align="left"><p class="R">CONTRIBUINTE ICMS</p><p class="C">
		        <% intIdx = 0 %>
		        <input type="radio" id="orcamento_endereco_contribuinte_icms_status_nao" name="orcamento_endereco_contribuinte_icms_status" value="<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO%>"><span class="C" style="cursor:default" onclick="fORC.orcamento_endereco_contribuinte_icms_status[<%=Cstr(intIdx)%>].click();">Não</span>
		        <% intIdx = intIdx + 1 %>
		        <input type="radio" id="orcamento_endereco_contribuinte_icms_status_sim" name="orcamento_endereco_contribuinte_icms_status" value="<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM%>"><span class="C" style="cursor:default" onclick="fORC.orcamento_endereco_contribuinte_icms_status[<%=Cstr(intIdx)%>].click();">Sim</span>
		        <% intIdx = intIdx + 1 %>
		        <input type="radio" id="orcamento_endereco_contribuinte_icms_status_isento" name="orcamento_endereco_contribuinte_icms_status" value="<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_ISENTO%>"><span class="C" style="cursor:default" onclick="fORC.orcamento_endereco_contribuinte_icms_status[<%=Cstr(intIdx)%>].click();">Isento</span></p></td>
	        <td class="MDE" width="215" align="left"><p class="R">IE</p><p class="C">
		        <input id="orcamento_endereco_ie" name="orcamento_endereco_ie" class="TA" type="text" maxlength="20" size="25" onkeypress="if (digitou_enter(true)) fORC.orcamento_endereco_logradouro.focus(); filtra_nome_identificador();"></p></td>
	        </tr>
        </table>

    <%end if%>

    <!-- ************   ENDEREÇO   ************ -->
    <table width="649" class="QS" cellspacing="0">
	    <tr>
	    <td width="100%" align="left"><p class="R">ENDEREÇO</p><p class="C">
		    <input id="orcamento_endereco_logradouro" name="orcamento_endereco_logradouro" class="TA" maxlength="60" style="width:635px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fORC.orcamento_endereco_numero.focus(); filtra_nome_identificador();"></p></td>
	    </tr>
    </table>

    <!-- ************   Nº/COMPLEMENTO   ************ -->
    <table width="649" class="QS" cellspacing="0">
	    <tr>
	    <td class="MD" width="50%" align="left"><p class="R">Nº</p><p class="C">
		    <input id="orcamento_endereco_numero" name="orcamento_endereco_numero" class="TA" maxlength="20" style="width:310px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fORC.orcamento_endereco_complemento.focus(); filtra_nome_identificador();"></p></td>
	    <td align="left"><p class="R">COMPLEMENTO</p><p class="C">
		    <input id="orcamento_endereco_complemento" name="orcamento_endereco_complemento" class="TA" maxlength="60" style="width:310px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fORC.orcamento_endereco_bairro.focus(); filtra_nome_identificador();"></p></td>
	    </tr>
    </table>

    <!-- ************   BAIRRO/CIDADE   ************ -->
    <table width="649" class="QS" cellspacing="0">
	    <tr>
	    <td class="MD" width="50%" align="left"><p class="R">BAIRRO</p><p class="C">
		    <input id="orcamento_endereco_bairro" name="orcamento_endereco_bairro" class="TA" maxlength="72" style="width:310px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fORC.orcamento_endereco_cidade.focus(); filtra_nome_identificador();"></p></td>
	    <td align="left"><p class="R">CIDADE</p><p class="C">
		    <input id="orcamento_endereco_cidade" name="orcamento_endereco_cidade" class="TA" maxlength="60" style="width:310px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fORC.orcamento_endereco_uf.focus(); filtra_nome_identificador();"></p></td>
	    </tr>
    </table>

    <!-- ************   UF/CEP   ************ -->
    <table width="649" class="QS" cellspacing="0">
	    <tr>
	    <td class="MD" width="50%" align="left"><p class="R">UF</p><p class="C">
		    <input id="orcamento_endereco_uf" name="orcamento_endereco_uf" class="TA" maxlength="2" size="3" onkeypress="if (digitou_enter(true) && tem_info(this.value) && uf_ok(this.value)) fORC.orcamento_endereco_ddd_res.focus();" onblur="this.value=trim(this.value); if (!uf_ok(this.value)) {alert('UF inválida!!');this.focus();} else this.value=ucase(this.value);"></p></td>
	    <td width="50%" align="left">
		    <table width="100%" cellpadding="0" cellspacing="0">
		    <tr>
			    <td width="50%" align="left"><p class="R">CEP</p><p class="C">
				    <input id="orcamento_endereco_cep" name="orcamento_endereco_cep" readonly tabindex=-1 class="TA" maxlength="9" size="11" onkeypress="if (digitou_enter(true) && cep_ok(this.value)) fORC.orcamento_endereco_ddd_res.focus(); filtra_cep();" onblur="if (!cep_ok(this.value)) {alert('CEP inválido!!');this.focus();} else this.value=cep_formata(this.value);"></p></td>
			    <td align="center" width="50%">
				    <% if blnPesquisaCEPAntiga then %>
				    <button type="button" name="bPesqCepOrcamento" id="bPesqCepOrcamento" style="width:130px;font-size:10pt;" class="Botao" onclick="OrcamentoAbrePesquisaCep();">Pesquisar CEP</button>
				    <% end if %>
				    <% if blnPesquisaCEPAntiga and blnPesquisaCEPNova then Response.Write "&nbsp;" %>
				    <% if blnPesquisaCEPNova then %>
				    <button type="button" name="bPesqCepNovoOrcamento" id="bPesqCepNovoOrcamento" style="width:130px;font-size:10pt;" class="Botao" onclick="exibeJanelaCEP_Orc();">&nbsp;Busca de CEP&nbsp;</button>
				    <% end if %>
			    </td>
		    </tr>
		    </table>
	    </td>
	    </tr>
    </table>

    <%if eh_cpf then %>
        <!-- ************   TELEFONES PESSOA FÍSICA   ************ -->
        <input type="hidden" name="orcamento_endereco_ddd_com_2" id="orcamento_endereco_ddd_com_2" value="<%=rs("ddd_com_2") %>" />
        <input type="hidden" name="orcamento_endereco_tel_com_2" id="orcamento_endereco_tel_com_2" value="<%=rs("tel_com_2") %>" />
        <input type="hidden" name="orcamento_endereco_ramal_com_2" id="orcamento_endereco_ramal_com_2" value="<%=rs("ramal_com_2") %>" />
        <input type="hidden" name="orcamento_endereco_tipo_pessoa" id="orcamento_endereco_tipo_pessoa" value="<%=rs("tipo") %>" />
        <input type="hidden" name="orcamento_endereco_cnpj_cpf" id="orcamento_endereco_cnpj_cpf" value="<%=rs("cnpj_cpf") %>" />
        <input type="hidden" name="orcamento_endereco_produtor_rural_status" id="orcamento_endereco_produtor_rural_status" value="<%=rs("produtor_rural_status") %>" />
        <input type="hidden" name="orcamento_endereco_rg" id="orcamento_endereco_rg" value="<%=rs("rg") %>" />
        <input type="hidden" name="orcamento_endereco_contato" id="orcamento_endereco_contato" value="<%=rs("contato") %>" />

        <!-- ************   TELEFONE RESIDENCIAL   ************ -->
        <table width="649" class="QS" cellspacing="0">
	        <tr>
	        <td class="MD" width="20%" align="left"><p class="R">DDD</p><p class="C">
		        <input id="orcamento_endereco_ddd_res" name="orcamento_endereco_ddd_res" class="TA" maxlength="4" size="5" onkeypress="if (digitou_enter(true) && ddd_ok(this.value)) fORC.orcamento_endereco_tel_res.focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inválido!!');this.focus();}"></p></td>
	        <td align="left"><p class="R">TELEFONE RESIDENCIAL</p><p class="C">
		        <input id="orcamento_endereco_tel_res" name="orcamento_endereco_tel_res" class="TA" maxlength="11" size="12" onkeypress="if (digitou_enter(true) && telefone_ok(this.value)) fORC.orcamento_endereco_ddd_cel.focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('Telefone inválido!!');this.focus();} else this.value=telefone_formata(this.value);"></p></td>
	        </tr>
	        <tr>
	        <td class="MD MC" width="20%" align="left"><p class="R">DDD</p><p class="C">
		        <input id="orcamento_endereco_ddd_cel" name="orcamento_endereco_ddd_cel" class="TA" maxlength="4" size="5" onkeypress="if (digitou_enter(true) && ddd_ok(this.value)) fORC.orcamento_endereco_tel_cel.focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inválido!!');this.focus();}"></p></td>
	        <td align="left" class="MC"><p class="R">CELULAR</p><p class="C">
		        <input id="orcamento_endereco_tel_cel" name="orcamento_endereco_tel_cel" class="TA" maxlength="9" size="12" onkeypress="if (digitou_enter(true) && telefone_ok(this.value)) fORC.orcamento_endereco_ddd_com.focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('Número de celular inválido!!');this.focus();} else this.value=telefone_formata(this.value);"></p></td>
	        </tr>
        </table>
	
        <!-- ************   TELEFONE COMERCIAL   ************ -->
        <table width="649" class="QS" cellspacing="0">
	        <tr>
	        <td class="MD" width="20%" align="left"><p class="R">DDD</p><p class="C">
		        <input id="orcamento_endereco_ddd_com" name="orcamento_endereco_ddd_com" class="TA" maxlength="4" size="5" onkeypress="if (digitou_enter(true) && ddd_ok(this.value)) fORC.orcamento_endereco_tel_com.focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inválido!!');this.focus();}"></p></td>
	        <td class="MD" align="left"><p class="R">TELEFONE COMERCIAL</p><p class="C">
		        <input id="orcamento_endereco_tel_com" name="orcamento_endereco_tel_com" class="TA" maxlength="11" size="12" onkeypress="if (digitou_enter(true) && telefone_ok(this.value)) fORC.orcamento_endereco_ramal_com.focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('Telefone inválido!!');this.focus();} else this.value=telefone_formata(this.value);"></p></td>
	        <td align="left"><p class="R">RAMAL</p><p class="C">
		        <input id="orcamento_endereco_ramal_com" name="orcamento_endereco_ramal_com" class="TA" maxlength="4" size="6" onkeypress="if (digitou_enter(true)) fORC.orcamento_endereco_email.focus(); filtra_numerico();"></p></td>
	        </tr>
        </table>
    <%else %>
        <!-- ************   TELEFONES PESSOA JURÍDICA   ************ -->
        <input type="hidden" name="orcamento_endereco_ddd_res" id="orcamento_endereco_ddd_res" value="<%=rs("ddd_res") %>" />
        <input type="hidden" name="orcamento_endereco_tel_res" id="orcamento_endereco_tel_res" value="<%=rs("tel_res") %>" />
        <input type="hidden" name="orcamento_endereco_ddd_cel" id="orcamento_endereco_ddd_cel" value="<%=rs("ddd_cel") %>" />
        <input type="hidden" name="orcamento_endereco_tel_cel" id="orcamento_endereco_tel_cel" value="<%=rs("tel_cel") %>" />
        <input type="hidden" name="orcamento_endereco_tipo_pessoa" id="orcamento_endereco_tipo_pessoa" value="<%=rs("tipo") %>" />
        <input type="hidden" name="orcamento_endereco_cnpj_cpf" id="orcamento_endereco_cnpj_cpf" value="<%=rs("cnpj_cpf") %>" />
        <input type="hidden" name="orcamento_endereco_produtor_rural_status" id="orcamento_endereco_produtor_rural_status" value="<%=rs("produtor_rural_status") %>" />
        <input type="hidden" name="orcamento_endereco_rg" id="orcamento_endereco_rg" value="<%=rs("rg") %>" />

        <!-- ************   TELEFONE COMERCIAL   ************ -->
        <table width="649" class="QS" cellspacing="0">
	        <tr>
	        <td class="MD" width="20%" align="left"><p class="R">DDD</p><p class="C">
		        <input id="orcamento_endereco_ddd_com" name="orcamento_endereco_ddd_com" class="TA" maxlength="4" size="5" onkeypress="if (digitou_enter(true) && ddd_ok(this.value)) fORC.orcamento_endereco_tel_com.focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inválido!!');this.focus();}"></p></td>
	        <td class="MD" align="left"><p class="R">TELEFONE</p><p class="C">
		        <input id="orcamento_endereco_tel_com" name="orcamento_endereco_tel_com" class="TA" maxlength="11" size="12" onkeypress="if (digitou_enter(true) && telefone_ok(this.value)) fORC.orcamento_endereco_ramal_com.focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('Telefone inválido!!');this.focus();} else this.value=telefone_formata(this.value);"></p></td>
	        <td align="left"><p class="R">RAMAL</p><p class="C">
		        <input id="orcamento_endereco_ramal_com" name="orcamento_endereco_ramal_com" class="TA" maxlength="4" size="6" onkeypress="if (digitou_enter(true)) fORC.orcamento_endereco_ddd_com_2.focus(); filtra_numerico();"></p></td>
	        </tr>
	        <tr>
	        <td class="MD MC" width="20%" align="left"><p class="R">DDD</p><p class="C">
	        <input id="orcamento_endereco_ddd_com_2" name="orcamento_endereco_ddd_com_2" class="TA" maxlength="4" size="5" onkeypress="if (digitou_enter(true) && ddd_ok(this.value)) fORC.orcamento_endereco_tel_com_2.focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inválido!!!');this.focus();}" /></p>  
	        </td>
	        <td class="MD MC" align="left"><p class="R">TELEFONE</p><p class="C">
	        <input id="orcamento_endereco_tel_com_2" name="orcamento_endereco_tel_com_2" class="TA" maxlength="9" size="12" onkeypress="if (digitou_enter(true) && telefone_ok(this.value)) fORC.orcamento_endereco_ramal_com_2.focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('Telefone inválido!!');this.focus();} else this.value=telefone_formata(this.value);"></p>
	        </td>
	        <td align="left" class="MC"><p class="R">RAMAL</p><p class="C">
	        <input id="orcamento_endereco_ramal_com_2" name="orcamento_endereco_ramal_com_2" class="TA" maxlength="4" size="6" onkeypress="if (digitou_enter(true)) fORC.orcamento_endereco_contato.focus(); filtra_numerico();" /></p>
	        </td>
	        </tr>
        </table>

        <!-- ************   CONTATO   ************ -->
        <table width="649" class="QS" cellspacing="0">
	        <tr>
	        <td width="100%" align="left"><p class="R">NOME DA PESSOA PARA CONTATO NA EMPRESA</p><p class="C">
		        <input id="orcamento_endereco_contato" name="orcamento_endereco_contato" class="TA" maxlength="30" size="45" onkeypress="if (digitou_enter(true)) fORC.orcamento_endereco_email.focus(); filtra_nome_identificador();"></p></td>
	        </tr>
        </table>

    <%end if %>

    <!-- ************   E-MAIL   ************ -->
    <table width="649" class="QS" cellspacing="0">
	    <tr>
	    <td width="100%" align="left"><p class="R">E-MAIL</p><p class="C">
		    <input id="orcamento_endereco_email" name="orcamento_endereco_email" class="TA" maxlength="60" size="74" onkeypress="if (digitou_enter(true)) fORC.orcamento_endereco_email_xml.focus(); filtra_email();"></p></td>
        </tr>
    </table>

    <!-- ************   E-MAIL (XML)  ************ -->
    <table width="649" class="QS" cellspacing="0">
	    <tr>
	    <td width="100%" align="left"><p class="R">E-MAIL (XML)</p><p class="C">
		    <input id="orcamento_endereco_email_xml" name="orcamento_endereco_email_xml" class="TA" maxlength="60" size="74" onkeypress="if (digitou_enter(true)) fORC.rb_end_entrega_nao.focus(); filtra_email();"></p></td>
	    </tr>
    </table>

<%end if%>
        

<!-- ************   ENDEREÇO DE ENTREGA: S/N   ************ -->
<br>
<table width="649" class="Q" cellspacing="0">
	<tr>
		<td align="left">
		<p class="R">ENDEREÇO DE ENTREGA</p><p class="C">
			<% intIdx = 0 %>
			<input type="radio" id="rb_end_entrega_nao" name="rb_end_entrega" value="N" onclick="Disabled_True(fORC);"><span class="C" style="cursor:default" onclick="fORC.rb_end_entrega[<%=Cstr(intIdx)%>].click();Disabled_True(fORC);">O mesmo endereço do cadastro</span>
			<% intIdx = intIdx + 1 %>
			<br><input type="radio" id="rb_end_entrega_sim" name="rb_end_entrega" value="S" onclick="Disabled_False(fORC);"><span class="C" style="cursor:default" onclick="fORC.rb_end_entrega[<%=Cstr(intIdx)%>].click();Disabled_False(fORC);">Outro endereço</span>
		</p>
		</td>
	</tr>
</table>


<!--  ************  TIPO DO ENDEREÇO DE ENTREGA: PF/PJ (SOMENTE SE O CLIENTE FOR PJ)   ************ -->

<%if blnUsarMemorizacaoCompletaEnderecos then%>
    <%if eh_cpf then%>
        <!-- ************   ENDEREÇO DE ENTREGA PARA CLIENTE PF   ************ -->
        <!-- Pegamos todos os atuais. Sem campos editáveis. -->
    <input type="hidden" id="EndEtg_tipo_pessoa" name="EndEtg_tipo_pessoa" value="PF"/>
    <input type="hidden" id="EndEtg_cnpj_cpf" name="EndEtg_cnpj_cpf" value="<%=Trim("" & rs("cnpj_cpf"))%>"/>
    <input type="hidden" id="EndEtg_ie" name="EndEtg_ie" value="<%=Trim("" & rs("ie"))%>"/>
    <input type="hidden" id="EndEtg_contribuinte_icms_status" name="EndEtg_contribuinte_icms_status" value="<%=Trim("" & rs("contribuinte_icms_status"))%>"/>
    <input type="hidden" id="EndEtg_rg" name="EndEtg_rg" value="<%=Trim("" & rs("rg"))%>"/>
    <input type="hidden" id="EndEtg_produtor_rural_status" name="EndEtg_produtor_rural_status" value="<%=Trim("" & rs("produtor_rural_status"))%>"/>
    <input type="hidden" id="EndEtg_email" name="EndEtg_email" value="<%=Trim("" & rs("email"))%>"/>
    <input type="hidden" id="EndEtg_email_xml" name="EndEtg_email_xml" value="<%=Trim("" & rs("email_xml"))%>"/>
    <input type="hidden" id="EndEtg_nome" name="EndEtg_nome" value="<%=Trim("" & rs("nome"))%>"/>


    <%else%>

    <table width="649" class="QS Habilitar_EndEtg_outroendereco" cellspacing="0">
	    <tr>
		    <td align="left">
		    <p class="R">TIPO</p><p class="C">
			    <input type="radio" id="EndEtg_tipo_pessoa_PJ" name="EndEtg_tipo_pessoa" value="PJ" onclick="trocarEndEtgTipoPessoa(null);" checked>
			    <span class="C" style="cursor:default" onclick="trocarEndEtgTipoPessoa('PJ');">Pessoa Jurídica</span>
			    &nbsp;
			    <input type="radio" id="EndEtg_tipo_pessoa_PF" name="EndEtg_tipo_pessoa" value="PF" onclick="trocarEndEtgTipoPessoa(null);">
			    <span class="C" style="cursor:default" onclick="trocarEndEtgTipoPessoa('PF');">Pessoa Física</span>
		    </p>
		    </td>
	    </tr>
    </table>

            <!-- ************   PJ: CNPJ/CONTRIBUINTE ICMS/IE - DO ENDEREÇO DE ENTREGA DE PJ ************ -->
            <!-- ************   PF: CPF/PRODUTOR RURAL/CONTRIBUINTE ICMS/IE - DO ENDEREÇO DE ENTREGA DE PJ  ************ -->
            <!-- fizemos dois conjuntos diferentes de campos porque a ordem é muito diferente -->
            <!-- EndEtg_rg EndEtg_email e EndEtg_email_xml vem diretamente do t_CLIENTE -->

    <input type="hidden" id="EndEtg_cnpj_cpf" name="EndEtg_cnpj_cpf" />
    <input type="hidden" id="EndEtg_ie" name="EndEtg_ie" />
    <input type="hidden" id="EndEtg_contribuinte_icms_status" name="EndEtg_contribuinte_icms_status" />
    <input type="hidden" id="EndEtg_rg" name="EndEtg_rg" value="<%=Trim("" & rs("rg"))%>"/>
    <input type="hidden" id="EndEtg_produtor_rural_status" name="EndEtg_produtor_rural_status" />
    <input type="hidden" id="EndEtg_email" name="EndEtg_email" value="<%=Trim("" & rs("email"))%>"/>
    <input type="hidden" id="EndEtg_email_xml" name="EndEtg_email_xml" value="<%=Trim("" & rs("email_xml"))%>"/>


    <table width="649" class="QS Habilitar_EndEtg_outroendereco Mostrar_EndEtg_pj" cellspacing="0">
	    <tr>
		    <td width="210" align="left">
	    <p class="R">CNPJ</p><p class="C">
	    <input id="EndEtg_cnpj_cpf_PJ" name="EndEtg_cnpj_cpf_PJ" class="TA" value="" size="22" style="text-align:center; color:#0000ff"></p></td>

	    <td class="MDE" width="215" align="left"><p class="R">IE</p><p class="C">
		    <input id="EndEtg_ie_PJ" name="EndEtg_ie_PJ" class="TA" type="text" maxlength="20" size="25" value="" onkeypress="if (digitou_enter(true)) fCAD.EndEtg_nome.focus(); filtra_nome_identificador();"></p></td>

	    <td align="left" class="Mostrar_EndEtg_contribuinte_icms_PJ"><p class="R">CONTRIBUINTE ICMS</p><p class="C">
		    <input type="radio" id="EndEtg_contribuinte_icms_status_PJ_nao" name="EndEtg_contribuinte_icms_status_PJ" value="<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO%>" ><span class="C" style="cursor:default" onclick="trataContribuinteIcmsEndEtg_PJ('<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO%>');">Não</span>
		    <input type="radio" id="EndEtg_contribuinte_icms_status_PJ_sim" name="EndEtg_contribuinte_icms_status_PJ" value="<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM%>" ><span class="C" style="cursor:default" onclick="trataContribuinteIcmsEndEtg_PJ('<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM%>');">Sim</span>
		    <input type="radio" id="EndEtg_contribuinte_icms_status_PJ_isento" name="EndEtg_contribuinte_icms_status_PJ" value="<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_ISENTO%>" ><span class="C" style="cursor:default" onclick="trataContribuinteIcmsEndEtg_PJ('<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_ISENTO%>');">Isento</span></p></td>
	    </tr>
    </table>

    <table width="649" class="QS Habilitar_EndEtg_outroendereco Mostrar_EndEtg_pf" cellspacing="0">
	    <tr>
		    <td width="210" align="left">
	    <p class="R">CPF</p><p class="C">
	    <input id="EndEtg_cnpj_cpf_PF" name="EndEtg_cnpj_cpf_PF" class="TA" value="" size="22" style="text-align:center; color:#0000ff"></p></td>

	    <td align="left" class="ME" style="min-width: 110px;" ><p class="R">PRODUTOR RURAL</p><p class="C">
		    <input type="radio" id="EndEtg_produtor_rural_status_PF_nao" name="EndEtg_produtor_rural_status_PF" value="<%=COD_ST_CLIENTE_PRODUTOR_RURAL_NAO%>" onclick="trataProdutorRuralEndEtg_PF(null);"><span class="C" style="cursor:default" onclick="trataProdutorRuralEndEtg_PF('<%=COD_ST_CLIENTE_PRODUTOR_RURAL_NAO%>');">Não</span>
		    <input type="radio" id="EndEtg_produtor_rural_status_PF_sim" name="EndEtg_produtor_rural_status_PF" value="<%=COD_ST_CLIENTE_PRODUTOR_RURAL_SIM%>" onclick="trataProdutorRuralEndEtg_PF(null);"><span class="C" style="cursor:default" onclick="trataProdutorRuralEndEtg_PF('<%=COD_ST_CLIENTE_PRODUTOR_RURAL_SIM%>')">Sim</span></p></td>

	    <td align="left" class="MDE Mostrar_EndEtg_contribuinte_icms_PF"><p class="R">IE</p><p class="C">
		    <input id="EndEtg_ie_PF" name="EndEtg_ie_PF" class="TA" type="text" maxlength="20" size="13" value="" onkeypress="if (digitou_enter(true)) fCAD.EndEtg_nome.focus(); filtra_nome_identificador();"></p>
	    </td>

	    <td align="left" class="Mostrar_EndEtg_contribuinte_icms_PF" ><p class="R">CONTRIBUINTE ICMS</p><p class="C">
		    <input type="radio" id="EndEtg_contribuinte_icms_status_PF_nao" name="EndEtg_contribuinte_icms_status_PF" value="<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO%>" ><span class="C" style="cursor:default" onclick="trataContribuinteIcmsEndEtg_PF('<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO%>');">Não</span>
		    <input type="radio" id="EndEtg_contribuinte_icms_status_PF_sim" name="EndEtg_contribuinte_icms_status_PF" value="<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM%>" ><span class="C" style="cursor:default" onclick="trataContribuinteIcmsEndEtg_PF('<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM%>');">Sim</span>
		    <input type="radio" id="EndEtg_contribuinte_icms_status_PF_isento" name="EndEtg_contribuinte_icms_status_PF" value="<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_ISENTO%>" ><span class="C" style="cursor:default" onclick="trataContribuinteIcmsEndEtg_PF('<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_ISENTO%>');">Isento</span></p>
	    </td>
	    </tr>
    </table>



    <!-- ************   ENDEREÇO DE ENTREGA: NOME  ************ -->
    <table width="649" class="QS" cellspacing="0">
	    <tr>
	    <td width="100%" align="left"><p class="R" id="Label_EndEtg_nome">RAZÃO SOCIAL</p><p class="C">
		    <input id="EndEtg_nome" name="EndEtg_nome" class="TA" value="" maxlength="60" size="85" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCAD.EndEtg_endereco.focus(); filtra_nome_identificador();"></p></td>
	    </tr>
    </table>


    <%end if%>
<%end if%>


<!-- ************   ENDEREÇO DE ENTREGA: ENDEREÇO   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
	<td width="100%" align="left"><p class="R">ENDEREÇO</p><p class="C">
		<input id="EndEtg_endereco" name="EndEtg_endereco" class="TA" value="" maxlength="60" style="width:635px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fORC.EndEtg_endereco_numero.focus(); filtra_nome_identificador();"></p></td>
	</tr>
</table>

<!-- ************   ENDEREÇO DE ENTREGA: Nº/COMPLEMENTO   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
	<td class="MD" width="50%" align="left"><p class="R">Nº</p><p class="C">
		<input id="EndEtg_endereco_numero" name="EndEtg_endereco_numero" class="TA" value="" maxlength="20" style="width:310px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fORC.EndEtg_endereco_complemento.focus(); filtra_nome_identificador();"></p></td>
	<td width="50%" align="left"><p class="R">COMPLEMENTO</p><p class="C">
		<input id="EndEtg_endereco_complemento" name="EndEtg_endereco_complemento" class="TA" value="" maxlength="60" style="width:310px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fORC.EndEtg_bairro.focus(); filtra_nome_identificador();"></p></td>
	</tr>
</table>

<!-- ************   ENDEREÇO DE ENTREGA: BAIRRO/CIDADE   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
	<td class="MD" width="50%" align="left"><p class="R">BAIRRO</p><p class="C">
		<input id="EndEtg_bairro" name="EndEtg_bairro" class="TA" value="" maxlength="72" style="width:310px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fORC.EndEtg_cidade.focus(); filtra_nome_identificador();"></p></td>
	<td width="50%" align="left"><p class="R">CIDADE</p><p class="C">
		<input id="EndEtg_cidade" name="EndEtg_cidade" class="TA" value="" maxlength="60" style="width:310px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fORC.EndEtg_uf.focus(); filtra_nome_identificador();"></p></td>
	</tr>
</table>

<!-- ************   ENDEREÇO DE ENTREGA: UF/CEP   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
	<td width="50%" class="MD" align="left"><p class="R">UF</p><p class="C">
		<input id="EndEtg_uf" name="EndEtg_uf" class="TA" value="" maxlength="2" size="3" onkeypress="if (digitou_enter(true) && tem_info(this.value) && uf_ok(this.value)) fORC.EndEtg_cep.focus();" onblur="this.value=trim(this.value); if (!uf_ok(this.value)) {alert('UF inválida!!');this.focus();} else this.value=ucase(this.value);"></p></td>
	<td width="50%" align="left">
		<table width="100%" cellpadding="0" cellspacing="0">
		<tr>
			<td width="50%" align="left"><p class="R">CEP</p><p class="C">
				<input id="EndEtg_cep" name="EndEtg_cep" readonly tabindex=-1 class="TA" value="" maxlength="9" size="11" onkeypress="if (digitou_enter(true) && cep_ok(this.value)) filtra_cep();" onblur="if (!cep_ok(this.value)) {alert('CEP inválido!!');this.focus();} else this.value=cep_formata(this.value);"></p></td>
			<td align="center" width="50%">
				<% if blnPesquisaCEPAntiga then %>
				<button type="button" name="bPesqCepEndEtg" id="bPesqCepEndEtg" style='width:130px;font-size:10pt;' class="Botao" onclick="AbrePesquisaCepEndEtg();">Pesquisar CEP</button>
				<% end if %>
				<% if blnPesquisaCEPAntiga and blnPesquisaCEPNova then Response.Write "&nbsp;" %>
				<% if blnPesquisaCEPNova then %>
				<button type="button" name="bPesqCepEndEtgNovo" id="bPesqCepEndEtgNovo" style='width:130px;font-size:10pt;' class="Botao" onclick="exibeJanelaCEP_Etg();">&nbsp;Busca de CEP&nbsp;</button>
				<% end if %>
			</td>
		</tr>
		</table>
	</td>
	</tr>
</table>

<%if blnUsarMemorizacaoCompletaEnderecos then%>
    <%if eh_cpf then%>

        <!-- ************   ENDEREÇO DE ENTREGA PARA PF: TELEFONES   ************ -->
        <!-- pegamos todos em branco (o usuário não poderá preencher eles) -->
        <input type="hidden" id="EndEtg_ddd_res" name="EndEtg_ddd_res" value=""/>
        <input type="hidden" id="EndEtg_tel_res" name="EndEtg_tel_res" value=""/>
        <input type="hidden" id="EndEtg_ddd_cel" name="EndEtg_ddd_cel" value=""/>
        <input type="hidden" id="EndEtg_tel_cel" name="EndEtg_tel_cel" value=""/>
        <input type="hidden" id="EndEtg_ddd_com" name="EndEtg_ddd_com" value=""/>
        <input type="hidden" id="EndEtg_tel_com" name="EndEtg_tel_com" value=""/>
        <input type="hidden" id="EndEtg_ramal_com" name="EndEtg_ramal_com" value=""/>
        <input type="hidden" id="EndEtg_ddd_com_2" name="EndEtg_ddd_com_2" value=""/>
        <input type="hidden" id="EndEtg_tel_com_2" name="EndEtg_tel_com_2" value=""/>
        <input type="hidden" id="EndEtg_ramal_com_2" name="EndEtg_ramal_com_2" value=""/>

    <%else%>
        
        <!-- ************   ENDEREÇO DE ENTREGA: TELEFONE RESIDENCIAL   ************ -->
        <table width="649" class="QS Mostrar_EndEtg_pf Habilitar_EndEtg_outroendereco" cellspacing="0">
	        <tr>
	        <td class="MD" width="20%" align="left"><p class="R">DDD</p><p class="C">
		        <input id="EndEtg_ddd_res" name="EndEtg_ddd_res" class="TA" value="" maxlength="4" size="5" onkeypress="if (digitou_enter(true) && ddd_ok(this.value)) fCAD.EndEtg_tel_res.focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inválido!!');this.focus();}"></p></td>
	        <td align="left"><p class="R">TELEFONE RESIDENCIAL</p><p class="C">
		        <input id="EndEtg_tel_res" name="EndEtg_tel_res" class="TA" value="" maxlength="11" size="12" onkeypress="if (digitou_enter(true) && telefone_ok(this.value)) fCAD.EndEtg_ddd_cel.focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('Telefone inválido!!');this.focus();} else this.value=telefone_formata(this.value);"></p></td>
	        </tr>
	        <tr>
	        <td class="MD MC" width="20%" align="left"><p class="R">DDD</p><p class="C">
		        <input id="EndEtg_ddd_cel" name="EndEtg_ddd_cel" class="TA" value="" maxlength="4" size="5" onkeypress="if (digitou_enter(true) && ddd_ok(this.value)) fCAD.EndEtg_tel_cel.focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inválido!!');this.focus();}"></p></td>
	        <td align="left" class="MC"><p class="R">CELULAR</p><p class="C">
		        <input id="EndEtg_tel_cel" name="EndEtg_tel_cel" class="TA" value="" maxlength="9" size="12" onkeypress="if (digitou_enter(true) && telefone_ok(this.value)) fCAD.EndEtg_obs.focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('Número de celular inválido!!');this.focus();} else this.value=telefone_formata(this.value);"></p></td>
	        </tr>
        </table>
	
        
        <!-- ************   ENDEREÇO DE ENTREGA: TELEFONE COMERCIAL   ************ -->
        <table width="649" class="QS Mostrar_EndEtg_pj Habilitar_EndEtg_outroendereco" cellspacing="0">
	        <tr>
	        <td class="MD" width="20%" align="left"><p class="R">DDD</p><p class="C">
		        <input id="EndEtg_ddd_com" name="EndEtg_ddd_com" class="TA" value="" maxlength="4" size="5" onkeypress="if (digitou_enter(true) && ddd_ok(this.value)) fCAD.EndEtg_tel_com.focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inválido!!');this.focus();}"></p></td>
	        <td class="MD" align="left"><p class="R">TELEFONE </p><p class="C">
		        <input id="EndEtg_tel_com" name="EndEtg_tel_com" class="TA" value="" maxlength="11" size="12" onkeypress="if (digitou_enter(true) && telefone_ok(this.value)) fCAD.EndEtg_ramal_com.focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('Telefone inválido!!');this.focus();} else this.value=telefone_formata(this.value);"></p></td>
	        <td align="left"><p class="R">RAMAL</p><p class="C">
		        <input id="EndEtg_ramal_com" name="EndEtg_ramal_com" class="TA" value="" maxlength="4" size="6" onkeypress="if (digitou_enter(true)) fCAD.EndEtg_ddd_com_2.focus(); filtra_numerico();"></p></td>
	        </tr>
	        <tr>
	            <td class="MD MC" width="20%" align="left"><p class="R">DDD</p><p class="C">
	            <input id="EndEtg_ddd_com_2" name="EndEtg_ddd_com_2" class="TA" value="" maxlength="4" size="5" onkeypress="if (digitou_enter(true) && ddd_ok(this.value)) fCAD.EndEtg_tel_com_2.focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inválido!!!');this.focus();}" /></p>  
	            </td>
	            <td class="MD MC" align="left"><p class="R">TELEFONE</p><p class="C">
	            <input id="EndEtg_tel_com_2" name="EndEtg_tel_com_2" class="TA" value="" maxlength="9" size="12" onkeypress="if (digitou_enter(true) && telefone_ok(this.value)) fCAD.EndEtg_ramal_com_2.focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('Telefone inválido!!');this.focus();} else this.value=telefone_formata(this.value);"></p>
	            </td>
	            <td align="left" class="MC"><p class="R">RAMAL</p><p class="C">
	            <input id="EndEtg_ramal_com_2" name="EndEtg_ramal_com_2" class="TA" value="" maxlength="4" size="6" onkeypress="if (digitou_enter(true)) fCAD.EndEtg_obs.focus(); filtra_numerico();" /></p>
	            </td>
	        </tr>
        </table>

    <% end if %>
<% end if %>


<!-- ************   JUSTIFIQUE O ENDEREÇO   ************ -->
<table id="obs_endereco" width="649" class="QS" cellspacing="0">
	<tr >
	<td class="M" width="50%" align="left"><p class="R">JUSTIFIQUE O ENDEREÇO</p><p class="C">
		<select id="EndEtg_obs" name="EndEtg_obs" style="margin-right:225px;">			
			 <%=codigo_descricao_monta_itens_select_por_loja(GRUPO_T_CODIGO_DESCRICAO__ENDETG_JUSTIFICATIVA, "", loja)%>
		</select></td>
	</tr>
</table>
	</form>
	</form>
<% end if %>



<!-- ************   SEPARADOR   ************ -->
<table width="698" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc" align="left">&nbsp;</td></tr>
</table>
<br>


<table class="notPrint" width="649" cellspacing="0">
<tr>
	<td align="left"><a href="javascript:history.back();" title="volta para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>

<% if (operacao_selecionada = OP_CONSULTA) Or blnEdicaoBloqueada then %>
	<td align="center"><div name="dATUALIZACONTRIB" id="dATUALIZACONTRIB">
		<a name="bATUALIZA" id="bATUALIZACONTRIB" href="javascript:AtualizaClienteContrib(fCAD)" title="atualiza o cadastro deste cliente">
		<img src="../botao/confirmar.gif" width="176" height="55" border="0"></a></div>
	</td>
	<td align="right"><div name="dORCAMENTO" id="dORCAMENTO">
		<a name="bORCAMENTO" id="bORCAMENTO" href="javascript:fORCConcluir(fORC);" title="cadastra um novo pré-pedido para este cliente">
		<img src="../botao/orcamento.gif" width="176" height="55" border="0"></a></div>
	</td>
<% else %>
	<td align="right"><div name="dATUALIZA" id="dATUALIZA">
		<a name="bATUALIZA" id="bATUALIZA" href="javascript:AtualizaCliente(fCAD)" title="atualiza o cadastro deste cliente">
		<img src="../botao/confirmar.gif" width="176" height="55" border="0"></a></div>
	</td>
<% end if %>
	
</tr>
</table>

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