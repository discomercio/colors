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
'			I N I C I A L I Z A     P � G I N A     A S P    N O    S E R V I D O R
' _____________________________________________________________________________________________


	On Error GoTo 0
	Err.Clear
	
'	EXIBI��O DE BOT�ES DE PESQUISA DE CEP
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
	
'	EST� DEFINIDA A P�GINA QUE DEVE SER EXIBIDA AP�S A ATUALIZA��O NO CADASTRO?
	pagina_retorno = trim(request("pagina_retorno"))

'	EDI��O BLOQUEADA?
	dim edicao_bloqueada, blnEdicaoBloqueada
	edicao_bloqueada = ucase(trim(request("edicao_bloqueada")))
	blnEdicaoBloqueada = False
	if edicao_bloqueada = "S" then blnEdicaoBloqueada = True


'	CONECTA COM O BANCO DE DADOS
	dim cn,rs,tRefBancaria,tRefComercial,tRefProfissional
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

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

'	REF BANC�RIA
	dim blnCadRefBancaria
	dim int_MAX_REF_BANCARIA_CLIENTE
	dim strRefBancariaBanco, strRefBancariaAgencia, strRefBancariaConta
	dim strRefBancariaDdd, strRefBancariaTelefone, strRefBancariaContato
'	O cadastro de Refer�ncia Banc�ria ser� exibido p/ PF e PJ
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
	
'	PJ: DADOS DO S�CIO MAJORIT�RIO
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
			f.EndEtg_endereco.disabled = true;
			f.EndEtg_endereco_numero.disabled = true;
			f.EndEtg_bairro.disabled = true;
			f.EndEtg_cidade.disabled = true;
			f.EndEtg_obs.disabled = true;
			f.EndEtg_uf.disabled = true;
			f.EndEtg_cep.disabled = true;
			f.bPesqCepEndEtgNovo.disabled = true;
			f.EndEtg_endereco_complemento.disabled = true;
		}
	}

	// Trata o problema em que os campos do formul�rio s�o limpos ap�s retornar � esta p�gina c/ o history.back() pela 2� vez quando ocorre erro de consist�ncia
	if (trim(fCAD.c_FormFieldValues.value) != "") {
		stringToForm(fCAD.c_FormFieldValues.value, $('#fCAD'));
	}

	if ((typeof (fORC) !== "undefined") && (fORC !== null)) {
		if (trim(fORC.c_FormFieldValues.value) != "") {
			stringToForm(fORC.c_FormFieldValues.value, $('#fORC'));
		}
	}
});
function Disabled_True(f) {

    f.EndEtg_endereco.disabled = true;
    f.EndEtg_endereco_numero.disabled = true;
    f.EndEtg_bairro.disabled = true;
    f.EndEtg_cidade.disabled = true;
    f.EndEtg_obs.disabled = true;
    f.EndEtg_uf.disabled = true;
    f.EndEtg_cep.disabled = true;
    f.bPesqCepEndEtgNovo.disabled = true;
    f.EndEtg_endereco_complemento.disabled = true;
}
function Disabled_False(f) {

    f.EndEtg_endereco.disabled = false;
    f.EndEtg_endereco_numero.disabled = false;
    f.EndEtg_bairro.disabled = false;
    f.EndEtg_cidade.disabled = false;
    f.EndEtg_obs.disabled = false;
    f.EndEtg_uf.disabled = false;
    f.EndEtg_cep.disabled = false;
    f.bPesqCepEndEtgNovo.disabled = false;
    f.EndEtg_endereco_complemento.disabled = false;
}
function ProcessaSelecaoCEP(){};

function AbrePesquisaCep(){
var f, strUrl;
	try
		{
	//  SE J� HOUVER UMA JANELA DE PESQUISA DE CEP ABERTA, GARANTE QUE ELA SER� FECHADA 
	// E UMA NOVA SER� CRIADA (EVITA PROBLEMAS C/ O 'WINDOW.OPENER')	
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
	window.status="Conclu�do";
}

function AbrePesquisaCepEndEtg(){
var f, strUrl;
	try
		{
	//  SE J� HOUVER UMA JANELA DE PESQUISA DE CEP ABERTA, GARANTE QUE ELA SER� FECHADA 
	// E UMA NOVA SER� CRIADA (EVITA PROBLEMAS C/ O 'WINDOW.OPENER')	
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
	window.status="Conclu�do";
}

function consiste_endereco_cadastro( f ) {
	
	if (trim(f.endereco.value)=="") {
		alert('Preencha o endere�o!!');
		f.endereco.focus();
		return false;
		}

	if (trim(f.endereco_numero.value)=="") {
		alert('Preencha o n�mero do endere�o!!');
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
		alert('UF inv�lida!!');
		f.uf.focus();
		return false;
		}
		
	if (trim(f.cep.value)=="") {
		alert('Informe o CEP!!');
		return false;
		}
		
	if (!cep_ok(f.cep.value)) {
		alert('CEP inv�lido!!');
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
		if (!confirm("As altera��es feitas ser�o perdidas!!\nContinua mesmo assim?")) return;
		}
	
	if ((!f.rb_end_entrega[0].checked)&&(!f.rb_end_entrega[1].checked)) {
		alert('Informe se o endere�o de entrega ser� o mesmo endere�o do cadastro ou n�o!!');
		return;
		}

	if (f.rb_end_entrega[1].checked) {
		if (trim(f.EndEtg_endereco.value)=="") {
			alert('Preencha o endere�o de entrega!!');
			f.EndEtg_endereco.focus();
			return;
			}

		if (trim(f.EndEtg_endereco_numero.value)=="") {
			alert('Preencha o n�mero do endere�o de entrega!!');
			f.EndEtg_endereco_numero.focus();
			return;
			}

		if (trim(f.EndEtg_bairro.value)=="") {
			alert('Preencha o bairro do endere�o de entrega!!');
			f.EndEtg_bairro.focus();
			return;
			}

		if (trim(f.EndEtg_cidade.value)=="") {
			alert('Preencha a cidade do endere�o de entrega!!');
			f.EndEtg_cidade.focus();
			return;
			}
		if (trim(f.EndEtg_obs.value) == "") {
		    alert('Selecione a justificativa do endere�o de entrega!!');
		    f.EndEtg_cep.focus();
		    return;
		}
		s=trim(f.EndEtg_uf.value);
		if ((s=="")||(!uf_ok(s))) {
			alert('UF inv�lida no endere�o de entrega!!');
			f.EndEtg_uf.focus();
			return;
			}
			
		if (trim(f.EndEtg_cep.value)=="") {
			alert('Informe o CEP do endere�o de entrega!!');
			f.EndEtg_cep.focus();
			return;
			}
			
		if (!cep_ok(f.EndEtg_cep.value)) {
			alert('CEP inv�lido no endere�o de entrega!!');
			f.EndEtg_cep.focus();
			return;
			}
		}

	//trecho comentado por Luiz para evitar bloqueio de altera��o de clientes sem CEP
	//if (trim(fCAD.cep.value)=="") {
	//	alert('� necess�rio preencher o CEP no cadastro do cliente!!');
	//	return;
	//	}
		
	if (eh_cpf) {
		if ((trim(fCAD.produtor_rural_cadastrado.value) == "0") ||
			((fCAD.rb_produtor_rural[0].checked) && (fCAD.produtor_rural_cadastrado.value != fCAD.rb_produtor_rural[0].value)) ||
			((fCAD.rb_produtor_rural[1].checked) && (fCAD.produtor_rural_cadastrado.value != fCAD.rb_produtor_rural[1].value))) {
			alert('� necess�rio gravar os dados do cadastro do cliente para que a op��o Produtor Rural seja gravada!');
			return;
			}
		if ((!fCAD.rb_produtor_rural[0].checked) && (!fCAD.rb_produtor_rural[1].checked)) {
			alert('Informe se o cliente � produtor rural ou n�o!!');
			return;
			}
		if (fCAD.rb_produtor_rural[1].checked) {
			if (!fCAD.rb_contribuinte_icms[1].checked) {
				alert('Para ser cadastrado como Produtor Rural, � necess�rio ser contribuinte do ICMS e possuir n� de IE!!');
				return;
			}
			if ((!fCAD.rb_contribuinte_icms[0].checked) && (!fCAD.rb_contribuinte_icms[1].checked) && (!fCAD.rb_contribuinte_icms[2].checked)) {
				alert('Informe se o cliente � contribuinte do ICMS, n�o contribuinte ou isento!!');
				return;
				}
			if ((trim(fCAD.contribuinte_icms_cadastrado.value) == "0") ||
				((fCAD.rb_contribuinte_icms[0].checked) && (fCAD.contribuinte_icms_cadastrado.value != fCAD.rb_contribuinte_icms[0].value)) ||
				((fCAD.rb_contribuinte_icms[1].checked) && (fCAD.contribuinte_icms_cadastrado.value != fCAD.rb_contribuinte_icms[1].value)) ||
				((fCAD.rb_contribuinte_icms[2].checked) && (fCAD.contribuinte_icms_cadastrado.value != fCAD.rb_contribuinte_icms[2].value))) {
				alert('� necess�rio gravar os dados do cadastro do cliente para que a op��o Contribuinte ICMS seja gravada!');
				return;
				}
			}
		}
	else {
		if ((!fCAD.rb_contribuinte_icms[0].checked) && (!fCAD.rb_contribuinte_icms[1].checked) && (!fCAD.rb_contribuinte_icms[2].checked)) {
			alert('Informe se o cliente � contribuinte do ICMS, n�o contribuinte ou isento!!');
			return;
		}
		if ((trim(fCAD.contribuinte_icms_cadastrado.value) == "0") ||
			((fCAD.rb_contribuinte_icms[0].checked) && (fCAD.contribuinte_icms_cadastrado.value != fCAD.rb_contribuinte_icms[0].value)) ||
			((fCAD.rb_contribuinte_icms[1].checked) && (fCAD.contribuinte_icms_cadastrado.value != fCAD.rb_contribuinte_icms[1].value)) ||
			((fCAD.rb_contribuinte_icms[2].checked) && (fCAD.contribuinte_icms_cadastrado.value != fCAD.rb_contribuinte_icms[2].value))) {
			alert('� necess�rio gravar os dados do cadastro do cliente para que a op��o Contribuinte ICMS seja gravada!');
			return;
			}
		}
		// Verifica se o campo IE est� vazio quando contribuinte ICMS = isento
		if (eh_cpf) {
			if (!fCAD.rb_produtor_rural[0].checked) {
				if ((fCAD.rb_contribuinte_icms[0].checked) && (fCAD.ie.value.toUpperCase().indexOf('ISEN') >= 0)) {
					alert('Se cliente � n�o contribuinte do ICMS, n�o pode ter o valor ISENTO no campo de Inscri��o Estadual!!');
					fCAD.ie.focus();
					return;
				}
				if ((fCAD.rb_contribuinte_icms[1].checked) && (fCAD.ie.value.toUpperCase().indexOf('ISEN') >= 0)) {
					alert('Se cliente � contribuinte do ICMS, n�o pode ter o valor ISENTO no campo de Inscri��o Estadual!!');
					fCAD.ie.focus();
					return;
				}
		        if (fCAD.rb_contribuinte_icms[2].checked) {
		            if (fCAD.ie.value != "") {
		                alert("Se o Contribuinte ICMS � isento, o campo IE deve ser vazio!");
		                fCAD.ie.focus();
		                return;
		            }
		        }
		    }
		}
		else {
			if ((fCAD.rb_contribuinte_icms[0].checked) && (fCAD.ie.value.toUpperCase().indexOf('ISEN') >= 0)) {
				alert('Se cliente � n�o contribuinte do ICMS, n�o pode ter o valor ISENTO no campo de Inscri��o Estadual!!');
				fCAD.ie.focus();
				return;
			}
			if ((fCAD.rb_contribuinte_icms[1].checked) && (fCAD.ie.value.toUpperCase().indexOf('ISEN') >= 0)) {
				alert('Se cliente � contribuinte do ICMS, n�o pode ter o valor ISENTO no campo de Inscri��o Estadual!!');
				fCAD.ie.focus();
				return;
			}
		    if (fCAD.rb_contribuinte_icms[2].checked) {
		        if (fCAD.ie.value != "") {
		            alert("Se o Contribuinte ICMS � isento, o campo IE deve ser vazio!");
		            fCAD.ie.focus();
		            return;
		        }
		    }
		}

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
		alert('CNPJ/CPF inv�lido!!');
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
			alert('Data inv�lida!!');
			f.dt_nasc.focus();
			return;
			}
		if ((!f.rb_produtor_rural[0].checked) && (!f.rb_produtor_rural[1].checked)) {
			alert('Informe se o cliente � produtor rural ou n�o!!');
			return;
			}
		if (!f.rb_produtor_rural[0].checked) {
			if (!fCAD.rb_contribuinte_icms[1].checked) {
				alert('Para ser cadastrado como Produtor Rural, � necess�rio ser contribuinte do ICMS e possuir n� de IE!!');
				return;
			}
			if ((!f.rb_contribuinte_icms[0].checked) && (!f.rb_contribuinte_icms[1].checked) && (!f.rb_contribuinte_icms[2].checked)) {
				alert('Informe se o cliente � contribuinte do ICMS, n�o contribuinte ou isento!!');
				return;
				}
			if ((f.rb_contribuinte_icms[1].checked) && (trim(f.ie.value) == "")) {
				alert('Se o cliente � contribuinte do ICMS a inscri��o estadual deve ser preenchida!!');
				f.ie.focus();
				return;
			}
			if ((f.rb_contribuinte_icms[0].checked) && (f.ie.value.toUpperCase().indexOf('ISEN') >= 0)) {
				alert('Se cliente � n�o contribuinte do ICMS, n�o pode ter o valor ISENTO no campo de Inscri��o Estadual!!');
				f.ie.focus();
				return;
			}
			if ((f.rb_contribuinte_icms[1].checked) && (f.ie.value.toUpperCase().indexOf('ISEN') >= 0)) {
				alert('Se cliente � contribuinte do ICMS, n�o pode ter o valor ISENTO no campo de Inscri��o Estadual!!');
				f.ie.focus();
				return;
				}
			}
		}
	else {
		//deixar de exigir preenchimento se cliente n�o � contribuinte?
		//s=trim(f.ie.value);
		//if (s=="") {
		//	alert('Preencha a Inscri��o Estadual!!');
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
			alert('Informe se o cliente � contribuinte do ICMS, n�o contribuinte ou isento!!');
			return;
			}
		if ((f.rb_contribuinte_icms[1].checked) && (trim(f.ie.value) == "")) {
			alert('Se o cliente � contribuinte do ICMS a inscri��o estadual deve ser preenchida!!');
			f.ie.focus();
			return;
		}
		if ((f.rb_contribuinte_icms[0].checked) && (f.ie.value.toUpperCase().indexOf('ISEN') >= 0)) {
			alert('Se cliente � n�o contribuinte do ICMS, n�o pode ter o valor ISENTO no campo de Inscri��o Estadual!!');
			f.ie.focus();
			return;
		}
		if ((f.rb_contribuinte_icms[1].checked) && (f.ie.value.toUpperCase().indexOf('ISEN') >= 0)) {
			alert('Se cliente � contribuinte do ICMS, n�o pode ter o valor ISENTO no campo de Inscri��o Estadual!!');
			f.ie.focus();
			return;
			}
		}

		// Verifica se o campo IE est� vazio quando contribuinte ICMS = isento
		if (eh_cpf) {
			if (!fCAD.rb_produtor_rural[0].checked) {
				if (fCAD.rb_contribuinte_icms[2].checked) {
					if (fCAD.ie.value != "") {
						alert("Se o Contribuinte ICMS � isento, o campo IE deve ser vazio!");
						fCAD.ie.focus();
						return;
					}
				}
			}
		}
		else {
			if (fCAD.rb_contribuinte_icms[2].checked) {
				if (fCAD.ie.value != "") {
					alert("Se o Contribuinte ICMS � isento, o campo IE deve ser vazio!");
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
		alert('Preencha o endere�o!!');
		f.endereco.focus();
		return;
		}

	if (trim(f.endereco_numero.value)=="") {
		alert('Preencha o n�mero do endere�o!!');
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
		alert('UF inv�lida!!');
		f.uf.focus();
		return;
		}
		
	if (trim(f.cep.value)=="") {
		alert('Informe o CEP!!');
		return;
		}
		
	if (!cep_ok(f.cep.value)) {
		alert('CEP inv�lido!!');
		f.cep.focus();
		return;
		}

	if (eh_cpf) {
		if (!ddd_ok(f.ddd_res.value)) {
			alert('DDD inv�lido!!');
			f._res.focus();
			return;
			}
		if (!telefone_ok(f.tel_res.value)) {
			alert('Telefone inv�lido!!');
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
		alert('DDD inv�lido!!');
		f.ddd_com.focus();
		return;
		}

	if (!telefone_ok(f.tel_com.value)) {
		alert('Telefone comercial inv�lido!!');
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

}
if (eh_cpf) {
    if (!ddd_ok(f.ddd_cel.value)) {
        alert('DDD inv�lido!!');
        f.ddd_cel.focus();
        return;
    }
    if (!telefone_ok(f.tel_cel.value)) {
        alert('Telefone inv�lido!!');
        f.tel_res.focus();
        return;
    }
    if ((f.ddd_cel.value == "") && (f.tel_cel.value != "")) {
        alert('Preencha o DDD do celular.');
        f.ddd_cel.focus();
        return;
    }
    if ((f.tel_cel.value == "") && (f.ddd_cel.value != "")) {
        alert('Preencha o n�mero do celular.');
        f.tel_cel.focus();
        return;
    }
}
if (!eh_cpf) {
    if (!ddd_ok(f.ddd_com_2.value)) {
        alert('DDD inv�lido!!');
        f.ddd_com_2.focus();
        return;
    }
    if (!telefone_ok(f.tel_com_2.value)) {
        alert('Telefone inv�lido!!');
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
		alert('E-mail inv�lido!!');
		f.email.focus();
		return;
		}

	if ( (trim(f.email_xml.value)!="") && (!email_ok(f.email_xml.value)) ) {
		alert('E-mail (XML) inv�lido!!');
		f.email_xml.focus();
		return;
	}

/*	if (trim(f.midia.options[f.midia.selectedIndex].value)=="") {
		alert('Indique a forma pela qual conheceu a Bonshop!!');
		return;
		}
*/

//  Ref Bancaria
		//  O cadastro de Refer�ncia Banc�ria ser� feito p/ PJ
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
		                    alert('Informe o banco no cadastro de Refer�ncia Banc�ria!!');
		                    f.c_RefBancariaBanco[i].focus();
		                    return;
		                }
		                if (trim(f.c_RefBancariaAgencia[i].value) == "") {
		                    alert('Informe a ag�ncia no cadastro de Refer�ncia Banc�ria!!');
		                    f.c_RefBancariaAgencia[i].focus();
		                    return;
		                }
		                if (trim(f.c_RefBancariaConta[i].value) == "") {
		                    alert('Informe o n�mero da conta no cadastro de Refer�ncia Banc�ria!!');
		                    f.c_RefBancariaConta[i].focus();
		                    return;
		                }
		            }
		        }
		    }
		}

//  Ref Profissional
//  O cadastro de Refer�ncia Profissional ser� feito apenas p/ PF
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
					alert('Informe o nome da empresa no cadastro de Refer�ncia Profissional!!');
					f.c_RefProfNomeEmpresa[i].focus();
					return;
					}
				if (trim(f.c_RefProfCargo[i].value)=="") {
					alert('Informe o cargo no cadastro de Refer�ncia Profissional!!');
					f.c_RefProfCargo[i].focus();
					return;
					}
				if (trim(f.c_RefProfCnpj[i].value)!="") {
					if (!cnpj_ok(f.c_RefProfCnpj[i].value)) {
						alert('CNPJ inv�lido!!');
						f.c_RefProfCnpj[i].focus();
						return;
						}
					}
				}
			}
		}
*/

//  Ref Comercial
//  O cadastro de Refer�ncia Comercial ser� feito apenas p/ PJ
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
					alert('Informe o nome da empresa no cadastro de Refer�ncia Comercial!!');
					f.c_RefComercialNomeEmpresa[i].focus();
					return;
					}
				}
			}
		}

//  Dados do S�cio Majorit�rio
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
				alert('Informe o nome do s�cio majorit�rio!!');
				f.c_SocioMajNome.focus();
				return;
				}
			}
		if (blnConsistirDadosBancarios) {
			if (trim(f.c_SocioMajBanco.value)=="") {
				alert('Informe o banco nos dados banc�rios do s�cio majorit�rio!!');
				f.c_SocioMajBanco.focus();
				return;
				}
			if (trim(f.c_SocioMajAgencia.value)=="") {
				alert('Informe a ag�ncia nos dados banc�rios do s�cio majorit�rio!!');
				f.c_SocioMajAgencia.focus();
				return;
				}
			if (trim(f.c_SocioMajConta.value)=="") {
				alert('Informe o n�mero da conta nos dados banc�rios do s�cio majorit�rio!!');
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


function AtualizaClienteContrib(f) {
	var s, eh_cpf, i;

	if (!ja_carregou) return;

	s = retorna_so_digitos(f.cnpj_cpf_selecionado.value);
	eh_cpf = false;
	if (s.length == 11) eh_cpf = true;

	if (eh_cpf) {
		if ((!f.rb_produtor_rural[0].checked) && (!f.rb_produtor_rural[1].checked)) {
			alert('Informe se o cliente � produtor rural ou n�o!!');
			return;
		}
		if (!f.rb_produtor_rural[0].checked) {
			if ((!f.rb_contribuinte_icms[0].checked) && (!f.rb_contribuinte_icms[1].checked) && (!f.rb_contribuinte_icms[2].checked)) {
				alert('Informe se o cliente � contribuinte do ICMS, n�o contribuinte ou isento!!');
				return;
			}
			if ((f.rb_contribuinte_icms[1].checked) && (trim(f.ie.value) == "")) {
				alert('Se o cliente � contribuinte do ICMS a inscri��o estadual deve ser preenchida!!');
				f.ie.focus();
				return;
			}
			if ((f.rb_contribuinte_icms[0].checked) && (f.ie.value.toUpperCase().indexOf('ISEN') >= 0)) {
				alert('Se cliente � n�o contribuinte do ICMS, n�o pode ter o valor ISENTO no campo de Inscri��o Estadual!!');
				f.ie.focus();
				return;
			}
			if ((f.rb_contribuinte_icms[1].checked) && (f.ie.value.toUpperCase().indexOf('ISEN') >= 0)) {
				alert('Se cliente � contribuinte do ICMS, n�o pode ter o valor ISENTO no campo de Inscri��o Estadual!!');
				f.ie.focus();
				return;
			}
		}
	}
	else {
		if ((!f.rb_contribuinte_icms[0].checked) && (!f.rb_contribuinte_icms[1].checked) && (!f.rb_contribuinte_icms[2].checked)) {
			alert('Informe se o cliente � contribuinte do ICMS, n�o contribuinte ou isento!!');
			return;
		}
		if ((f.rb_contribuinte_icms[1].checked) && (trim(f.ie.value) == "")) {
			alert('Se o cliente � contribuinte do ICMS a inscri��o estadual deve ser preenchida!!');
			f.ie.focus();
			return;
		}
		if ((f.rb_contribuinte_icms[0].checked) && (f.ie.value.toUpperCase().indexOf('ISEN') >= 0)) {
			alert('Se cliente � n�o contribuinte do ICMS, n�o pode ter o valor ISENTO no campo de Inscri��o Estadual!!');
			f.ie.focus();
			return;
		}
		if ((f.rb_contribuinte_icms[1].checked) && (f.ie.value.toUpperCase().indexOf('ISEN') >= 0)) {
			alert('Se cliente � contribuinte do ICMS, n�o pode ter o valor ISENTO no campo de Inscri��o Estadual!!');
			f.ie.focus();
			return;
		}
}

// Verifica se o campo IE est� vazio quando contribuinte ICMS = isento
if (eh_cpf) {
    if (!fCAD.rb_produtor_rural[0].checked) {
        if (fCAD.rb_contribuinte_icms[2].checked) {
            if (fCAD.ie.value != "") {
                alert("Se o Contribuinte ICMS � isento, o campo IE deve ser vazio!");
                fCAD.ie.focus();
                return;
            }
        }
    }
}
else {
    if (fCAD.rb_contribuinte_icms[2].checked) {
        if (fCAD.ie.value != "") {
            alert("Se o Contribuinte ICMS � isento, o campo IE deve ser vazio!");
            fCAD.ie.focus();
            return;
        }
    }
}

	fCAD.c_FormFieldValues.value = formToString($("#fCAD"));

	dATUALIZACONTRIB.style.visibility = "hidden";
	window.status = "Aguarde ...";
	f.submit();
}
</script>

<script type="text/javascript">
	function exibeJanelaCEP_Cli() {
		$.mostraJanelaCEP("cep", "uf", "cidade", "bairro", "endereco", "endereco_numero", "endereco_complemento");
	}

	function exibeJanelaCEP_Etg() {
		$.mostraJanelaCEP("EndEtg_cep", "EndEtg_uf", "EndEtg_cidade", "EndEtg_bairro", "EndEtg_endereco", "EndEtg_endereco_numero", "EndEtg_endereco_complemento");
	}

	function trataProdutorRural() {
		//ao clicar na op��o Produtor Rural, exibir/ocultar os campos apropriados
		if ((typeof (fCAD.rb_produtor_rural) !== "undefined") && (fCAD.rb_produtor_rural !== null)) {
			if (!fCAD.rb_produtor_rural[1].checked) {
				$("#t_contribuinte_icms").css("display", "none");
			}
			else {
				$("#t_contribuinte_icms").css("display", "block");
			}
		}
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

<!-- ************   EXIBE OBSERVA��ES CREDIT�CIAS?  ************ -->
<%	if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("obs_crediticias")) else s=""
	if s <> "" then %>
		<span class="Lbl" style="display:none">OBSERVA��ES CREDIT�CIAS</span>
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
		<input id="dt_nasc" name="dt_nasc" <%=s_readonly%> class="TA" type="text" maxlength="10" size="14" value="<%=s%>" onkeypress="if (digitou_enter(true) && isDate(this)) fCAD.sexo.focus(); filtra_data();" onblur="if (tem_info(this.value)) if (!isDate(this)) {alert('Data inv�lida!!');this.focus();}"></p></td>
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
		<input type="radio" id="rb_contribuinte_icms_nao" name="rb_contribuinte_icms" value="<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO%>" <%=s_aux%>><span class="C" style="cursor:default" onclick="fNEW.rb_contribuinte_icms[<%=Cstr(intIdx)%>].click();">N�o</span>
		<%if s=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM then s_aux="checked" else s_aux=""%>
		<% intIdx = intIdx + 1 %>
		<input type="radio" id="rb_contribuinte_icms_sim" name="rb_contribuinte_icms" value="<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM%>" <%=s_aux%>><span class="C" style="cursor:default" onclick="fNEW.rb_contribuinte_icms[<%=Cstr(intIdx)%>].click();">Sim</span>
		<%if s=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_ISENTO then s_aux="checked" else s_aux=""%>
		<% intIdx = intIdx + 1 %>
		<input type="radio" id="rb_contribuinte_icms_isento" name="rb_contribuinte_icms" value="<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_ISENTO%>" <%=s_aux%>><span class="C" style="cursor:default" onclick="fNEW.rb_contribuinte_icms[<%=Cstr(intIdx)%>].click();">Isento</span></p></td>
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
		<input type="radio" id="rb_produtor_rural_nao" name="rb_produtor_rural" value="<%=COD_ST_CLIENTE_PRODUTOR_RURAL_NAO%>" <%=s_aux%> onclick="trataProdutorRural();"><span class="C" style="cursor:default" onclick="fNEW.rb_produtor_rural[<%=Cstr(intIdx)%>].click();">N�o</span>
		<%if s=COD_ST_CLIENTE_PRODUTOR_RURAL_SIM then s_aux="checked" else s_aux=""%>
		<% intIdx = intIdx + 1 %>
		<input type="radio" id="rb_produtor_rural_sim" name="rb_produtor_rural" value="<%=COD_ST_CLIENTE_PRODUTOR_RURAL_SIM%>" <%=s_aux%> onclick="trataProdutorRural();"><span class="C" style="cursor:default" onclick="fNEW.rb_produtor_rural[<%=Cstr(intIdx)%>].click();">Sim</span></p></td>
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
		<input type="radio" id="rb_contribuinte_icms_nao" name="rb_contribuinte_icms" value="<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO%>" <%=s_aux%>><span class="C" style="cursor:default" onclick="fNEW.rb_contribuinte_icms[<%=Cstr(intIdx)%>].click();">N�o</span>
		<%if s=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM then s_aux="checked" else s_aux=""%>
		<% intIdx = intIdx + 1 %>
		<input type="radio" id="rb_contribuinte_icms_sim" name="rb_contribuinte_icms" value="<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM%>" <%=s_aux%>><span class="C" style="cursor:default" onclick="fNEW.rb_contribuinte_icms[<%=Cstr(intIdx)%>].click();">Sim</span>
		<%if s=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_ISENTO then s_aux="checked" else s_aux=""%>
		<% intIdx = intIdx + 1 %>
		<input type="radio" id="rb_contribuinte_icms_isento" name="rb_contribuinte_icms" value="<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_ISENTO%>" <%=s_aux%>><span class="C" style="cursor:default" onclick="fNEW.rb_contribuinte_icms[<%=Cstr(intIdx)%>].click();">Isento</span></p></td>
	</tr>
</table>
<% end if %>

<!-- ************   NOME  ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
	<%if eh_cpf then s="NOME" else s="RAZ�O SOCIAL"%>
	<td width="100%" align="left"><p class="R"><%=s%></p><p class="C">
		<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("nome")) else s=""%>
		<input id="nome" name="nome" <%=s_readonly%> class="TA" value="<%=s%>" maxlength="60" size="85" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCAD.endereco.focus(); filtra_nome_identificador();"></p></td>
	</tr>
</table>

<!-- ************   ENDERE�O   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
	<td width="100%" align="left"><p class="R">ENDERE�O</p><p class="C">
		<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("endereco")) else s=""%>
		<input id="endereco" name="endereco" <%=s_readonly%> class="TA" value="<%=s%>" maxlength="60" style="width:635px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCAD.endereco_numero.focus(); filtra_nome_identificador();"></p></td>
	</tr>
</table>

<!-- ************   N�/COMPLEMENTO   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
	<td class="MD" width="50%" align="left"><p class="R">N�</p><p class="C">
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
			<%if eh_cpf then Response.write "fCAD.ddd_res.focus();" else Response.Write "fCAD.ddd_com.focus();"%>" onblur="this.value=trim(this.value); if (!uf_ok(this.value)) {alert('UF inv�lida!!');this.focus();} else this.value=ucase(this.value);"></p></td>
	<td width="50%" align="left">
		<table width="100%" cellpadding="0" cellspacing="0">
		<tr>
			<td width="50%" align="left"><p class="R">CEP</p><p class="C">
				<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("cep")) else s=""%>
				<input id="cep" name="cep" readonly tabindex=-1 class="TA" value="<%=cep_formata(s)%>" maxlength="9" size="11" onkeypress="if (digitou_enter(true) && cep_ok(this.value)) 
					<%if eh_cpf then Response.write "fCAD.ddd_res.focus();" else Response.Write "fCAD.ddd_com.focus();"%> 
					filtra_cep();" onblur="if (!cep_ok(this.value)) {alert('CEP inv�lido!!');this.focus();} else this.value=cep_formata(this.value);"></p></td>
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
		<input id="ddd_res" name="ddd_res" <%=s_readonly%> class="TA" value="<%=s%>" maxlength="4" size="5" onkeypress="if (digitou_enter(true) && ddd_ok(this.value)) fCAD.tel_res.focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inv�lido!!');this.focus();}"></p></td>
	<td align="left"><p class="R">TELEFONE RESIDENCIAL</p><p class="C">
		<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("tel_res")) else s=""%>
		<input id="tel_res" name="tel_res" <%=s_readonly%> class="TA" value="<%=telefone_formata(s)%>" maxlength="11" size="12" onkeypress="if (digitou_enter(true) && telefone_ok(this.value)) fCAD.ddd_cel.focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('Telefone inv�lido!!');this.focus();} else this.value=telefone_formata(this.value);"></p></td>
	</tr>
	<tr>
	<td class="MD MC" width="20%" align="left"><p class="R">DDD</p><p class="C">
		<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("ddd_cel")) else s=""%>
		<input id="ddd_cel" name="ddd_cel" <%=s_readonly%> class="TA" value="<%=s%>" maxlength="4" size="5" onkeypress="if (digitou_enter(true) && ddd_ok(this.value)) fCAD.tel_cel.focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inv�lido!!');this.focus();}"></p></td>
	<td align="left" class="MC"><p class="R">CELULAR</p><p class="C">
		<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("tel_cel")) else s=""%>
		<input id="tel_cel" name="tel_cel" <%=s_readonly%> class="TA" value="<%=telefone_formata(s)%>" maxlength="9" size="12" onkeypress="if (digitou_enter(true) && telefone_ok(this.value)) fCAD.ddd_com.focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('N�mero de celular inv�lido!!');this.focus();} else this.value=telefone_formata(this.value);"></p></td>
	</tr>
</table>
<% end if %>
	
<!-- ************   TELEFONE COMERCIAL   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
	<td class="MD" width="20%" align="left"><p class="R">DDD</p><p class="C">
		<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("ddd_com")) else s=""%>
		<input id="ddd_com" name="ddd_com" <%=s_readonly%> class="TA" value="<%=s%>" maxlength="4" size="5" onkeypress="if (digitou_enter(true) && ddd_ok(this.value)) fCAD.tel_com.focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inv�lido!!');this.focus();}"></p></td>
	<%if eh_cpf then s=" COMERCIAL" else s=""%>
	<td class="MD" align="left"><p class="R">TELEFONE<%=s%></p><p class="C">
		<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("tel_com")) else s=""%>
		<input id="tel_com" name="tel_com" <%=s_readonly%> class="TA" value="<%=telefone_formata(s)%>" maxlength="11" size="12" onkeypress="if (digitou_enter(true) && telefone_ok(this.value)) fCAD.ramal_com.focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('Telefone inv�lido!!');this.focus();} else this.value=telefone_formata(this.value);"></p></td>
	<td align="left"><p class="R">RAMAL</p><p class="C">
		<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("ramal_com")) else s=""%>
		<input id="ramal_com" name="ramal_com" <%=s_readonly%> class="TA" value="<%=s%>" maxlength="4" size="6" onkeypress="if (digitou_enter(true))
			<%if Not eh_cpf then Response.Write "fCAD.ddd_com_2.focus();" else Response.Write "filiacao.focus();" %> filtra_numerico();"></p></td>
	</tr>
	<% if Not eh_cpf then %>
	<tr>
	    <td class="MD MC" width="20%" align="left"><p class="R">DDD</p><p class="C">
	    <% if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("ddd_com_2")) else s="" %>
	    <input id="ddd_com_2" name="ddd_com_2" <%=s_readonly%> class="TA" value="<%=s%>" maxlength="4" size="5" onkeypress="if (digitou_enter(true) && ddd_ok(this.value)) fCAD.tel_com_2.focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inv�lido!!!');this.focus();}" /></p>  
	    </td>
	    <td class="MD MC" align="left"><p class="R">TELEFONE</p><p class="C">
	    <% if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("tel_com_2")) else s=""%>
	    <input id="tel_com_2" name="tel_com_2" <%=s_readonly%> class="TA" value="<%=telefone_formata(s)%>" maxlength="9" size="12" onkeypress="if (digitou_enter(true) && telefone_ok(this.value)) fCAD.ramal_com_2.focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('Telefone inv�lido!!');this.focus();} else this.value=telefone_formata(this.value);"></p>
	    </td>
	    <td align="left" class="MC"><p class="R">RAMAL</p><p class="C">
	    <% if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("ramal_com_2")) else s=""%>
	    <input id="ramal_com_2" name="ramal_com_2" <%=s_readonly%> class="TA" value="<%=s%>" maxlength="4" size="6" onkeypress="if (digitou_enter(true)) <%if eh_cpf then Response.Write "fCAD.filiacao.focus();" else Response.Write "fCAD.contato.focus();"%> filtra_numerico();" /></p>
	    </td>
	</tr>
	<% end if %>
</table>

<% if eh_cpf then %>
<!-- ************   OBSERVA��O (ANTIGO CAMPO FILIA��O)   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
	<td width="100%" align="left"><p class="R">OBSERVA��O</p><p class="C">
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
	</tr>
</table>

<!-- ************   E-MAIL (XML)  ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
	<td width="100%" align="left"><p class="R">E-MAIL (XML)</p><p class="C">
		<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("email_xml")) else s=""%>
		<input id="email_xml" name="email_xml" <%=s_readonly%> class="TA" value="<%=s%>" maxlength="60" size="74" onkeypress="if (digitou_enter(true)) fCAD.obs_crediticias.focus(); filtra_email();"></p></td>
	</tr>
</table>

<!-- ************   OBS CREDIT�CIAS (INATIVO)  ************ -->
<table width="649" class="QS" cellspacing="0" style="display:none">
	<tr>
	<td width="100%" align="left"><p class="R">OBSERVA��ES CREDIT�CIAS</p><p class="C">
		<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("obs_crediticias")) else s=""%>
		<input id="obs_crediticias" name="obs_crediticias" <%=s_readonly%> class="TA" value="<%=s%>" maxlength="50" size="65" onkeypress="if (digitou_enter(true)) fCAD.midia.focus(); filtra_nome_identificador();"></p></td>
	</tr>
</table>

<!-- ************   M�DIA (INATIVO)  ************ -->
<table width="649" class="QS" cellspacing="0" style="display:none">
	<tr>
	<td width="100%" align="left"><p class="R">FORMA PELA QUAL CONHECEU A BONSHOP</p><p class="C">
		<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("midia")) else s=""%>
		<select id="midia" name="midia" <%=s_disabled%> style="margin-top:4pt; margin-bottom:4pt;">
			<%=midia_monta_itens_select(s)%>
		</select>
	</tr>
</table>

<!-- ************   REF BANC�RIA   ************ -->
<%if blnCadRefBancaria then%>
<!--  ASSEGURA CRIA��O DE UM ARRAY DE CAMPOS, MESMO QUANDO HOUVER SOMENTE 1 LINHA!! -->
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
						<p class="R">REFER�NCIA BANC�RIA<%if int_MAX_REF_BANCARIA_CLIENTE > 1 then Response.Write " (" & CStr(intCounter) & ")"%></p>
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
						<p class="R">AG�NCIA</p>
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
							<input name="c_RefBancariaDdd" id="c_RefBancariaDdd" <%=s_readonly%> class="TA" maxlength="2" size="4" value="<%=strRefBancariaDdd%>" onkeypress="if (digitou_enter(true)) fCAD.c_RefBancariaTelefone[<%=CStr(intCounter)%>].focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inv�lido!!');this.focus();}">
						</p>
					</td>
					<td align="left">
						<p class="R">TELEFONE</p>
						<p class="C">
							<input name="c_RefBancariaTelefone" id="c_RefBancariaTelefone" <%=s_readonly%> class="TA" maxlength="9" value="<%=telefone_formata(strRefBancariaTelefone)%>" onkeypress="if (digitou_enter(true)) fCAD.c_RefBancariaContato[<%=CStr(intCounter)%>].focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('Telefone inv�lido!!');this.focus();} else this.value=telefone_formata(this.value);">
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
<!--  ASSEGURA CRIA��O DE UM ARRAY DE CAMPOS, MESMO QUANDO HOUVER SOMENTE 1 LINHA!! -->
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
						<p class="R">REFER�NCIA PROFISSIONAL<%if int_MAX_REF_PROFISSIONAL_CLIENTE > 1 then Response.Write " (" & CStr(intCounter) & ")"%></p>
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
							<input name="c_RefProfCnpj" id="c_RefProfCnpj" <%=s_readonly%> class="TA" maxlength="18"  size="24" value="<%=strRefProfCnpj%>" onkeypress="if (digitou_enter(true) && cnpj_ok(this.value)) fCAD.c_RefProfCargo[<%=CStr(intCounter)%>].focus(); filtra_cnpj();" onblur="if (!cnpj_ok(this.value)) {alert('CNPJ inv�lido'); this.focus();} else this.value=cnpj_formata(this.value);">
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
							<input name="c_RefProfDdd" id="c_RefProfDdd" <%=s_readonly%> class="TA" maxlength="2" size="4" value="<%=strRefProfDdd%>" onkeypress="if (digitou_enter(true)) fCAD.c_RefProfTelefone[<%=CStr(intCounter)%>].focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inv�lido!!');this.focus();}">
						</p>
					</td>
					<td align="left">
						<p class="R">TELEFONE</p>
						<p class="C">
							<input name="c_RefProfTelefone" id="c_RefProfTelefone" <%=s_readonly%> class="TA" maxlength="9" value="<%=telefone_formata(strRefProfTelefone)%>" onkeypress="if (digitou_enter(true)) fCAD.c_RefProfPeriodoRegistro[<%=CStr(intCounter)%>].focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('Telefone inv�lido!!');this.focus();} else this.value=telefone_formata(this.value);">
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
							<input name="c_RefProfPeriodoRegistro" id="c_RefProfPeriodoRegistro" <%=s_readonly%> class="TA" maxlength="10" value="<%=strRefProfPeriodoRegistro%>" onkeypress="if (digitou_enter(true)) fCAD.c_RefProfRendimentos[<%=CStr(intCounter)%>].focus(); filtra_data();" onblur="if (!isDate(this)) {alert('Data inv�lida!!');this.focus();}">
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
<!--  ASSEGURA CRIA��O DE UM ARRAY DE CAMPOS, MESMO QUANDO HOUVER SOMENTE 1 LINHA!! -->
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
						<p class="R">REFER�NCIA COMERCIAL<%if int_MAX_REF_COMERCIAL_CLIENTE > 1 then Response.Write " (" & CStr(intCounter) & ")"%></p>
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
							<input name="c_RefComercialDdd" id="c_RefComercialDdd" <%=s_readonly%> class="TA" maxlength="2" size="4" value="<%=strRefComercialDdd%>" onkeypress="if (digitou_enter(true)) fCAD.c_RefComercialTelefone[<%=CStr(intCounter)%>].focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inv�lido!!');this.focus();}">
						</p>
					</td>
					<td align="left">
						<p class="R">TELEFONE</p>
						<p class="C">
							<input name="c_RefComercialTelefone" id="c_RefComercialTelefone" <%=s_readonly%> class="TA" maxlength="9" value="<%=telefone_formata(strRefComercialTelefone)%>" onkeypress="if (digitou_enter(true)) {if (<%=CStr(intCounter+1)%>==fCAD.c_RefComercialNomeEmpresa.length) this.focus(); else fCAD.c_RefComercialNomeEmpresa[<%=CStr(intCounter+1)%>].focus();} filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('Telefone inv�lido!!');this.focus();} else this.value=telefone_formata(this.value);">
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


<!-- ************   PJ: DADOS DO S�CIO MAJORIT�RIO (INATIVO)  ************ -->
<%if blnCadSocioMaj then%>
<table width="649" cellpadding="0" cellspacing="0" style="display:none">
	<tr>
		<td width="100%" align="left">
			<table width="100%" cellspacing="0">
				<tr>
					<td width="100%" align="left">
						<p class="R">DADOS DO S�CIO MAJORIT�RIO</p>
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
					<input id="c_SocioMajCpf" name="c_SocioMajCpf" <%=s_readonly%> class="TA" value="<%=cnpj_cpf_formata(s)%>" maxlength="14" size="15" onkeypress="if (digitou_enter(true) && cpf_ok(this.value)) fCAD.c_SocioMajBanco.focus(); filtra_numerico();" onblur="if (!cpf_ok(this.value)) {alert('CPF inv�lido!!');this.focus();} else this.value=cnpj_cpf_formata(this.value);"></p></td>
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
						<p class="R">AG�NCIA</p>
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
							<input name="c_SocioMajDdd" id="c_SocioMajDdd" <%=s_readonly%> class="TA" maxlength="2" size="4" value="<%=s%>" onkeypress="if (digitou_enter(true)) fCAD.c_SocioMajTelefone.focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inv�lido!!');this.focus();}">
						</p>
					</td>
					<td align="left">
						<p class="R">TELEFONE</p>
						<p class="C">
							<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("SocMaj_telefone")) else s=""%>
							<input name="c_SocioMajTelefone" id="c_SocioMajTelefone" <%=s_readonly%> class="TA" maxlength="9" value="<%=telefone_formata(s)%>" onkeypress="if (digitou_enter(true)) fCAD.c_SocioMajContato.focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('Telefone inv�lido!!');this.focus();} else this.value=telefone_formata(this.value);">
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


<!-- ************   FORM PARA OP��O DE CADASTRAR NOVO OR�AMENTO?  ************ -->
<% if operacao_selecionada = OP_CONSULTA then %>
	<% if isLojaHabilitadaProdCompostoECommerce(loja) then
			s_dest = "OrcamentoNovoProdCompostoMask.asp"
		else
			s_dest = "OrcamentoNovo.asp"
		end if %>
	<form action="<%=s_dest%>" method="post" id="fORC" name="fORC" onsubmit="if (!fORCConcluir(fORC)) return false">
	<input type="hidden" name="cliente_selecionado" id="cliente_selecionado" value='<%=id_cliente%>'>
	<input type="hidden" name="operacao_selecionada" id="operacao_selecionada" value='<%=OP_INCLUI%>'>
	<input type="hidden" name="c_FormFieldValues" id="c_FormFieldValues" value="" />

<!-- ************   ENDERE�O DE ENTREGA: S/N   ************ -->
<br>
<table width="649" class="Q" cellspacing="0">
	<tr>
		<td align="left">
		<p class="R">ENDERE�O DE ENTREGA</p><p class="C">
			<% intIdx = 0 %>
			<input type="radio" id="rb_end_entrega_nao" name="rb_end_entrega" value="N" onclick="Disabled_True(fORC);"><span class="C" style="cursor:default" onclick="fORC.rb_end_entrega[<%=Cstr(intIdx)%>].click();Disabled_True(fORC);">O mesmo endere�o do cadastro</span>
			<% intIdx = intIdx + 1 %>
			<br><input type="radio" id="rb_end_entrega_sim" name="rb_end_entrega" value="S" onclick="Disabled_False(fORC);"><span class="C" style="cursor:default" onclick="fORC.rb_end_entrega[<%=Cstr(intIdx)%>].click();Disabled_False(fORC);">Outro endere�o</span>
		</p>
		</td>
	</tr>
</table>

<!-- ************   ENDERE�O DE ENTREGA: ENDERE�O   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
	<td width="100%" align="left"><p class="R">ENDERE�O</p><p class="C">
		<input id="EndEtg_endereco" name="EndEtg_endereco" class="TA" value="" maxlength="60" style="width:635px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fORC.EndEtg_endereco_numero.focus(); filtra_nome_identificador();"></p></td>
	</tr>
</table>

<!-- ************   ENDERE�O DE ENTREGA: N�/COMPLEMENTO   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
	<td class="MD" width="50%" align="left"><p class="R">N�</p><p class="C">
		<input id="EndEtg_endereco_numero" name="EndEtg_endereco_numero" class="TA" value="" maxlength="20" style="width:310px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fORC.EndEtg_endereco_complemento.focus(); filtra_nome_identificador();"></p></td>
	<td width="50%" align="left"><p class="R">COMPLEMENTO</p><p class="C">
		<input id="EndEtg_endereco_complemento" name="EndEtg_endereco_complemento" class="TA" value="" maxlength="60" style="width:310px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fORC.EndEtg_bairro.focus(); filtra_nome_identificador();"></p></td>
	</tr>
</table>

<!-- ************   ENDERE�O DE ENTREGA: BAIRRO/CIDADE   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
	<td class="MD" width="50%" align="left"><p class="R">BAIRRO</p><p class="C">
		<input id="EndEtg_bairro" name="EndEtg_bairro" class="TA" value="" maxlength="72" style="width:310px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fORC.EndEtg_cidade.focus(); filtra_nome_identificador();"></p></td>
	<td width="50%" align="left"><p class="R">CIDADE</p><p class="C">
		<input id="EndEtg_cidade" name="EndEtg_cidade" class="TA" value="" maxlength="60" style="width:310px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fORC.EndEtg_uf.focus(); filtra_nome_identificador();"></p></td>
	</tr>
</table>

<!-- ************   ENDERE�O DE ENTREGA: UF/CEP   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
	<td width="50%" class="MD" align="left"><p class="R">UF</p><p class="C">
		<input id="EndEtg_uf" name="EndEtg_uf" class="TA" value="" maxlength="2" size="3" onkeypress="if (digitou_enter(true) && tem_info(this.value) && uf_ok(this.value)) fORC.EndEtg_cep.focus();" onblur="this.value=trim(this.value); if (!uf_ok(this.value)) {alert('UF inv�lida!!');this.focus();} else this.value=ucase(this.value);"></p></td>
	<td width="50%" align="left">
		<table width="100%" cellpadding="0" cellspacing="0">
		<tr>
			<td width="50%" align="left"><p class="R">CEP</p><p class="C">
				<input id="EndEtg_cep" name="EndEtg_cep" readonly tabindex=-1 class="TA" value="" maxlength="9" size="11" onkeypress="if (digitou_enter(true) && cep_ok(this.value)) filtra_cep();" onblur="if (!cep_ok(this.value)) {alert('CEP inv�lido!!');this.focus();} else this.value=cep_formata(this.value);"></p></td>
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
<!-- ************   JUSTIFIQUE O ENDERE�O   ************ -->
<table id="obs_endereco" width="649" class="QS" cellspacing="0">
	<tr >
	<td class="M" width="50%" align="left"><p class="R">JUSTIFIQUE O ENDERE�O</p><p class="C">
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
	<td align="left"><a href="javascript:history.back();" title="volta para a p�gina anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>

<% if (operacao_selecionada = OP_CONSULTA) Or blnEdicaoBloqueada then %>
	<td align="center"><div name="dATUALIZACONTRIB" id="dATUALIZACONTRIB">
		<a name="bATUALIZA" id="bATUALIZACONTRIB" href="javascript:AtualizaClienteContrib(fCAD)" title="atualiza o cadastro deste cliente">
		<img src="../botao/confirmar.gif" width="176" height="55" border="0"></a></div>
	</td>
	<td align="right"><div name="dORCAMENTO" id="dORCAMENTO">
		<a name="bORCAMENTO" id="bORCAMENTO" href="javascript:fORCConcluir(fORC);" title="cadastra um novo pr�-pedido para este cliente">
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