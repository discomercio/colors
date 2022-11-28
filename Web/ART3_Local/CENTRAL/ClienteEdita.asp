<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/Global.asp"    -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

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
	dim s, s_aux, usuario, pagina_retorno, strDisabled
	usuario = trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 
	
'	ESTÁ DEFINIDA A PÁGINA QUE DEVE SER EXIBIDA APÓS A ATUALIZAÇÃO NO CADASTRO?
	pagina_retorno = trim(request("pagina_retorno"))


'	CONECTA COM O BANCO DE DADOS
	dim cn,rs,tRefBancaria,tRefComercial,tRefProfissional
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	dim intIdx
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	EDIÇÃO BLOQUEADA?
	dim edicao_bloqueada, blnEdicaoBloqueada
	edicao_bloqueada = ucase(trim(request("edicao_bloqueada")))
	blnEdicaoBloqueada = False
	if edicao_bloqueada = "S" then blnEdicaoBloqueada = True
	if Not operacao_permitida(OP_CEN_EDITA_CLIENTE_DADOS_CADASTRAIS, s_lista_operacoes_permitidas) then blnEdicaoBloqueada = True

'	CLIENTE A EDITAR
	Dim id_cliente
	id_cliente = trim(request("cliente_selecionado"))
	if id_cliente = "" then Response.Redirect("aviso.asp?id=" & ERR_ID_INVALIDO)

	s = "select * from t_CLIENTE where (id='" & id_cliente & "')"
	set rs = cn.Execute(s)
	if Err <> 0 then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
	if rs.EOF then Response.Redirect("aviso.asp?id=" & ERR_CLIENTE_NAO_CADASTRADO)

	dim eh_cpf
	s=Trim("" & rs("cnpj_cpf"))
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

'	INDICADOR
	dim blnCampoIndicadorEditavel
	blnCampoIndicadorEditavel = False
	if operacao_permitida(OP_CEN_EDITA_CLIENTE_CAMPO_INDICADOR, s_lista_operacoes_permitidas) then blnCampoIndicadorEditavel = True
	if Not operacao_permitida(OP_CEN_EDITA_CLIENTE_DADOS_CADASTRAIS, s_lista_operacoes_permitidas) then blnCampoIndicadorEditavel = False
	




' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________

' ___________________________________
' L I S T A _ M I D I A
'
function lista_midia(byval id_default)
dim x,r,s,ha_default
	id_default = Trim("" & id_default)
	ha_default=False
	set r = cn.Execute("SELECT * FROM t_MIDIA WHERE indisponivel=0 ORDER BY apelido")
	s= ""
	do while Not r.eof 
		x = Trim("" & r("id"))
		if (id_default<>"") And (id_default=x) then
			s = s & "<OPTION SELECTED"
			ha_default=True
		else
			s = s & "<OPTION"
			end if
		s = s & " VALUE='" & x & "'>"
		s = s & Trim("" & r("apelido"))
		s = s & "</OPTION>" & chr(13)
		r.MoveNext
		loop

	if Not ha_default then
		s = "<OPTION SELECTED VALUE=''>&nbsp;</OPTION>" & chr(13) & s
		end if
	
	lista_midia = s
	r.close
	set r=nothing
	
end function
%>



<%=DOCTYPE_LEGADO%>

<html>


<head>
	<title>CENTRAL</title>
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

	if ((!eh_cpf) || (eh_cpf && f.rb_produtor_rural[1].checked)) {
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
		if (fCAD.rb_contribuinte_icms[2].checked) {
			if (fCAD.ie.value != "") {
				alert("Se o Contribuinte ICMS é isento, o campo IE deve ser vazio!");
				fCAD.ie.focus();
				return;
			}
		}
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
		}
	else {
		//deixar de exigir preenchimento se cliente não é contribuinte?
		//s = trim(f.ie.value);
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

		if (eh_cpf) {
		    if ((trim(f.tel_res.value) == "") && (trim(f.tel_com.value) == "") && (trim(f.tel_cel.value) == "")) {
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

/*	if (trim(f.midia.options[f.midia.selectedIndex].value)=="") {
		alert('Indique a forma pela qual conheceu a DIS!!');
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
/*
	if (!eh_cpf) blnCadSocioMaj=true; else blnCadSocioMaj=false;
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
</script>

<script type="text/javascript">
	function exibeJanelaCEP_Cli() {
		$.mostraJanelaCEP("cep", "uf", "cidade", "bairro", "endereco", "endereco_numero", "endereco_complemento");
	}

	function trataProdutorRural() {
		//ao clicar na opção Produtor Rural, exibir/ocultar os campos apropriados
		if (!fCAD.rb_produtor_rural[1].checked) {
			$("#t_contribuinte_icms").css("display", "none");
		}
		else {
			$("#t_contribuinte_icms").css("display", "block");
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

<link href="<%=URL_FILE__E_CSS%>" Rel="stylesheet" Type="text/css">
<link href="<%=URL_FILE__EPRINTER_CSS%>" Rel="stylesheet" Type="text/css" media="print">
<link href="<%=URL_FILE__E_JANELABUSCACEP_CSS%>" rel="stylesheet" type="text/css">

<body id="corpoPagina" onload="focus();ja_carregou=true;trataProdutorRural();">
<center>

<!-- #include file = "../global/JanelaBuscaCEP.htm"    -->


<!--  CADASTRO DO CLIENTE -->

<table width="698" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="center" valign="bottom"><p class="PEDIDO">Consulta/Edição de Cliente Cadastrado<br><span class="C">&nbsp;</span></p></td>
</tr>
</table>
<br>

<!-- ************   EXIBE OBSERVAÇÕES CREDITÍCIAS?  ************ -->
 <%	s = Trim("" & rs("obs_crediticias"))
	if s <> "" then %>
		<span class="Lbl" style="display:none">OBSERVAÇÕES CREDITÍCIAS</span>
		<div class='MtAviso' style="width:649px;FONT-WEIGHT:bold;border:1pt solid black;display:none;" align="CENTER"><P style='margin:5px 2px 5px 2px;'><%=s%></p></div>
		<br>
	<% end if %>



<!-- ************  CAMPOS DO CADASTRO  ************ -->
<form id="fCAD" name="fCAD" METHOD="POST" ACTION="ClienteAtualiza.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name='cliente_selecionado' id="cliente_selecionado" value='<%=id_cliente%>'>
<input type="hidden" name='pagina_retorno' id="pagina_retorno" value='<%=pagina_retorno%>'>

<%if blnCampoIndicadorEditavel then%>
<input type="hidden" name='CampoIndicadorEditavel' id="CampoIndicadorEditavel" value='S'>
<%else%>
<input type="hidden" name='CampoIndicadorEditavel' id="CampoIndicadorEditavel" value='N'>
<%end if%>

<INPUT type="hidden" name='contribuinte_icms_cadastrado' id="contribuinte_icms_cadastrado" value='<%=Trim("" & rs("contribuinte_icms_status"))%>'>
<INPUT type="hidden" name='produtor_rural_cadastrado' id="produtor_rural_cadastrado" value='<%=Trim("" & rs("produtor_rural_status"))%>'>
<input type="hidden" name="c_FormFieldValues" id="c_FormFieldValues" value="" />


<!-- ************   CNPJ/IE OU CPF/RG/NASCIMENTO/SEXO  ************ -->
<table width="649" class="Q" cellspacing="0">
	<tr>
		<td width="210" align="left">
	<%if eh_cpf then s="CPF" else s="CNPJ"%>
	<p class="R"><%=s%></p><p class="C">
	<%	s=Trim("" & rs("cnpj_cpf"))
		s=cnpj_cpf_formata(s)
	%>
	<input id="cnpj_cpf_selecionado" name="cnpj_cpf_selecionado" class="TA" value="<%=s%>" readonly size="22" style="text-align:center; color:#0000ff"></p></td>

<%if eh_cpf then%>
	<td class="MDE" width="210" align="left"><p class="R">RG</p><p class="C">
		<input id="rg" name="rg" class="TA" type="text" maxlength="20" size="22" value="<%=Trim("" & rs("rg"))%>" onkeypress="if (digitou_enter(true)) fCAD.dt_nasc.focus(); filtra_nome_identificador();"></p></td>
	<td class="MD" align="left"><p class="R">NASCIMENTO</p><p class="C">
		<input id="dt_nasc" name="dt_nasc" class="TA" type="text" maxlength="10" size="14" value="<%=formata_data(rs("dt_nasc"))%>" onkeypress="if (digitou_enter(true) && isDate(this)) fCAD.sexo.focus(); filtra_data();" onblur="if (tem_info(this.value)) if (!isDate(this)) {alert('Data inválida!!');this.focus();}"></p></td>
	<td align="left"><p class="R">SEXO</p><p class="C">
		<input id="sexo" name="sexo" class="TA" type="text" maxlength="1" size="2" value="<%=Trim("" & rs("sexo"))%>" onkeypress="if (digitou_enter(true)) if (!tem_info(this.value)) fCAD.nome.focus(); else if (sexo_ok(this.value)) fCAD.nome.focus(); filtra_sexo();" onkeyup="this.value=ucase(this.value);"></p></td>

<%else%>
	<td class="MDE" width="215" align="left"><p class="R">IE</p><p class="C">
		<input id="ie" name="ie" class="TA" type="text" maxlength="20" size="25" value="<%=Trim("" & rs("ie"))%>" onkeypress="if (digitou_enter(true)) fCAD.nome.focus(); filtra_nome_identificador();"></p></td>
	<td align="left"><p class="R">CONTRIBUINTE ICMS</p><p class="C">
		<%s=Trim("" & rs("contribuinte_icms_status"))%>
		<%if s=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO then s_aux="checked" else s_aux=""%>
		<% intIdx = 0 %>
		<input type="radio" id="rb_contribuinte_icms_nao" name="rb_contribuinte_icms" value="<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO%>" <%=s_aux%>><span class="C" style="cursor:default" onclick="fNEW.rb_contribuinte_icms[<%=Cstr(intIdx)%>].click();">Não</span>
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
		<%s=Trim("" & rs("produtor_rural_status"))%>
		<%if s=COD_ST_CLIENTE_PRODUTOR_RURAL_NAO then s_aux="checked" else s_aux=""%>
		<% intIdx = 0 %>
		<input type="radio" id="rb_produtor_rural_nao" name="rb_produtor_rural" value="<%=COD_ST_CLIENTE_PRODUTOR_RURAL_NAO%>" <%=s_aux%> onclick="trataProdutorRural();"><span class="C" style="cursor:default" onclick="fNEW.rb_produtor_rural[<%=Cstr(intIdx)%>].click();">Não</span>
		<%if s=COD_ST_CLIENTE_PRODUTOR_RURAL_SIM then s_aux="checked" else s_aux=""%>
		<% intIdx = intIdx + 1 %>
		<input type="radio" id="rb_produtor_rural_sim" name="rb_produtor_rural" value="<%=COD_ST_CLIENTE_PRODUTOR_RURAL_SIM%>" <%=s_aux%> onclick="trataProdutorRural();"><span class="C" style="cursor:default" onclick="fNEW.rb_produtor_rural[<%=Cstr(intIdx)%>].click();">Sim</span></p></td>
	</tr>
</table>

<table width="649" class="QS" cellspacing="0" id="t_contribuinte_icms" onload="trataProdutorRural();">
	<tr>
	<td width="210" class="MD" align="left"><p class="R">IE</p><p class="C">
		<input id="ie" name="ie" class="TA" type="text" maxlength="20" size="25" value="<%=Trim("" & rs("ie"))%>" onkeypress="if (digitou_enter(true)) fCAD.nome.focus(); filtra_nome_identificador();"></p></td>

	<td align="left"><p class="R">CONTRIBUINTE ICMS</p><p class="C">
		<%s=Trim("" & rs("contribuinte_icms_status"))%>
		<%if s=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO then s_aux="checked" else s_aux=""%>
		<% intIdx = 0 %>
		<input type="radio" id="rb_contribuinte_icms_nao" name="rb_contribuinte_icms" value="<%=COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO%>" <%=s_aux%>><span class="C" style="cursor:default" onclick="fNEW.rb_contribuinte_icms[<%=Cstr(intIdx)%>].click();">Não</span>
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
	<%if eh_cpf then s="NOME" else s="RAZÃO SOCIAL"%>
	<td width="100%" align="left"><p class="R"><%=s%></p><p class="C">
		<input id="nome" name="nome" class="TA" value="<%=Trim("" & rs("nome"))%>" maxlength="60" size="85" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCAD.endereco.focus(); filtra_nome_identificador();"></p></td>
	</tr>
</table>

<!-- ************   ENDEREÇO   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
	<td width="100%" align="left"><p class="R">ENDEREÇO</p><p class="C">
		<input id="endereco" name="endereco" class="TA" value="<%=Trim("" & rs("endereco"))%>" maxlength="60" style="width:635px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCAD.endereco_numero.focus(); filtra_nome_identificador();"></p></td>
	</tr>
</table>

<!-- ************   Nº/COMPLEMENTO   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
	<td class="MD" width="50%" align="left"><p class="R">Nº</p><p class="C">
		<input id="endereco_numero" name="endereco_numero" class="TA" value="<%=Trim("" & rs("endereco_numero"))%>" maxlength="20" style="width:310px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCAD.endereco_complemento.focus(); filtra_nome_identificador();"></p></td>
	<td align="left"><p class="R">COMPLEMENTO</p><p class="C">
		<input id="endereco_complemento" name="endereco_complemento" class="TA" value="<%=Trim("" & rs("endereco_complemento"))%>" maxlength="60" style="width:310px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCAD.bairro.focus(); filtra_nome_identificador();"></p></td>
	</tr>
</table>

<!-- ************   BAIRRO/CIDADE   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
	<td class="MD" width="50%" align="left"><p class="R">BAIRRO</p><p class="C">
		<input id="bairro" name="bairro" class="TA" value="<%=Trim("" & rs("bairro"))%>" maxlength="72" style="width:310px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCAD.cidade.focus(); filtra_nome_identificador();"></p></td>
	<td align="left"><p class="R">CIDADE</p><p class="C">
		<input id="cidade" name="cidade" class="TA" value="<%=Trim("" & rs("cidade"))%>" maxlength="60" style="width:310px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCAD.uf.focus(); filtra_nome_identificador();"></p></td>
	</tr>
</table>

<!-- ************   UF/CEP   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
	<td class="MD" width="50%" align="left"><p class="R">UF</p><p class="C">
		<input id="uf" name="uf" class="TA" value="<%=Trim("" & rs("uf"))%>" maxlength="2" size="3" onkeypress="if (digitou_enter(true) && tem_info(this.value) && uf_ok(this.value)) 
			<%if eh_cpf then Response.write "fCAD.ddd_res.focus();" else Response.Write "fCAD.ddd_com.focus();"%>" 
			onblur="this.value=trim(this.value); if (!uf_ok(this.value)) {alert('UF inválida!!');this.focus();} else this.value=ucase(this.value);"></p></td>
	<td width="50%" align="left">
		<table width="100%" cellpadding="0" cellspacing="0">
		<tr>
			<td width="50%" align="left"><p class="R">CEP</p><p class="C">
				<input id="cep" name="cep" class="TA" readonly tabindex=-1 value="<%=cep_formata(Trim("" & rs("cep")))%>" maxlength="9" size="11" onkeypress="if (digitou_enter(true) && cep_ok(this.value)) 
					<%if eh_cpf then Response.write "fCAD.ddd_res.focus();" else Response.Write "fCAD.ddd_com.focus();"%> 
					filtra_cep();" onblur="if (!cep_ok(this.value)) {alert('CEP inválido!!');this.focus();} else this.value=cep_formata(this.value);"></p></td>
			<td align="center" width="50%">
				<% if blnPesquisaCEPAntiga then %>
				<button type="button" name="bPesqCep" id="bPesqCep" style='width:130px;font-size:10pt;' class="Botao" onclick="AbrePesquisaCep();">Pesquisar CEP</button>
				<% end if %>
				<% if blnPesquisaCEPAntiga and blnPesquisaCEPNova then Response.Write "&nbsp;" %>
				<% if blnPesquisaCEPNova then %>
				<button type="button" name="bPesqCep" id="bPesqCep" style='width:130px;font-size:10pt;' class="Botao" onclick="exibeJanelaCEP_Cli();">Pesquisar CEP</button>
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
		<input id="ddd_res" name="ddd_res" class="TA" value="<%=Trim("" & rs("ddd_res"))%>" maxlength="4" size="5" onkeypress="if (digitou_enter(true) && ddd_ok(this.value)) fCAD.tel_res.focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inválido!!');this.focus();}"></p></td>
	<td align="left"><p class="R">TELEFONE RESIDENCIAL</p><p class="C">
		<input id="tel_res" name="tel_res" class="TA" value="<%=telefone_formata(Trim("" & rs("tel_res")))%>" maxlength="11" size="12" onkeypress="if (digitou_enter(true) && telefone_ok(this.value)) fCAD.ddd_cel.focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('Telefone inválido!!');this.focus();} else this.value=telefone_formata(this.value);"></p></td>
	</tr>
	<tr>
	<td class="MD MC" width="20%" align="left"><p class="R">DDD</p><p class="C">
		<input id="ddd_cel" name="ddd_cel" class="TA" value="<%=Trim("" & rs("ddd_cel"))%>" maxlength="4" size="5" onkeypress="if (digitou_enter(true) && ddd_ok(this.value)) fCAD.tel_cel.focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inválido!!');this.focus();}"></p></td>
	<td align="left" class="MC"><p class="R">CELULAR</p><p class="C">
		<input id="tel_cel" name="tel_cel" class="TA" value="<%=telefone_formata(Trim("" & rs("tel_cel")))%>" maxlength="10" size="12" onkeypress="if (digitou_enter(true) && telefone_ok(this.value)) fCAD.ddd_com.focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('Número de celular inválido!!');this.focus();} else this.value=telefone_formata(this.value);"></p></td>
	</tr>
</table>
<% end if %>
	
<!-- ************   TELEFONE COMERCIAL   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
	<td class="MD" width="20%" align="left"><p class="R">DDD</p><p class="C">
		<input id="ddd_com" name="ddd_com" class="TA" value="<%=Trim("" & rs("ddd_com"))%>" maxlength="4" size="5" onkeypress="if (digitou_enter(true) && ddd_ok(this.value)) fCAD.tel_com.focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inválido!!');this.focus();}"></p></td>
	<%if eh_cpf then s=" COMERCIAL" else s=""%>
	<td class="MD" align="left"><p class="R">TELEFONE<%=s%></p><p class="C">
		<input id="tel_com" name="tel_com" class="TA" value="<%=telefone_formata(Trim("" & rs("tel_com")))%>" maxlength="11" size="12" onkeypress="if (digitou_enter(true) && telefone_ok(this.value)) fCAD.ramal_com.focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('Telefone inválido!!');this.focus();} else this.value=telefone_formata(this.value);"></p></td>
	<td align="left"><p class="R">RAMAL</p><p class="C">
		<input id="ramal_com" name="ramal_com" class="TA" value="<%=Trim("" & rs("ramal_com"))%>" maxlength="4" size="6" onkeypress="if (digitou_enter(true))
			 <%if Not eh_cpf then Response.Write "fCAD.ddd_com_2.focus();" else Response.Write "filiacao.focus();" %> filtra_numerico();"></p></td>
	</tr>
	<% if Not eh_cpf then %>
	<tr>
	    <td class="MD MC" width="20%" align="left"><p class="R">DDD</p><p class="C">
	    <%s=Trim("" & rs("ddd_com_2"))%>
	    <input id="ddd_com_2" name="ddd_com_2" class="TA" value="<%=s%>" maxlength="4" size="5" onkeypress="if (digitou_enter(true) && ddd_ok(this.value)) fCAD.tel_com_2.focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inválido!!!');this.focus();}" /></p>  
	    </td>
	    <td class="MD MC" align="left"><p class="R">TELEFONE</p><p class="C">
	    <%s=Trim("" & rs("tel_com_2"))%>
	    <input id="tel_com_2" name="tel_com_2" class="TA" value="<%=telefone_formata(s)%>" maxlength="9" size="12" onkeypress="if (digitou_enter(true) && telefone_ok(this.value)) fCAD.ramal_com_2.focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('Telefone inválido!!');this.focus();} else this.value=telefone_formata(this.value);"></p>
	    </td>
	    <td align="left" class="MC"><p class="R">RAMAL</p><p class="C">
	    <%s=Trim("" & rs("ramal_com_2"))%>
	    <input id="ramal_com_2" name="ramal_com_2" class="TA" value="<%=s%>" maxlength="4" size="6" onkeypress="if (digitou_enter(true)) <%if eh_cpf then Response.Write "fCAD.filiacao.focus();" else Response.Write "fCAD.contato.focus();"%> filtra_numerico();" /></p>
	    </td>
	</tr>
	<% end if %>
</table>

<% if eh_cpf then %>
<!-- ************   OBSERVAÇÃO (ANTIGO CAMPO FILIAÇÃO)   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
	<td width="100%" align="left"><p class="R">OBSERVAÇÃO</p><p class="C">
		<input id="filiacao" name="filiacao" class="TA" value="<%=Trim("" & rs("filiacao"))%>" maxlength="60" size="74" onkeypress="if (digitou_enter(true)) fCAD.email.focus(); filtra_nome_identificador();"></p></td>
	</tr>
</table>
<% else %>
<!-- ************   CONTATO   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
	<td width="100%" align="left"><p class="R">NOME DA PESSOA PARA CONTATO NA EMPRESA</p><p class="C">
		<input id="contato" name="contato" class="TA" value="<%=Trim("" & rs("contato"))%>" maxlength="30" size="45" onkeypress="if (digitou_enter(true)) fCAD.email.focus(); filtra_nome_identificador();"></p></td>
	</tr>
</table>
<% end if %>

<!-- ************   E-MAIL   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
	<td width="100%" align="left"><p class="R">E-MAIL</p><p class="C">
		<input id="email" name="email" class="TA" value="<%=Trim("" & rs("email"))%>" maxlength="60" size="74" onkeypress="if (digitou_enter(true)) fCAD.email_xml.focus(); filtra_email();"></p></td>
	</tr>
</table>

<!-- ************   E-MAIL (XML)  ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
	<td width="100%" align="left"><p class="R">E-MAIL (XML)</p><p class="C">
		<input id="email_xml" name="email_xml" class="TA" value="<%=Trim("" & rs("email_xml"))%>" maxlength="60" size="74" onkeypress="if (digitou_enter(true)) fCAD.obs_crediticias.focus(); filtra_email();"></p></td>
	</tr>
</table>

<!-- ************   OBS CREDITÍCIAS (INATIVO)  ************ -->
<table width="649" class="QS" cellspacing="0" style="display:none">
	<tr>
	<td width="100%" align="left"><p class="R">OBSERVAÇÕES CREDITÍCIAS</p><p class="C">
		<input id="obs_crediticias" name="obs_crediticias" class="TA" value="<%=Trim("" & rs("obs_crediticias"))%>" maxlength="50" size="65" onkeypress="if (digitou_enter(true)) fCAD.midia.focus(); filtra_nome_identificador();"></p></td>
	</tr>
</table>

<!-- ************   MÍDIA (INATIVO)  ************ -->
<table width="649" class="QS" cellspacing="0" style="display:none">
	<tr>
	<% if rs("spc_negativado_status") = 1 then %>
	<td class="MD" width="50%" align="left"><p class="R">FORMA PELA QUAL CONHECEU A DIS</p><p class="C">
		<select id="Select1" name="midia" style="margin-top:4pt; margin-bottom:4pt;">
			<%=lista_midia(Trim("" & rs("midia")))%>
		</select></p>
	</td>
	<td width="50%" align="left" valign="top"><p class="R">SPC</p><p class="C">
			<input id="infoSPC" name="infoSPC" class="TA" style="color: #FF0000; margin-top:4pt; margin-bottom:4pt;" value="Cliente Negativado (em <%=formata_data(rs("spc_negativado_data"))%>)" size="30"></p>
	</td>
	<% else %>
	<td width="100%" align="left"><p class="R">FORMA PELA QUAL CONHECEU A DIS</p><p class="C">
		<select id="midia" name="midia" style="margin-top:4pt; margin-bottom:4pt;">
			<%=lista_midia(Trim("" & rs("midia")))%>
		</select>
	</td>
	<% end if %>
	</tr>
</table>

<!-- ************   SPC   ************ -->
<% if rs("spc_negativado_status") = 1 then %>
<table width="649" class="QS" cellspacing="0">
    <tr>
        <td width="50%" align="left" valign="top"><p class="R">SPC</p><p class="C">
			        <input id="Text1" name="infoSPC" class="TA" style="color: #FF0000; margin-top:4pt; margin-bottom:4pt;" value="Cliente Negativado (em <%=formata_data(rs("spc_negativado_data"))%>)" size="30"></p>
        </td>
    </tr>
</table>
<% end if %>

<!-- ************   INDICADOR   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
	<td width="100%" align="left"><p class="R">INDICADOR</p><p class="C">
		<%	strDisabled = ""
			if Not blnCampoIndicadorEditavel then strDisabled = " DISABLED"
		%>
		<select id="indicador" name="indicador" <%=strDisabled%> style="margin-top:4pt; margin-bottom:4pt;">
			<%=indicadores_monta_itens_select(Trim("" & rs("indicador")))%>
		</select>
	</tr>
</table>


<!-- ************   REF BANCÁRIA   ************ -->
<%if blnCadRefBancaria then%>
<!--  ASSEGURA CRIAÇÃO DE UM ARRAY DE CAMPOS, MESMO QUANDO HOUVER SOMENTE 1 LINHA!! -->
<input type=HIDDEN name="c_RefBancariaBanco" id="c_RefBancariaBanco" value="">
<input type=HIDDEN name="c_RefBancariaAgencia" id="c_RefBancariaAgencia" value="">
<input type=HIDDEN name="c_RefBancariaConta" id="c_RefBancariaConta" value="">
<input type=HIDDEN name="c_RefBancariaDdd" id="c_RefBancariaDdd" value="">
<input type=HIDDEN name="c_RefBancariaTelefone" id="c_RefBancariaTelefone" value="">
<input type=HIDDEN name="c_RefBancariaContato" id="c_RefBancariaContato" value="">
	<% 
		s="SELECT * FROM t_CLIENTE_REF_BANCARIA WHERE (id_cliente='" & Trim("" & rs("id")) & "') ORDER BY ordem"
		set tRefBancaria = cn.Execute(s)
	%>
	
	<% for intCounter=1 to int_MAX_REF_BANCARIA_CLIENTE %>
		<%
		strRefBancariaBanco=""
		strRefBancariaAgencia=""
		strRefBancariaConta=""
		strRefBancariaDdd=""
		strRefBancariaTelefone=""
		strRefBancariaContato=""
		if Not tRefBancaria.Eof then 
			strRefBancariaBanco=Trim("" & tRefBancaria("banco"))
			strRefBancariaAgencia=Trim("" & tRefBancaria("agencia"))
			strRefBancariaConta=Trim("" & tRefBancaria("conta"))
			strRefBancariaDdd=Trim("" & tRefBancaria("ddd"))
			strRefBancariaTelefone=Trim("" & tRefBancaria("telefone"))
			strRefBancariaContato=Trim("" & tRefBancaria("contato"))
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
							    <select name="c_RefBancariaBanco" id="c_RefBancariaBanco" style="margin:1px 10px 6px 10px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
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
							    <input name="c_RefBancariaAgencia" id="c_RefBancariaAgencia" class="TA" maxlength="8" size="12" value="<%=strRefBancariaAgencia%>" onkeypress="if (digitou_enter(true)) fCAD.c_RefBancariaConta[<%=CStr(intCounter)%>].focus(); filtra_agencia_bancaria();">
						    </p>
					    </td>
					    <td class="MD" align="left">
						    <p class="R">CONTA</p>
						    <p class="C">
							    <input name="c_RefBancariaConta" id="c_RefBancariaConta" class="TA" maxlength="12" value="<%=strRefBancariaConta%>" onkeypress="if (digitou_enter(true)) fCAD.c_RefBancariaDdd[<%=CStr(intCounter)%>].focus(); filtra_conta_bancaria();">
						    </p>
					    </td>
					    <td class="MD" align="left">
						    <p class="R">DDD</p>
						    <p class="C">
							    <input name="c_RefBancariaDdd" id="c_RefBancariaDdd" class="TA" maxlength="2" size="4" value="<%=strRefBancariaDdd%>" onkeypress="if (digitou_enter(true)) fCAD.c_RefBancariaTelefone[<%=CStr(intCounter)%>].focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inválido!!');this.focus();}">
						    </p>
					    </td>
					    <td align="left">
						    <p class="R">TELEFONE</p>
						    <p class="C">
							    <input name="c_RefBancariaTelefone" id="c_RefBancariaTelefone" class="TA" maxlength="9" value="<%=telefone_formata(strRefBancariaTelefone)%>" onkeypress="if (digitou_enter(true)) fCAD.c_RefBancariaContato[<%=CStr(intCounter)%>].focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('Telefone inválido!!');this.focus();} else this.value=telefone_formata(this.value);">
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
							    <input name="c_RefBancariaContato" id="c_RefBancariaContato" class="TA" maxlength="40"  style="width:600px;" value="<%=strRefBancariaContato%>" onkeypress="if (digitou_enter(true)) {if (<%=Cstr(intCounter+1)%>==fCAD.c_RefBancariaAgencia.length) this.focus(); else fCAD.c_RefBancariaAgencia[<%=Cstr(intCounter+1)%>].focus();} filtra_nome_identificador();">
						    </p>
					    </td>
				    </tr>
			    </table>
		    </td>
	    </tr>
    </table>
<% end if %>
		<% 
			if Not tRefBancaria.Eof then tRefBancaria.MoveNext
		%>
		
	<% next %>
	
	<% 
		tRefBancaria.Close
	%>
<%end if%>


<!-- ************   REF PROFISSIONAL   ************ -->
<%if blnCadRefProfissional then%>
<input type=HIDDEN name="c_RefProfNomeEmpresa" id="c_RefProfNomeEmpresa" value="">
<input type=HIDDEN name="c_RefProfCargo" id="c_RefProfCargo" value="">
<input type=HIDDEN name="c_RefProfDdd" id="c_RefProfDdd" value="">
<input type=HIDDEN name="c_RefProfTelefone" id="c_RefProfTelefone" value="">
<input type=HIDDEN name="c_RefProfPeriodoRegistro" id="c_RefProfPeriodoRegistro" value="">
<input type=HIDDEN name="c_RefProfRendimentos" id="c_RefProfRendimentos" value="">
<input type=HIDDEN name="c_RefProfCnpj" id="c_RefProfCnpj" value="">
	<% 
		s="SELECT * FROM t_CLIENTE_REF_PROFISSIONAL WHERE (id_cliente='" & Trim("" & rs("id")) & "') ORDER BY ordem"
		set tRefProfissional = cn.Execute(s)
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
		if Not tRefProfissional.Eof then 
			strRefProfNomeEmpresa=Trim("" & tRefProfissional("nome_empresa"))
			strRefProfCargo=Trim("" & tRefProfissional("cargo"))
			strRefProfDdd=Trim("" & tRefProfissional("ddd"))
			strRefProfTelefone=Trim("" & tRefProfissional("telefone"))
			strRefProfPeriodoRegistro=formata_data(tRefProfissional("periodo_registro"))
			strRefProfRendimentos=formata_moeda(tRefProfissional("rendimentos"))
			strRefProfCnpj=cnpj_cpf_formata(Trim("" & tRefProfissional("cnpj")))
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
							<input name="c_RefProfNomeEmpresa" id="c_RefProfNomeEmpresa" class="TA" maxlength="60"  style="width:450px;" value="<%=strRefProfNomeEmpresa%>" onkeypress="if (digitou_enter(true)) fCAD.c_RefProfCnpj[<%=CStr(intCounter)%>].focus(); filtra_nome_identificador();">
						</p>
					</td>
					<td class="MC" align="left">
						<p class="R">CNPJ</p>
						<p class="C">
							<input name="c_RefProfCnpj" id="c_RefProfCnpj" class="TA" maxlength="18" size="24" value="<%=strRefProfCnpj%>" onkeypress="if (digitou_enter(true) && cnpj_ok(this.value)) fCAD.c_RefProfCargo[<%=CStr(intCounter)%>].focus(); filtra_cnpj();" onblur="if (!cnpj_ok(this.value)) {alert('CNPJ inválido'); this.focus();} else this.value=cnpj_formata(this.value);">
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
							<input name="c_RefProfCargo" id="c_RefProfCargo" class="TA" maxlength="40" style="width:350px;" value="<%=strRefProfCargo%>" onkeypress="if (digitou_enter(true)) fCAD.c_RefProfDdd[<%=CStr(intCounter)%>].focus(); filtra_nome_identificador();">
						</p>
					</td>
					<td class="MD" align="left">
						<p class="R">DDD</p>
						<p class="C">
							<input name="c_RefProfDdd" id="c_RefProfDdd" class="TA" maxlength="2" size=4 value="<%=strRefProfDdd%>" onkeypress="if (digitou_enter(true)) fCAD.c_RefProfTelefone[<%=CStr(intCounter)%>].focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inválido!!');this.focus();}">
						</p>
					</td>
					<td align="left">
						<p class="R">TELEFONE</p>
						<p class="C">
							<input name="c_RefProfTelefone" id="c_RefProfTelefone" class="TA" maxlength="9" value="<%=telefone_formata(strRefProfTelefone)%>" onkeypress="if (digitou_enter(true)) fCAD.c_RefProfPeriodoRegistro[<%=CStr(intCounter)%>].focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('Telefone inválido!!');this.focus();} else this.value=telefone_formata(this.value);">
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
							<input name="c_RefProfPeriodoRegistro" id="c_RefProfPeriodoRegistro" class="TA" maxlength="10" value="<%=strRefProfPeriodoRegistro%>" onkeypress="if (digitou_enter(true)) fCAD.c_RefProfRendimentos[<%=CStr(intCounter)%>].focus(); filtra_data();" onblur="if (!isDate(this)) {alert('Data inválida!!');this.focus();}">
						</p>
					</td>
					<td width="50%" align="left">
						<p class="R">RENDIMENTOS (<%=SIMBOLO_MONETARIO%>)</p>
						<p class="C">
							<input name="c_RefProfRendimentos" id="c_RefProfRendimentos" class="TA" maxlength="18" value="<%=strRefProfRendimentos%>" onkeypress="if (digitou_enter(true)) {if (<%=Cstr(intCounter+1)%>==fCAD.c_RefProfNomeEmpresa.length) this.focus(); else fCAD.c_RefProfNomeEmpresa[<%=Cstr(intCounter+1)%>].focus();} filtra_moeda_positivo();" onblur="this.value=formata_moeda(this.value);">
						</p>
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
		<% 
			if Not tRefProfissional.Eof then tRefProfissional.MoveNext
		%>
		
	<% next %>
	
	<% 
		tRefProfissional.Close
	%>
<%end if%>


<!-- ************   REF COMERCIAL   ************ -->
<%if blnCadRefComercial then%>
<!--  ASSEGURA CRIAÇÃO DE UM ARRAY DE CAMPOS, MESMO QUANDO HOUVER SOMENTE 1 LINHA!! -->
<input type=HIDDEN name="c_RefComercialNomeEmpresa" id="c_RefComercialNomeEmpresa" value="">
<input type=HIDDEN name="c_RefComercialContato" id="c_RefComercialContato" value="">
<input type=HIDDEN name="c_RefComercialDdd" id="c_RefComercialDdd" value="">
<input type=HIDDEN name="c_RefComercialTelefone" id="c_RefComercialTelefone" value="">
	<% 
		s="SELECT * FROM t_CLIENTE_REF_COMERCIAL WHERE (id_cliente='" & Trim("" & rs("id")) & "') ORDER BY ordem"
		set tRefComercial = cn.Execute(s)
	%>
	
	<% for intCounter=1 to int_MAX_REF_COMERCIAL_CLIENTE %>
		<%
		strRefComercialNomeEmpresa=""
		strRefComercialContato=""
		strRefComercialDdd=""
		strRefComercialTelefone=""
		if Not tRefComercial.Eof then 
			strRefComercialNomeEmpresa=Trim("" & tRefComercial("nome_empresa"))
			strRefComercialContato=Trim("" & tRefComercial("contato"))
			strRefComercialDdd=Trim("" & tRefComercial("ddd"))
			strRefComercialTelefone=Trim("" & tRefComercial("telefone"))
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
							<input name="c_RefComercialNomeEmpresa" id="c_RefComercialNomeEmpresa" class="TA" maxlength="60"  style="width:600px;" value="<%=strRefComercialNomeEmpresa%>" onkeypress="if (digitou_enter(true)) fCAD.c_RefComercialContato[<%=CStr(intCounter)%>].focus(); filtra_nome_identificador();">
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
							<input name="c_RefComercialContato" id="c_RefComercialContato" class="TA" maxlength="40" value="<%=strRefComercialContato%>" onkeypress="if (digitou_enter(true)) fCAD.c_RefComercialDdd[<%=CStr(intCounter)%>].focus(); filtra_nome_identificador();">
						</p>
					</td>
					<td class="MD" align="left">
						<p class="R">DDD</p>
						<p class="C">
							<input name="c_RefComercialDdd" id="c_RefComercialDdd" class="TA" maxlength="2" size="4" value="<%=strRefComercialDdd%>" onkeypress="if (digitou_enter(true)) fCAD.c_RefComercialTelefone[<%=CStr(intCounter)%>].focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inválido!!');this.focus();}">
						</p>
					</td>
					<td align="left">
						<p class="R">TELEFONE</p>
						<p class="C">
							<input name="c_RefComercialTelefone" id="c_RefComercialTelefone" class="TA" maxlength="9" value="<%=telefone_formata(strRefComercialTelefone)%>" onkeypress="if (digitou_enter(true)) {if (<%=Cstr(intCounter+1)%>==fCAD.c_RefComercialNomeEmpresa.length) this.focus(); else fCAD.c_RefComercialNomeEmpresa[<%=Cstr(intCounter+1)%>].focus();} filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('Telefone inválido!!');this.focus();} else this.value=telefone_formata(this.value);">
						</p>
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
		<% 
			if Not tRefComercial.Eof then tRefComercial.MoveNext
		%>
		
	<% next %>
	
	<% 
		tRefComercial.Close
	%>
<%end if%>


<!-- ************   PJ: DADOS DO SÓCIO MAJORITÁRIO (INATIVO)  ************ -->
<%if blnCadSocioMaj then%>
<br>
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
						<input id="c_SocioMajNome" name="c_SocioMajNome" class="TA" value="<%=Trim("" & rs("SocMaj_Nome"))%>" maxlength="60" size="61" onkeypress="if (digitou_enter(true)) fCAD.c_SocioMajCpf.focus(); filtra_nome_identificador();" onblur="this.value=trim(this.value);"></p>
					</td>
					<td class="MC" align="left"><p class="R">CPF</p><p class="C">
						<input id="c_SocioMajCpf" name="c_SocioMajCpf" class="TA" value="<%=cnpj_cpf_formata(Trim("" & rs("SocMaj_CPF")))%>" maxlength="14" size="15" onkeypress="if (digitou_enter(true) && cpf_ok(this.value)) fCAD.c_SocioMajBanco.focus(); filtra_numerico();" onblur="if (!cpf_ok(this.value)) {alert('CPF inválido!!');this.focus();} else this.value=cnpj_cpf_formata(this.value);"></p>
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
						<p class="R">BANCO</p>
						<p class="C">
							<select name="c_SocioMajBanco" id="c_SocioMajBanco" style="margin:1px 10px 6px 10px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
							<%=banco_monta_itens_select(Trim("" & rs("SocMaj_banco"))) %>
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
							<input name="c_SocioMajAgencia" id="c_SocioMajAgencia" class="TA" maxlength="8" value="<%=Trim("" & rs("SocMaj_agencia"))%>" onkeypress="if (digitou_enter(true)) fCAD.c_SocioMajConta.focus(); filtra_agencia_bancaria();">
						</p>
					</td>
					<td class="MD" align="left">
						<p class="R">CONTA</p>
						<p class="C">
							<input name="c_SocioMajConta" id="c_SocioMajConta" class="TA" maxlength="12" value="<%=Trim("" & rs("SocMaj_conta"))%>" onkeypress="if (digitou_enter(true)) fCAD.c_SocioMajDdd.focus(); filtra_conta_bancaria();">
						</p>
					</td>
					<td class="MD" align="left">
						<p class="R">DDD</p>
						<p class="C">
							<input name="c_SocioMajDdd" id="c_SocioMajDdd" class="TA" maxlength="2" size="4" value="<%=Trim("" & rs("SocMaj_ddd"))%>" onkeypress="if (digitou_enter(true)) fCAD.c_SocioMajTelefone.focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inválido!!');this.focus();}">
						</p>
					</td>
					<td align="left">
						<p class="R">TELEFONE</p>
						<p class="C">
							<input name="c_SocioMajTelefone" id="c_SocioMajTelefone" class="TA" maxlength="9" value="<%=telefone_formata(Trim("" & rs("SocMaj_telefone")))%>" onkeypress="if (digitou_enter(true)) fCAD.c_SocioMajContato.focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('Telefone inválido!!');this.focus();} else this.value=telefone_formata(this.value);">
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
							<input name="c_SocioMajContato" id="c_SocioMajContato" class="TA" maxlength="40"  style="width:600px;" value="<%=Trim("" & rs("SocMaj_contato"))%>" onkeypress="filtra_nome_identificador();">
						</p>
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
<%end if%>


<!-- ************   SEPARADOR   ************ -->
<table width="698" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc" align="left">&nbsp;</td></tr>
</table>
<br>


<table class="notPrint" width="649" cellspacing="0">
<% if blnEdicaoBloqueada then %>
<tr>
	<td align="center"><a href="javascript:history.back();" title="volta para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
</tr>
<% else %>
<tr>
	<td align="left"><a href="javascript:history.back();" title="volta para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
	<td align="right"><div name="dATUALIZA" id="dATUALIZA">
		<a name="bATUALIZA" id="bATUALIZA" href="javascript:AtualizaCliente(fCAD)" title="atualiza o cadastro deste cliente">
		<img src="../botao/confirmar.gif" width="176" height="55" border="0"></a></div>
	</td>
</tr>
<% end if %>
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