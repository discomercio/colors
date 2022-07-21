<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ===============================
'	  L O J A E D I T A . A S P
'     ===============================
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
	dim s, usuario, loja_selecionada, operacao_selecionada
	usuario = trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 
	
'	LOJA A EDITAR
	loja_selecionada = trim(request("loja_selecionada"))
	operacao_selecionada = trim(request("operacao_selecionada"))
	
	if operacao_selecionada=OP_INCLUI then
		loja_selecionada=retorna_so_digitos(loja_selecionada)
		end if

	loja_selecionada=normaliza_codigo(loja_selecionada, TAM_MIN_LOJA)
	
	if (loja_selecionada="") Or (loja_selecionada="00") then Response.Redirect("aviso.asp?id=" & ERR_LOJA_NAO_ESPECIFICADA) 
	if (operacao_selecionada<>OP_INCLUI) And (operacao_selecionada<>OP_CONSULTA) then Response.Redirect("aviso.asp?id=" & ERR_OPERACAO_NAO_ESPECIFICADA)
	
'	CONECTA COM O BANCO DE DADOS
	dim cn,rs
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	set rs = cn.Execute("select * from t_LOJA where (loja='" & loja_selecionada & "')")
	if Err <> 0 then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
	
	if operacao_selecionada=OP_INCLUI then
		if Not rs.EOF then Response.Redirect("aviso.asp?id=" & ERR_LOJA_JA_CADASTRADA)
	'	GARANTE QUE O Nº DE LOJA NÃO ESTÁ EM USO
		rs.Close
		set rs = cn.Execute("select * from t_LOJA where (CONVERT(smallint, loja) = " & loja_selecionada & ")")
		if Err <> 0 then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
		if Not rs.EOF then Response.Redirect("aviso.asp?id=" & ERR_LOJA_JA_CADASTRADA)
	elseif operacao_selecionada=OP_CONSULTA then
		if rs.EOF then Response.Redirect("aviso.asp?id=" & ERR_LOJA_NAO_CADASTRADA)
		end if




	
' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________

' _____________________________________________
' finPlanoContasEmpresaMontaItensSelect
'
function finPlanoContasEmpresaMontaItensSelect(byval id_default)
dim x, r, strResp, ha_default
	id_default = Trim("" & id_default)
	ha_default=False
	set r = cn.Execute("SELECT * FROM t_FIN_PLANO_CONTAS_EMPRESA ORDER BY id")
	strResp = ""
	do while Not r.eof 
		x = Trim("" & r("id"))
		if (converte_numero(id_default)<>0) And (converte_numero(id_default)=converte_numero(x)) then
			strResp = strResp & "<OPTION SELECTED"
			ha_default=True
		else
			strResp = strResp & "<OPTION"
			end if
		strResp = strResp & " VALUE='" & x & "'>"
		strResp = strResp & normaliza_codigo(x,TAM_PLANO_CONTAS__EMPRESA) & " - " & Trim("" & r("descricao"))
		strResp = strResp & "</OPTION>" & chr(13)
		r.MoveNext
		loop

	if Not ha_default then
		strResp = "<OPTION SELECTED VALUE=''>&nbsp;</OPTION>" & chr(13) & strResp
	else
		strResp = "<OPTION VALUE=''>&nbsp;</OPTION>" & chr(13) & strResp
		end if
		
	finPlanoContasEmpresaMontaItensSelect = strResp
	r.close
	set r=nothing
end function


' _____________________________________________
' finPlanoContasContaMontaItensSelect
'
function finPlanoContasContaMontaItensSelect(byval id_default)
dim x, r, strSql, strResp, ha_default
	id_default = Trim("" & id_default)
	ha_default=False
	strSql = "SELECT " & _
				"*" & _
			" FROM t_FIN_PLANO_CONTAS_CONTA" & _
			" WHERE" & _
				" natureza = '" & COD_FIN_NATUREZA__CREDITO & "'" & _
			" ORDER BY" & _
				" id"
	set r = cn.Execute(strSql)
	strResp = ""
	do while Not r.eof 
		x = Trim("" & r("id"))
		if (converte_numero(id_default)<>0) And (converte_numero(id_default)=converte_numero(x)) then
			strResp = strResp & "<OPTION SELECTED"
			ha_default=True
		else
			strResp = strResp & "<OPTION"
			end if
		strResp = strResp & " VALUE='" & x & "|" & Trim("" & r("id_plano_contas_grupo")) & "'>"
		strResp = strResp & normaliza_codigo(x,TAM_PLANO_CONTAS__CONTA) & " - " & Trim("" & r("descricao"))
		strResp = strResp & "</OPTION>" & chr(13)
		r.MoveNext
		loop

	if Not ha_default then
		strResp = "<OPTION SELECTED VALUE=''>&nbsp;</OPTION>" & chr(13) & strResp
	else
		strResp = "<OPTION VALUE=''>&nbsp;</OPTION>" & chr(13) & strResp
		end if
		
	finPlanoContasContaMontaItensSelect = strResp
	r.close
	set r=nothing
end function


' _____________________________________________
' unidadeNegocioMontaItensSelect
'
function unidadeNegocioMontaItensSelect(byval id_default)
dim x, r, strResp, ha_default
	id_default = Trim("" & id_default)
	ha_default=False
	set r = cn.Execute("SELECT * FROM t_CFG_UNIDADE_NEGOCIO ORDER BY ordenacao")
	strResp = ""
	do while Not r.eof 
		x = Trim("" & r("Sigla"))
		if (id_default <> "") And (UCase(id_default) = UCase(x)) then
			strResp = strResp & "<OPTION SELECTED"
			ha_default=True
		else
			strResp = strResp & "<OPTION"
			end if
		strResp = strResp & " VALUE='" & x & "'>"
		strResp = strResp & Trim("" & r("NomeCurto"))
		strResp = strResp & "</OPTION>" & chr(13)
		r.MoveNext
		loop

	if Not ha_default then
		strResp = "<OPTION SELECTED VALUE=''>&nbsp;</OPTION>" & chr(13) & strResp
	else
		strResp = "<OPTION VALUE=''>&nbsp;</OPTION>" & chr(13) & strResp
		end if
		
	unidadeNegocioMontaItensSelect = strResp
	r.close
	set r=nothing
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

<script src="<%=URL_FILE__JQUERY%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_I18N%>" Language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>
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

function RemoveLoja( f ) {
var b;
	b=window.confirm('Confirma a exclusão da loja?');
	if (b){
		f.operacao_selecionada.value=OP_EXCLUI;
		dREMOVE.style.visibility="hidden";
		window.status = "Aguarde ...";
		f.submit();
		}
}

function AtualizaLoja( f ) {
	if (trim(f.nome.value)=="") {
		alert('Preencha o nome!!');
		f.nome.focus();
		return;
		}
	if (!cnpj_ok(f.cnpj.value)) {
		alert('CNPJ inválido!!');
		f.cnpj.focus();
		return;
		}
	if (!cep_ok(f.cep.value)) {
		alert('CEP inválido!!');
		f.cep.focus();
		return;
		}
	if (!uf_ok(f.uf.value)) {
		alert('UF inválida!!');
		f.uf.focus();
		return;
		}
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
	if (trim(f.c_plano_contas_empresa.value)=="") {
		alert('Selecione a empresa para os lançamentos do fluxo de caixa!!');
		return;
	}
	if (trim(f.c_plano_contas_conta.value)=="") {
		alert('Selecione o plano de conta para os lançamentos do fluxo de caixa!!');
		return;
	}
	
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


<%	if operacao_selecionada=OP_INCLUI then
		s = "fCAD.nome.focus()"
	else
		s = "focus()"
		end if
%>
<body id="corpoPagina" onload="<%=s%>">

<center>

<!-- #include file = "../global/JanelaBuscaCEP.htm"    -->


<!--  CADASTRO DA LOJA -->
<br />
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
<%	if operacao_selecionada=OP_INCLUI then
		s = "Cadastro de Nova Loja"
	else
		s = "Consulta/Edição de Loja Cadastrada"
		end if
%>
	<td align="center" valign="bottom"><span class="PEDIDO"><%=s%></span></td>
</tr>
</table>
<br>


<!--  CAMPOS DO CADASTRO  -->
<form id="fCAD" name="fCAD" method="post" action="LojaAtualiza.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name='operacao_selecionada' id="operacao_selecionada" value='<%=operacao_selecionada%>'>

<!-- ************   NÚMERO/NOME   ************ -->
<table width="649" class="Q" cellspacing="0">
	<tr>
		<td class="MD" width="10%" align="left"><p class="R">LOJA</p><p class="C"><input id="loja_selecionada" name="loja_selecionada" class="TA" value="<%=loja_selecionada%>" readonly size="6" style="text-align:center; color:#0000ff"></p></td>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("nome")) else s=""%>
		<td width="90%" align="left"><p class="R">NOME (apelido)</p><p class="C"><input id="nome" name="nome" class="TA" type="text" maxlength="30" size="60" value="<%=s%>" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCAD.razao_social.focus(); filtra_nome_identificador();"></p></td>
	</tr>
</table>

<!-- ************   RAZÃO SOCIAL   ************ -->
<table width="649" class="QS" cellSpacing="0">
	<tr>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("razao_social")) else s=""%>
		<td width="100%" align="left"><p class="R">RAZÃO SOCIAL</p><p class="C"><input id="razao_social" name="razao_social" class="TA" value="<%=s%>" maxlength="60" size="85" onkeypress="if (digitou_enter(true)) fCAD.cnpj.focus(); filtra_nome_identificador();"></p></td>
	</tr>
</table>

<!-- ************   CNPJ/IE   ************ -->
<table width="649" class="QS" cellSpacing="0">
	<tr>
<%if operacao_selecionada=OP_CONSULTA then s=cnpj_cpf_formata(Trim("" & rs("cnpj"))) else s=""%>
	<td class="MD" width="50%" align="left"><p class="R">CNPJ</p><p class="C">
		<input id="cnpj" name="cnpj" class="TA" value="<%=s%>" maxlength="18" size="24" onblur="if (!cnpj_ok(this.value)) {alert('CNPJ inválido'); this.focus();} else this.value=cnpj_formata(this.value);" onkeypress="if (digitou_enter(true) && cnpj_ok(this.value)) ie.focus(); filtra_cnpj();"></p></td>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("ie")) else s=""%>
		<td width="50%" align="left"><p class="R">IE</p><p class="C"><input id="ie" name="ie" class="TA" type="TEXT" maxlength="20" size="25" value="<%=s%>" onkeypress="if (digitou_enter(true)) fCAD.endereco.focus(); filtra_nome_identificador();"></p></td>
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
		<input id="endereco_complemento" name="endereco_complemento" class="TA" value="<%=s%>" maxlength="60" style="width:310px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCAD.bairro.focus(); filtra_nome_identificador();"></p></td>
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
		<td width="50%" class="MD" align="left"><p class="R">UF</p><p class="C"><input id="uf" name="uf" class="TA" value="<%=s%>" maxlength="2" size="3" onkeypress="if (digitou_enter(true) && uf_ok(this.value)) fCAD.ddd.focus();" onblur="this.value=trim(this.value); if (!uf_ok(this.value)) {alert('UF inválida!!');this.focus();} else this.value=ucase(this.value);"></p></td>
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

<!-- ************   DDD/TELEFONE/FAX   ************ -->
<table width="649" class="QS" cellSpacing="0">
	<tr>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("ddd")) else s=""%>
		<td width="15%" class="MD" align="left"><p class="R">DDD</p><p class="C"><input id="ddd" name="ddd" class="TA" value="<%=s%>" maxlength="4" size="5" onkeypress="if (digitou_enter(true) && ddd_ok(this.value)) fCAD.telefone.focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inválido!!');this.focus();}"></p></td>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("telefone")) else s=""%>
		<td width="35%" class="MD" align="left"><p class="R">TELEFONE</p><p class="C"><input id="telefone" name="telefone" class="TA" value="<%=telefone_formata(s)%>" maxlength="11" size="12" onkeypress="if (digitou_enter(true) && telefone_ok(this.value)) fCAD.fax.focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('Telefone inválido!!');this.focus();} else this.value=telefone_formata(this.value);"></p></td>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("fax")) else s=""%>
		<td align="left"><p class="R">FAX</p><p class="C"><input id="fax" name="fax" class="TA" value="<%=telefone_formata(s)%>" maxlength="11" size="12" onkeypress="if (digitou_enter(true) && telefone_ok(this.value)) fCAD.comissao_indicacao.focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('Fax inválido!!');this.focus();} else this.value=telefone_formata(this.value);"></p></td>
	</tr>
</table>

<!-- ************   PERCENTUAL DA COMISSÃO POR INDICAÇÃO   ************ -->
<table width="649" class="QS" cellSpacing="0">
	<tr>
		<%if operacao_selecionada=OP_CONSULTA then s=formata_perc_comissao(rs("comissao_indicacao")) else s=""%>
		<td class="MD" width="50%" align="left"><p class="R">COMISSÃO DA LOJA POR INDICAÇÕES (%)</p><p class="C"><input id="comissao_indicacao" name="comissao_indicacao" class="TA" value="<%=s%>" maxlength="5" style="width:70px;" 
			onkeypress="if (digitou_enter(true)) fCAD.c_max_senha_desconto.focus(); filtra_percentual();"
			onblur="this.value=formata_perc_desc(this.value); if ((converte_numero(this.value)>100)||(converte_numero(this.value)<0)) {alert('Percentual inválido!!');this.focus();}"></p></td>
		<%if operacao_selecionada=OP_CONSULTA then s=formata_perc(rs("PercMaxSenhaDesconto")) else s=""%>
		<td align="left"><p class="R">MÁX SENHA DE DESCONTO (%)</p><p class="C"><input id="c_PercMaxSenhaDesconto" name="c_PercMaxSenhaDesconto" readonly tabindex=-1 class="TA" value="<%=s%>" maxlength="6" style="width:70px;color:gray;" 
			onkeypress="if (digitou_enter(true)) bATUALIZA.focus(); filtra_percentual();"
			onblur="this.value=formata_numero(this.value,2); if ((converte_numero(this.value)>100)||(converte_numero(this.value)<0)) {alert('Percentual inválido!!');this.focus();}"></p>
		</td>
	</tr>
</table>

<!-- ************   PERCENTUAL MÁXIMO DE DESCONTO SEM ZERAR A COMISSÃO (ANTERIORMENTE CHAMADO DE RT)   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
	<%if operacao_selecionada=OP_CONSULTA then s=formata_perc(rs("PercMaxDescSemZerarRT")) else s=""%>
	<td align="left"><p class="R">MÁX DESCONTO SEM ZERAR COMISSÃO (%)</p><p class="C"><input id="c_PercMaxDescSemZerarRT" name="c_PercMaxDescSemZerarRT" readonly tabindex=-1 class="TA" value="<%=s%>" maxlength="6" style="width:70px;color:gray;"
			onkeypress="if (digitou_enter(true)) bATUALIZA.focus(); filtra_percentual();"
			onblur="this.value=formata_numero(this.value,2); if ((converte_numero(this.value)>100)||(converte_numero(this.value)<0)) {alert('Percentual inválido!!');this.focus();}"></p>
		</td>
	</tr>
</table>

<!-- ************   PLANO CONTAS EMPRESA / PLANO CONTAS CONTA   ************ -->
<table width="649" class="QS" cellSpacing="0">
	<tr>
		<td align="left">
			<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("id_plano_contas_empresa")) else s=""%>
			<p class="R">EMPRESA (LANÇAMENTOS DO FLUXO DE CAIXA)</p>
			<select id="c_plano_contas_empresa" name="c_plano_contas_empresa" style="margin-left:4px;margin-top:4px;margin-bottom:8px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
			<%=finPlanoContasEmpresaMontaItensSelect(s)%>
			</select>
		</td>
	</tr>
	<tr>
		<td class="MC" align="left">
			<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("id_plano_contas_conta")) else s=""%>
			<p class="R">PLANO DE CONTA (LANÇAMENTOS DO FLUXO DE CAIXA)</p>
			<select id="c_plano_contas_conta" name="c_plano_contas_conta" style="margin-left:4px;margin-top:4px;margin-bottom:8px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
			<%=finPlanoContasContaMontaItensSelect(s)%>
			</select>
		</td>
	</tr>
</table>

<!-- ************   COMISSÃO INDICADORES: PLANO CONTAS EMPRESA   ************ -->
<table width="649" class="QS" cellSpacing="0">
	<tr>
		<td align="left">
			<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("id_plano_contas_empresa_comissao_indicador")) else s=""%>
			<p class="R">EMPRESA (LANÇAMENTOS DO FLUXO DE CAIXA REF COMISSÃO INDICADOR)</p>
			<select id="c_plano_contas_empresa_comissao_indicador" name="c_plano_contas_empresa_comissao_indicador" style="margin-left:4px;margin-top:4px;margin-bottom:8px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
			<%=finPlanoContasEmpresaMontaItensSelect(s)%>
			</select>
		</td>
	</tr>
</table>

<!-- ************   UNIDADE DE NEGÓCIO   ************ -->
<table width="649" class="QS" cellSpacing="0">
	<tr>
		<td align="left">
			<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("unidade_negocio")) else s=""%>
			<p class="R">UNIDADE DE NEGÓCIO</p>
			<select id="c_unidade_negocio" name="c_unidade_negocio" style="margin-left:4px;margin-top:4px;margin-bottom:8px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
			<%=unidadeNegocioMontaItensSelect(s)%>
			</select>
		</td>
	</tr>
</table>


<!-- ************   SEPARADOR   ************ -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc" align="left">&nbsp;</td></tr>
</table>
<br>


<table class="notPrint" width="649" cellSpacing="0">
<tr>
	<td align="left"><a href="javascript:history.back()" title="cancela as alterações no cadastro">
		<img src="../botao/cancelar.gif" width="176" height="55" border="0"></a></td>
	<%
	s = ""
	if operacao_selecionada=OP_CONSULTA then
		s = "<td align='center'><div name='dREMOVE' id='dREMOVE'><a href='javascript:RemoveLoja(fCAD)' "
		s =s + "title='remove a loja cadastrada'><img src='../botao/remover.gif' width=176 height=55 border=0></a></div>"
		end if
	%><%=s%>
	<td align="right"><div name="dATUALIZA" id="dATUALIZA">
		<a name="bATUALIZA" id="bATUALIZA" href="javascript:AtualizaLoja(fCAD)" title="atualiza o cadastro da loja">
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