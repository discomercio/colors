<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =====================================
'	  F A B R I C A N T E E D I T A . A S P
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
	dim s, usuario, fabricante_selecionado, operacao_selecionada
	usuario = trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO)
	
'	FABRICANTE A EDITAR
	fabricante_selecionado = trim(request("fabricante_selecionado"))
	operacao_selecionada = trim(request("operacao_selecionada"))
	
	if operacao_selecionada=OP_INCLUI then
		fabricante_selecionado=retorna_so_digitos(fabricante_selecionado)
		end if

	fabricante_selecionado=normaliza_codigo(fabricante_selecionado, TAM_MIN_FABRICANTE)
	
	if (fabricante_selecionado="") Or (fabricante_selecionado="000") then Response.Redirect("aviso.asp?id=" & ERR_FABRICANTE_NAO_ESPECIFICADO) 
	if (operacao_selecionada<>OP_INCLUI) And (operacao_selecionada<>OP_CONSULTA) then Response.Redirect("aviso.asp?id=" & ERR_OPERACAO_NAO_ESPECIFICADA)
	
'	CONECTA COM O BANCO DE DADOS
	dim cn,rs
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	set rs = cn.Execute("select * from t_FABRICANTE where (fabricante='" & fabricante_selecionado & "')")
	if Err <> 0 then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
		
	if operacao_selecionada=OP_INCLUI then
		if Not rs.EOF then Response.Redirect("aviso.asp?id=" & ERR_FABRICANTE_JA_CADASTRADO)
	'	GARANTE QUE O Nº DO FABRICANTE NÃO ESTÁ EM USO
		rs.Close
		set rs = cn.Execute("select * from t_FABRICANTE where (CONVERT(smallint, fabricante) = " & fabricante_selecionado & ")")
		if Err <> 0 then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
		if Not rs.EOF then Response.Redirect("aviso.asp?id=" & ERR_FABRICANTE_JA_CADASTRADO)
	elseif operacao_selecionada=OP_CONSULTA then
		if rs.EOF then Response.Redirect("aviso.asp?id=" & ERR_FABRICANTE_NAO_CADASTRADO)
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

function RemoveFabricante( f ) {
var b;
	b=window.confirm('Confirma a exclusão do fabricante?');
	if (b){
		f.operacao_selecionada.value=OP_EXCLUI;
		dREMOVE.style.visibility="hidden";
		window.status = "Aguarde ...";
		f.submit();
		}
}

function AtualizaFabricante( f ) {
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


<!--  CADASTRO DO FABRICANTE -->
<br />
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
<%	if operacao_selecionada=OP_INCLUI then
		s = "Cadastro de Novo Fabricante"
	else
		s = "Consulta/Edição de Fabricante Cadastrado"
		end if
%>
	<td align="center" valign="bottom"><span class="PEDIDO"><%=s%></span></td>
</tr>
</table>
<br>


<!--  CAMPOS DO CADASTRO  -->
<form id="fCAD" name="fCAD" method="post" action="FabricanteAtualiza.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name='operacao_selecionada' id="operacao_selecionada" value='<%=operacao_selecionada%>'>

<!-- ************   NÚMERO/NOME   ************ -->
<table width="649" class="Q" cellspacing="0">
	<tr>
		<td class="MD" width="15%" align="left"><p class="R">FABRICANTE</p><p class="C"><input id="fabricante_selecionado" name="fabricante_selecionado" class="TA" value="<%=fabricante_selecionado%>" readonly size="6" style="text-align:center; color:#0000ff"></p></td>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("nome")) else s=""%>
		<td width="85%" align="left">
			<p class="R">NOME (apelido)</p>
			<p class="C">
			<input id="nome" name="nome" class="TA" type="text" maxlength="30" size="60" value="<%=s%>"
				onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCAD.razao_social.focus(); filtra_nome_identificador();">
			</p>
		</td>
	</tr>
</table>

<!-- ************   RAZÃO SOCIAL   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("razao_social")) else s=""%>
		<td width="100%" align="left">
			<p class="R">RAZÃO SOCIAL</p>
			<p class="C">
			<input id="razao_social" name="razao_social" class="TA" value="<%=s%>" maxlength="60" size="85"
				onkeypress="if (digitou_enter(true)) fCAD.cnpj.focus(); filtra_nome_identificador();">
			</p>
		</td>
	</tr>
</table>

<!-- ************   CNPJ/IE   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
<%if operacao_selecionada=OP_CONSULTA then s=cnpj_cpf_formata(Trim("" & rs("cnpj"))) else s=""%>
		<td class="MD" width="50%" align="left">
			<p class="R">CNPJ</p>
			<p class="C">
			<input id="cnpj" name="cnpj" class="TA" value="<%=s%>" maxlength="18" size="24"
				onblur="if (!cnpj_ok(this.value)) {alert('CNPJ inválido'); this.focus();} else this.value=cnpj_formata(this.value);"
				onkeypress="if (digitou_enter(true) && cnpj_ok(this.value)) fCAD.ie.focus(); filtra_cnpj();">
			</p>
		</td>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("ie")) else s=""%>
		<td width="50%" align="left">
			<p class="R">IE</p>
			<p class="C">
			<input id="ie" name="ie" class="TA" type="text" maxlength="20" size="25" value="<%=s%>"
				onkeypress="if (digitou_enter(true)) fCAD.endereco.focus(); filtra_nome_identificador();">
			</p>
		</td>
	</tr>
</table>

<!-- ************   ENDEREÇO   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("endereco")) else s=""%>
		<td width="100%" align="left">
			<p class="R">ENDEREÇO</p>
			<p class="C">
			<input id="endereco" name="endereco" class="TA" value="<%=s%>" maxlength="60" style="width:635px;"
				onkeypress="if (digitou_enter(true)) fCAD.endereco_numero.focus(); filtra_nome_identificador();">
			</p>
		</td>
	</tr>
</table>

<!-- ************   Nº/COMPLEMENTO   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
		<td class="MD" width="50%" align="left">
			<p class="R">Nº</p>
			<p class="C">
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("endereco_numero")) else s=""%>
			<input id="endereco_numero" name="endereco_numero" class="TA" value="<%=s%>" maxlength="20" style="width:310px;"
				onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCAD.endereco_complemento.focus(); filtra_nome_identificador();">
			</p>
		</td>
		<td align="left">
			<p class="R">COMPLEMENTO</p>
			<p class="C">
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("endereco_complemento")) else s=""%>
			<input id="endereco_complemento" name="endereco_complemento" class="TA" value="<%=s%>" maxlength="60" style="width:310px;"
				onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCAD.bairro.focus(); filtra_nome_identificador();">
			</p>
		</td>
	</tr>
</table>

<!-- ************   BAIRRO/CIDADE   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("bairro")) else s=""%>
		<td width="50%" class="MD" align="left">
			<p class="R">BAIRRO</p>
			<p class="C">
			<input id="bairro" name="bairro" class="TA" value="<%=s%>" maxlength="72" style="width:310px;"
				onkeypress="if (digitou_enter(true)) fCAD.cidade.focus(); filtra_nome_identificador();">
			</p>
		</td>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("cidade")) else s=""%>
		<td width="50%" align="left">
			<p class="R">CIDADE</p>
			<p class="C">
			<input id="cidade" name="cidade" class="TA" value="<%=s%>" maxlength="60" style="width:310px;"
				onkeypress="if (digitou_enter(true)) fCAD.uf.focus(); filtra_nome_identificador();">
			</p>
		</td>
	</tr>
</table>

<!-- ************   UF/CEP   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("uf")) else s=""%>
		<td width="50%" class="MD" align="left">
			<p class="R">UF</p>
			<p class="C">
			<input id="uf" name="uf" class="TA" value="<%=s%>" maxlength="2" size="3"
				onkeypress="if (digitou_enter(true) && uf_ok(this.value)) fCAD.ddd.focus();"
				onblur="this.value=trim(this.value); if (!uf_ok(this.value)) {alert('UF inválida!!');this.focus();} else this.value=ucase(this.value);">
			</p>
		</td>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("cep")) else s=""%>
		<td width="25%" align="left">
			<p class="R">CEP</p>
			<p class="C">
			<input id="cep" name="cep" readonly tabindex=-1 class="TA" value="<%=cep_formata(s)%>" maxlength="9" size="11"
				onkeypress="if (digitou_enter(true) && cep_ok(this.value)) fCAD.ddd.focus(); filtra_cep();"
				onblur="if (!cep_ok(this.value)) {alert('CEP inválido!!');this.focus();} else this.value=cep_formata(this.value);">
			</p>
		</td>
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
<table width="649" class="QS" cellspacing="0">
	<tr>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("ddd")) else s=""%>
		<td width="15%" class="MD" align="left">
			<p class="R">DDD</p>
			<p class="C">
			<input id="ddd" name="ddd" class="TA" value="<%=s%>" maxlength="4" size="5"
				onkeypress="if (digitou_enter(true) && ddd_ok(this.value)) fCAD.telefone.focus(); filtra_numerico();"
				onblur="if (!ddd_ok(this.value)) {alert('DDD inválido!!');this.focus();}">
			</p>
		</td>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("telefone")) else s=""%>
		<td width="35%" class="MD" align="left">
			<p class="R">TELEFONE</p>
			<p class="C">
			<input id="telefone" name="telefone" class="TA" value="<%=telefone_formata(s)%>" maxlength="11" size="12"
				onkeypress="if (digitou_enter(true) && telefone_ok(this.value)) fCAD.fax.focus(); filtra_numerico();"
				onblur="if (!telefone_ok(this.value)) {alert('Telefone inválido!!');this.focus();} else this.value=telefone_formata(this.value);">
			</p>
		</td>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("fax")) else s=""%>
		<td align="left">
			<p class="R">FAX</p>
			<p class="C">
			<input id="fax" name="fax" class="TA" value="<%=telefone_formata(s)%>" maxlength="11" size="12"
				onkeypress="if (digitou_enter(true) && telefone_ok(this.value)) fCAD.contato.focus(); filtra_numerico();"
				onblur="if (!telefone_ok(this.value)) {alert('Fax inválido!!');this.focus();} else this.value=telefone_formata(this.value);">
			</p>
		</td>
	</tr>
</table>

<!-- ************   CONTATO   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("contato")) else s=""%>
		<td width="100%" align="left">
			<p class="R">CONTATO</p>
			<p class="C">
			<input id="contato" name="contato" class="TA" value="<%=s%>" maxlength="40" size="85"
				onkeypress="if (digitou_enter(true)) fCAD.markup.focus(); filtra_nome_identificador();">
			</p>
		</td>
	</tr>
</table>

<!-- ************   MARKUP   ************ -->
<table width="649" class="QS" cellSpacing="0">
	<tr>
<%if operacao_selecionada=OP_CONSULTA then s=formata_perc_markup("" & rs("markup")) else s=""%>
		<td width="100%" align="left">
			<p class="R">MARK UP</p>
			<p class="C">
			<input id="markup" name="markup" class="TA" value="<%=s%>" maxlength="5" size="85" 
				onkeypress="if (digitou_enter(true)) bATUALIZA.focus(); filtra_percentual();"
				onblur="this.value=formata_perc_desc(this.value); if ((converte_numero(this.value)>100)||(converte_numero(this.value)<0)) {alert('Percentual inválido!!');this.focus();}">
			</p>
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
		s = "<td align='center'><div name='dREMOVE' id='dREMOVE'><a href='javascript:RemoveFabricante(fCAD)' "
		s =s + "title='remove o fabricante cadastrado'><img src='../botao/remover.gif' width=176 height=55 border=0></a></div>"
		end if
	%><%=s%>
	<td align="right"><div name="dATUALIZA" id="dATUALIZA">
		<a name="bATUALIZA" id="bATUALIZA" href="javascript:AtualizaFabricante(fCAD)" title="atualiza o cadastro do fabricante">
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