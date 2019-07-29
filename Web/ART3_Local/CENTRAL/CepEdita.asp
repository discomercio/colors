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
'	  C E P E D I T A . A S P
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
	
'	OBTEM O ID
	dim s, usuario, cep_selecionado, tipo_operacao
	usuario = trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))
	
	if Not operacao_permitida(OP_CEN_CAD_CEP, s_lista_operacoes_permitidas) then 
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if
	
'	CEP A EDITAR
	cep_selecionado = retorna_so_digitos(trim(Request.Form("c_cep")))
	
	if (cep_selecionado="") then Response.Redirect("aviso.asp?id=" & ERR_CEP_NAO_ESPECIFICADO) 
	if (len(cep_selecionado) <> 8) then Response.Redirect("aviso.asp?id=" & ERR_CEP_INVALIDO) 
	
'	CONECTA COM O BANCO DE DADOS
	dim cn, rs
	If Not bdd_cep_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	set rs = cn.Execute("SELECT * FROM t_CEP_LOGRADOURO WHERE (cep8_log='" & cep_selecionado & "')")
	if Err <> 0 then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
	
	if rs.Eof then
		tipo_operacao = OP_INCLUI
	else
		tipo_operacao = OP_CONSULTA
		end if
	





' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________


function UF_monta_itens_select(byval id_default)
dim strResp, ha_default, strListaUF, strUF, vUF, intContador
	id_default = UCase(Trim("" & id_default))
	ha_default=False
	strListaUF="AC|AL|AM|AP|BA|CE|DF|ES|GO|MA|MG|MS|MT|PA|PB|PE|PI|PR|RJ|RN|RO|RR|RS|SC|SE|SP|TO"
	vUF=Split(strListaUF,"|")
	for intContador=LBound(vUF) to UBound(vUF)
		strUF = vUF(intContador)
		if (id_default<>"") And (id_default=strUF) then
			strResp = strResp & "<OPTION SELECTED"
			ha_default=True
		else
			strResp = strResp & "<OPTION"
			end if
		strResp = strResp & " VALUE='" & strUF & "'>"
		strResp = strResp & strUF
		strResp = strResp & "</OPTION>" & chr(13)
		next

	if Not ha_default then
		strResp = "<OPTION SELECTED VALUE=''>&nbsp;</OPTION>" & chr(13) & strResp
		end if
	
	UF_monta_itens_select = strResp
end function


function TipoLogradouro_monta_itens_select(byval id_default)
dim strResp, ha_default, strListaTipo, strTipo, vTipo, intContador
	id_default = UCase(Trim("" & id_default))
	ha_default=False
	strListaTipo="ACESSO|ALAMEDA|AREA|AV|BECO|BL|CAM|CJ|COND|ESCADARIA|EST|JD|LD|LOT|LRG|PASSAGEM|PC|PQ|QD|R|ROD|SERVIDAO|SETOR|TV|VIA|VIADUTO|VIELA|VILA"
	vTipo=Split(strListaTipo,"|")
	for intContador=LBound(vTipo) to UBound(vTipo)
		strTipo = vTipo(intContador)
		if (id_default<>"") And (id_default=strTipo) then
			strResp = strResp & "<OPTION SELECTED"
			ha_default=True
		else
			strResp = strResp & "<OPTION"
			end if
		strResp = strResp & " VALUE='" & strTipo & "'>"
		strResp = strResp & strTipo
		strResp = strResp & "</OPTION>" & chr(13)
		next

	if Not ha_default then
		strResp = "<OPTION SELECTED VALUE=''>&nbsp;</OPTION>" & chr(13) & strResp
	else
		strResp = "<OPTION VALUE=''>&nbsp;</OPTION>" & chr(13) & strResp
		end if
		
	TipoLogradouro_monta_itens_select = strResp
end function

%>

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

<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
function RemoveItem( f ) {
var b;
	b=window.confirm('Confirma a exclusão deste CEP?');
	if (b){
		f.tipo_operacao.value=OP_EXCLUI;
		dREMOVE.style.visibility="hidden";
		window.status = "Aguarde ...";
		f.submit();
		}
}

function AtualizaItem( f ) {

//  CEP
	if (!cep_ok(f.c_cep.value)) {
		alert('CEP inválido!!');
		f.c_cep.focus();
		return;
		}
//  UF
	if (trim(f.c_uf.value)=="") {
		alert('UF não foi informada!!');
		f.c_uf.focus();
		return;
		}

	if (!uf_ok(f.c_uf.value)) {
		alert('UF inválida!!');
		f.c_uf.focus();
		return;
		}

//  CIDADE
	if (trim(f.c_cidade.value)=="") {
		alert('Informe a cidade!!');
		f.c_cidade.focus();
		return;
		}
	
//  BAIRRO
	if (trim(f.c_bairro.value)=="") {
		alert('Preencha o bairro!!');
		f.c_bairro.focus();
		return;
		}

//  LOGRADOURO
	if (trim(f.c_logradouro.value)=="") {
		alert('Preencha o logradouro!!');
		f.c_logradouro.focus();
		return;
		}

//  TIPO DO LOGRADOURO
	if (trim(f.c_tipo_logradouro.value)=="") {
		if (!confirm('O tipo do logradouro não foi informado!!\nContinua assim mesmo?')) {
			f.c_tipo_logradouro.focus();
			return;
			}
		}

	dATUALIZA.style.visibility="hidden";
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

<link href="<%=URL_FILE__E_CSS%>" Rel="stylesheet" Type="text/css">
<link href="<%=URL_FILE__EPRINTER_CSS%>" Rel="stylesheet" Type="text/css" media="print">

<style TYPE="text/css">
#c_uf,#c_tipo_logradouro {
	margin: 4pt 0pt 4pt 10pt;
	vertical-align: top;
	}
</style>

<%	if tipo_operacao=OP_INCLUI then
		s = "fCAD.c_uf.focus()"
	else
		s = "focus()"
		end if
%>
<body onLoad="<%=s%>">
<center>



<!--  CADASTRO DO CEP -->

<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
<%	if tipo_operacao=OP_INCLUI then
		s = "Cadastro de Novo CEP"
	else
		s = "Consulta/Edição de CEP Cadastrado"
		end if
%>
	<td align="CENTER" vAlign="BOTTOM"><p class="PEDIDO"><%=s%><br><span class="C">&nbsp;</span></p></td>
</tr>
</table>
<br>


<!--  CAMPOS DO CADASTRO  -->
<form id="fCAD" name="fCAD" METHOD="POST" ACTION="CepAtualiza.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<INPUT type=HIDDEN name='tipo_operacao' id="tipo_operacao" value='<%=tipo_operacao%>'>

<!-- ************   CEP   ************ -->
<table width="649" class="Q" cellSpacing="0">
	<tr>
		<td class="MD" valign="TOP" width="15%">
			<p class="R">CEP</p>
			<p class="C">
			<input id="c_cep" name="c_cep" class="TA" value="<%=cep_formata(cep_selecionado)%>" readonly size="11" style="margin-top:8px; text-align:center; color:#0000ff">
			</p>
		</td>
		<td class="MD" width="15%">
			<p class="R">UF</p>
			<p class="C">
			<%if tipo_operacao=OP_CONSULTA then s=Trim("" & rs("uf_log")) else s=""%>
			<select id="c_uf" name="c_uf" style="margin-right:10px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) {this.options[0].selected=true;}" onkeypress="if (digitou_enter(true)) fCAD.c_cidade.focus();">
			<% =UF_monta_itens_select(s) %>
			</select>
			</p>
		</td>
		<td valign="TOP">
			<p class="R">CIDADE</p>
			<p class="C">
			<%if tipo_operacao=OP_CONSULTA then s=Trim("" & rs("nome_local")) else s=""%>
			<input id="c_cidade" name="c_cidade" class="TA" type="TEXT" maxlength="60" size="60" value="<%=s%>" style="margin-top:8px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCAD.c_bairro.focus(); filtra_nome_identificador();">
			</p>
		</td>
	</tr>
</table>

<!-- ************   BAIRRO   ************ -->
<table width="649" class="QS" cellSpacing="0">
	<tr>
		<td>
			<p class="R">BAIRRO</p>
			<p class="C">
			<%if tipo_operacao=OP_CONSULTA then s=Trim("" & rs("extenso_bai")) else s=""%>
			<input id="c_bairro" name="c_bairro" class="TA" type="TEXT" maxlength="72" size="72" value="<%=s%>" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCAD.c_tipo_logradouro.focus(); filtra_nome_identificador();">
			</p>
		</td>
	</tr>
</table>

<!-- ************   TIPO LOGRADOURO / LOGRADOURO   ************ -->
<table width="649" class="QS" cellSpacing="0">
	<tr>
		<td class="MD">
			<p class="R">TIPO DE LOGRADOURO</p>
			<p class="C">
			<%if tipo_operacao=OP_CONSULTA then s=Trim("" & rs("abrev_tipo")) else s=""%>
			<select id="c_tipo_logradouro" name="c_tipo_logradouro" style="margin-right:10px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) {this.options[0].selected=true;}" onkeypress="if (digitou_enter(true)) fCAD.c_logradouro.focus();">
			<% =TipoLogradouro_monta_itens_select(s) %>
			</select>
			</p>
		</td>
		<td valign="TOP">
			<p class="R">LOGRADOURO</p>
			<p class="C">
			<%if tipo_operacao=OP_CONSULTA then s=Trim("" & rs("nome_log")) else s=""%>
			<input id="c_logradouro" name="c_logradouro" class="TA" type="TEXT" maxlength="70" size="70" value="<%=s%>" style="margin-top:8px;" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCAD.c_complemento_logradouro.focus(); filtra_nome_identificador();">
			</p>
		</td>
	</tr>
</table>

<!-- ************   COMPLEMENTO LOGRADOURO   ************ -->
<table width="649" class="QS" cellSpacing="0">
	<tr>
		<td width="100%">
			<p class="R">COMPLEMENTO DO LOGRADOURO</p>
			<p class="C">
			<%if tipo_operacao=OP_CONSULTA then s=Trim("" & rs("comple_log")) else s=""%>
			<input id="c_complemento_logradouro" name="c_complemento_logradouro" class="TA" value="<%=s%>" maxlength="100" style="width:635px;" onkeypress="if (digitou_enter(true)) bATUALIZA.focus(); filtra_nome_identificador();">
			</p>
		</td>
	</tr>
</table>


<!-- ************   SEPARADOR   ************ -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>
<br>


<table class="notPrint" width="649" cellSpacing="0">
<tr>
	<% if (tipo_operacao=OP_INCLUI) then %>
		<td align="LEFT">
			<a href="javascript:history.back()" title="cancela as alterações no cadastro">
			<img src="../botao/cancelar.gif" width="176" height="55" border="0"></a>
		</td>
		<td align="RIGHT"><div name="dATUALIZA" id="dATUALIZA">
			<a name="bATUALIZA" id="bATUALIZA" href="javascript:AtualizaItem(fCAD)" title="atualiza o cadastro do CEP">
			<img src="../botao/confirmar.gif" width="176" height="55" border="0"></a></div>
		</td>
	<% elseif (tipo_operacao=OP_CONSULTA) then %>
		<td align="LEFT">
			<a href="javascript:history.back()" title="cancela as alterações no cadastro">
			<img src="../botao/cancelar.gif" width="176" height="55" border="0"></a>
		</td>
		<% if Trim("" & rs("chvlocal_log")) = "" then %>
		<!-- PERMITE EXCLUSÃO APENAS DE CEP'S CADASTRADOS MANUALMENTE ATRAVÉS DESTA PÁGINA -->
		<td align="CENTER"><div name="dREMOVE" id="dREMOVE"><a href="javascript:RemoveItem(fCAD)" title="remove o CEP cadastrado">
			<img src="../botao/remover.gif" width=176 height=55 border=0></a></div>
		</td>
		<% end if %>
		<td align="RIGHT"><div name="dATUALIZA" id="dATUALIZA">
			<a name="bATUALIZA" id="bATUALIZA" href="javascript:AtualizaItem(fCAD)" title="atualiza o cadastro do CEP">
			<img src="../botao/confirmar.gif" width="176" height="55" border="0"></a></div>
		</td>
	<% else %>
		<td align="CENTER">
			<a href="javascript:history.back()" title="retorna para página anterior">
			<img src="../botao/voltar.gif" width="176" height="55" border="0"></a>
		</td>
	<% end if %>
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