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
'	  ClienteConsultaView.asp
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
	
'	OBTEM O ID
	dim intCounter
	dim s, s_aux, usuario
	usuario = trim(Session("usuario_atual"))
	If (usuario = "") then
		usuario = Trim(Request("usuario"))
		Session("usuario_atual") = usuario
		end if
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 
	
'	CLIENTE A CONSULTAR
	dim id_cliente
	id_cliente = trim(request("cliente_selecionado"))
	if id_cliente = "" then Response.Redirect("aviso.asp?id=" & ERR_ID_INVALIDO)
	
	dim ocultar_botoes
	ocultar_botoes = UCase(Trim(Request("ocultar_botoes")))
	
'	HÁ PEDIDO P/ ONDE RETORNAR?
	dim pedido_selecionado, pedido_selecionado_inicial
	pedido_selecionado = ucase(Trim(request("pedido_selecionado")))
	pedido_selecionado_inicial = Trim(Request("pedido_selecionado_inicial"))
	
'	HÁ ORÇAMENTO P/ ONDE RETORNAR?
	dim  orcamento_selecionado, orcamento_selecionado_inicial
	orcamento_selecionado = ucase(Trim(request("orcamento_selecionado")))
	orcamento_selecionado_inicial = Trim(Request("orcamento_selecionado_inicial"))
	
	dim pagina_retorno, s_url
	pagina_retorno = Trim(Request("pagina_retorno"))
	
'	CONECTA COM O BANCO DE DADOS
	dim cn,rs,tRefBancaria,tRefComercial,tRefProfissional
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))
	if s_lista_operacoes_permitidas = "" then
		s_lista_operacoes_permitidas = obtem_operacoes_permitidas_usuario(cn, usuario)
		Session("lista_operacoes_permitidas") = s_lista_operacoes_permitidas
		end if

	s = "SELECT * FROM t_CLIENTE WHERE (id='" & id_cliente & "')"
	set rs = cn.Execute(s)
	if Err <> 0 then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
	if rs.EOF then Response.Redirect("aviso.asp?id=" & ERR_CLIENTE_NAO_CADASTRADO)

	dim eh_cpf
	s=Trim("" & rs("cnpj_cpf"))
	if len(s)=11 then eh_cpf=True else eh_cpf=False
	
	dim s_codigo, s_descricao, s_codigo_e_descricao
	
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


<%=DOCTYPE_LEGADO%>

<html>


<head>
	<title>LOJA</title>
	</head>


<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
	window.status = "";
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

<body>

<center>


<!--  CADASTRO DO CLIENTE -->

<table width="698" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="center" valign="bottom"><span class="PEDIDO">Cadastro do Cliente</span></td>
</tr>
</table>
<br>

<!-- ************   EXIBE OBSERVAÇÕES CREDITÍCIAS?  ************ -->
<%	s = Trim("" & rs("obs_crediticias"))
	if s <> "" then %>
		<span class="Lbl" style="display:none">OBSERVAÇÕES CREDITÍCIAS</span>
		<div class='MtAviso' style="width:649px;FONT-WEIGHT:bold;border:1pt solid black;display:none;" align="CENTER"><p style='margin:5px 2px 5px 2px;'><%=s%></p></div>
		<br>
<% end if %>


<!-- ************  CAMPOS DO CADASTRO  ************ -->
<form id="fCAD" name="fCAD" method="post">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name='cliente_selecionado' id="cliente_selecionado" value='<%=id_cliente%>'>
<input type="hidden" name="pedido_selecionado" value="<%=pedido_selecionado%>">
<input type="hidden" name="pedido_selecionado_inicial" value="<%=pedido_selecionado_inicial%>">
<input type="hidden" name="usuario" value="<%=usuario%>">
<input type="hidden" name="ocultar_botoes" value="<%=ocultar_botoes%>">



<!-- ************   CNPJ/IE OU CPF/RG/NASCIMENTO/SEXO  ************ -->
<table width="649" class="Q" cellspacing="0">
	<tr>
	<td class="MD" align="left">
		<%if eh_cpf then s="CPF" else s="CNPJ"%>
		<p class="R"><%=s%></p>
		<p class="C">
		<input id="cnpj_cpf_selecionado" name="cnpj_cpf_selecionado" class="TA" value="<%=cnpj_cpf_formata(Trim("" & rs("cnpj_cpf")))%>" readonly size="22" style="text-align:center; color:#0000ff">
		</p>
	</td>

<%if eh_cpf then%>
	<td class="MD" align="left">
		<p class="R">RG</p>
		<p class="C">
		<input id="rg" name="rg" class="TA" type="text" maxlength="20" size="15" value="<%=Trim("" & rs("rg"))%>" readonly>
		</p>
	</td>
    <td width="50%" class="MD" align="left">
		<p class="R">PRODUTOR RURAL</p>
		<p class="C">
        <% if Trim("" & rs("produtor_rural_status")) = COD_ST_CLIENTE_PRODUTOR_RURAL_SIM then
                s = Trim("" & rs("contribuinte_icms_status"))
                if s = COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO then
                    s_aux = "Sim (Não contribuinte)"
                elseif s = COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM then
                    s_aux = "Sim (IE: " & Trim("" & rs("ie")) & ")"
                elseif s = COD_ST_CLIENTE_CONTRIBUINTE_ICMS_ISENTO then
                    s_aux = "Sim (Isento)"
                end if
            elseif Trim("" & rs("produtor_rural_status")) = COD_ST_CLIENTE_PRODUTOR_RURAL_NAO then
                s_aux = "Não"
            end if
            %>
		<input id="c_contribuinte_icms" name="c_contribuinte_icms" class="TA" type="TEXT" value="<%=s_aux%>" readonly>
		</p>
	</td>
	<td class="MD" align="left">
		<p class="R">NASCIMENTO</p>
		<p class="C">
		<input id="dt_nasc" name="dt_nasc" class="TA" type="text" maxlength="10" size="10" value="<%=formata_data(rs("dt_nasc"))%>" readonly>
		</p>
	</td>
	<td align="left">
		<p class="R">SEXO</p>
		<p class="C">
		<input id="sexo" name="sexo" class="TA" type="text" maxlength="1" size="2" value="<%=Trim("" & rs("sexo"))%>" readonly>
		</p>
	</td>

<%else%>
	<td width="50%" class="MD" align="left">
		<p class="R">IE</p>
		<p class="C">
		<input id="ie" name="ie" class="TA" type="TEXT" maxlength="20" size="25" value="<%=Trim("" & rs("ie"))%>" readonly>
		</p>
	</td>
    <td width="50%" align="left">        
		<p class="R">CONTRIBUINTE ICMS</p>
		<p class="C">
        <% s = Trim("" & rs("contribuinte_icms_status"))
            if s = COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO then
                s_aux = "Não"
            elseif s = COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM then
                s_aux = "Sim"
            elseif s = COD_ST_CLIENTE_CONTRIBUINTE_ICMS_ISENTO then
                s_aux = "Isento"
            end if            
            %>
		<input id="c_contribuinte_icms" name="c_contribuinte_icms" class="TA" type="TEXT" size="25" value="<%=s_aux%>" readonly>
		</p>
	</td>
<%end if%>
	</tr>
</table>

<!-- ************   NOME  ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
	<%if eh_cpf then s="NOME" else s="RAZÃO SOCIAL"%>
	<td width="100%" align="left">
		<p class="R"><%=s%></p>
		<p class="C">
		<input id="nome" name="nome" class="TA" value="<%=Trim("" & rs("nome"))%>" maxlength="60" size="85" readonly>
		</p>
	</td>
	</tr>
</table>

<!-- ************   ENDEREÇO   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
	<td width="100%" align="left">
		<p class="R">ENDEREÇO</p>
		<p class="C">
		<input id="endereco" name="endereco" class="TA" value="<%=Trim("" & rs("endereco"))%>" maxlength="60" style="width:635px;" readonly>
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
		<input id="endereco_numero" name="endereco_numero" class="TA" value="<%=Trim("" & rs("endereco_numero"))%>" maxlength="20" style="width:310px;" readonly>
		</p>
	</td>
	<td align="left">
		<p class="R">COMPLEMENTO</p>
		<p class="C">
		<input id="endereco_complemento" name="endereco_complemento" class="TA" value="<%=Trim("" & rs("endereco_complemento"))%>" maxlength="60" style="width:310px;" readonly>
		</p>
	</td>
	</tr>
</table>

<!-- ************   BAIRRO/CIDADE   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
	<td class="MD" width="50%" align="left">
		<p class="R">BAIRRO</p>
		<p class="C">
		<input id="bairro" name="bairro" class="TA" value="<%=Trim("" & rs("bairro"))%>" maxlength="72" style="width:310px;" readonly>
		</p>
	</td>
	<td align="left">
		<p class="R">CIDADE</p>
		<p class="C">
		<input id="cidade" name="cidade" class="TA" value="<%=Trim("" & rs("cidade"))%>" maxlength="60" style="width:310px;" readonly>
		</p>
	</td>
	</tr>
</table>

<!-- ************   UF/CEP   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
	<td class="MD" width="50%" align="left">
		<p class="R">UF</p>
		<p class="C">
		<input id="uf" name="uf" class="TA" value="<%=Trim("" & rs("uf"))%>" maxlength="2" size="3" readonly>
		</p>
	</td>
	<td width="50%" align="left">
		<table width="100%" cellpadding="0" cellspacing="0">
		<tr>
			<td width="50%" align="left">
				<p class="R">CEP</p>
				<p class="C">
				<input id="cep" name="cep" class="TA" readonly tabindex=-1 value="<%=cep_formata(Trim("" & rs("cep")))%>" maxlength="9" size="11">
				</p>
			</td>
			<td align="center" width="50%">&nbsp;</td>
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
		<input id="tel_cel" name="tel_cel" class="TA" value="<%=telefone_formata(Trim("" & rs("tel_cel")))%>" maxlength="9" size="12" onkeypress="if (digitou_enter(true) && telefone_ok(this.value)) fCAD.ddd_com.focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('Número de celular inválido!!');this.focus();} else this.value=telefone_formata(this.value);"></p></td>
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
	<td width="100%" align="left">
		<p class="R">OBSERVAÇÃO</p>
		<p class="C">
		<input id="filiacao" name="filiacao" class="TA" value="<%=Trim("" & rs("filiacao"))%>" maxlength="60" size="74" readonly>
		</p>
	</td>
	</tr>
</table>
<% else %>
<!-- ************   CONTATO   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
	<td width="100%" align="left">
		<p class="R">NOME DA PESSOA PARA CONTATO NA EMPRESA</p>
		<p class="C">
		<input id="contato" name="contato" class="TA" value="<%=Trim("" & rs("contato"))%>" maxlength="30" size="45" readonly>
		</p>
	</td>
	</tr>
</table>
<% end if %>

<!-- ************   E-MAIL   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
	<td width="100%" align="left">
		<p class="R">E-MAIL</p>
		<p class="C">
		<input id="email" name="email" class="TA" value="<%=Trim("" & rs("email"))%>" maxlength="60" size="74" readonly>
		</p>
	</td>
	</tr>
</table>

<!-- ************   OBS CREDITÍCIAS   ************ -->
<table width="649" class="QS" cellspacing="0" style="display:none">
	<tr>
	<td width="100%" align="left">
		<p class="R">OBSERVAÇÕES CREDITÍCIAS</p>
		<p class="C">
		<input id="obs_crediticias" name="obs_crediticias" class="TA" value="<%=Trim("" & rs("obs_crediticias"))%>" maxlength="50" size="65" readonly>
		</p>
	</td>
	</tr>
</table>

<!-- ************   MÍDIA   ************ -->
<%	s_codigo = Trim("" & rs("midia"))
	s_descricao = ""
	s_codigo_e_descricao = ""
	if s_codigo <> "" then s_descricao = x_midia(s_codigo)
	if (s_codigo <> "") And (s_descricao <> "") then s_codigo_e_descricao = s_codigo & " - " & s_descricao
%>
<table width="649" class="QS" cellspacing="0" style="display:none">
	<tr>
	<% if rs("spc_negativado_status") = 1 then %>
	<td class="MD" width="50%" align="left">
		<p class="R">FORMA PELA QUAL CONHECEU A DIS</p>
		<p class="C">
		<input id="midia" name="midia" class="TA" value="<%=s_codigo_e_descricao%>" style="width:310px;" readonly />
		</p>
	</td>
	<td width="50%" align="left" valign="top">
		<p class="R">SPC</p>
		<p class="C">
		<input id="infoSPC" name="infoSPC" class="TA" readonly style="color: #FF0000; margin-top:4pt; margin-bottom:4pt;" value="Cliente Negativado (em <%=formata_data(rs("spc_negativado_data"))%>)" size="30">
		</p>
	</td>
	<% else %>
	<td width="100%" align="left">
		<p class="R">FORMA PELA QUAL CONHECEU A DIS</p>
		<p class="C">
		<input id="midia" name="midia" class="TA" value="<%=s_codigo_e_descricao%>" style="width:310px;" readonly />
		</p>
	</td>
	<% end if %>
	</tr>
</table>

<!-- ************   INDICADOR   ************ -->
<%	s_codigo = Trim("" & rs("indicador"))
	s_descricao = ""
	s_codigo_e_descricao = ""
	if s_codigo <> "" then s_descricao = x_orcamentista_e_indicador(s_codigo)
	if (s_codigo <> "") And (s_descricao <> "") then s_codigo_e_descricao = s_codigo & " - " & s_descricao
%>
<table width="649" class="QS" cellspacing="0">
	<tr>
	<td width="100%" align="left">
		<p class="R">INDICADOR</p>
		<p class="C">
		<input id="indicador" name="indicador" class="TA" value="<%=s_codigo_e_descricao%>" style="width:620px;" readonly />
		</p>
	</td>
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
		s_codigo = ""
		s_descricao = ""
		s_codigo_e_descricao = ""
		if Not tRefBancaria.Eof then 
			strRefBancariaBanco=Trim("" & tRefBancaria("banco"))
			strRefBancariaAgencia=Trim("" & tRefBancaria("agencia"))
			strRefBancariaConta=Trim("" & tRefBancaria("conta"))
			strRefBancariaDdd=Trim("" & tRefBancaria("ddd"))
			strRefBancariaTelefone=Trim("" & tRefBancaria("telefone"))
			strRefBancariaContato=Trim("" & tRefBancaria("contato"))
			s_codigo = strRefBancariaBanco
			if s_codigo <> "" then s_descricao = x_banco(s_codigo)
			if (s_codigo <> "") And (s_descricao <> "") then s_codigo_e_descricao = s_codigo & " - " & s_descricao
			end if
		%>
<br>
<% if Not eh_cpf then %>
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
							<input id="c_RefBancariaBanco" name="c_RefBancariaBanco" class="TA" value="<%=s_codigo_e_descricao%>" style="width:620px;" readonly />
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
							<input name="c_RefBancariaAgencia" id="c_RefBancariaAgencia" class="TA" maxlength="8" size="12" value="<%=strRefBancariaAgencia%>" readonly>
						</p>
					</td>
					<td class="MD" align="left">
						<p class="R">CONTA</p>
						<p class="C">
							<input name="c_RefBancariaConta" id="c_RefBancariaConta" class="TA" maxlength="12" value="<%=strRefBancariaConta%>" readonly>
						</p>
					</td>
					<td class="MD" align="left">
						<p class="R">DDD</p>
						<p class="C">
							<input name="c_RefBancariaDdd" id="c_RefBancariaDdd" class="TA" maxlength="2" size="4" value="<%=strRefBancariaDdd%>" readonly>
						</p>
					</td>
					<td align="left">
						<p class="R">TELEFONE</p>
						<p class="C">
							<input name="c_RefBancariaTelefone" id="c_RefBancariaTelefone" class="TA" maxlength="9" value="<%=telefone_formata(strRefBancariaTelefone)%>" readonly>
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
							<input name="c_RefBancariaContato" id="c_RefBancariaContato" class="TA" maxlength="40"  style="width:600px;" value="<%=strRefBancariaContato%>" readonly>
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
<!--  ASSEGURA CRIAÇÃO DE UM ARRAY DE CAMPOS, MESMO QUANDO HOUVER SOMENTE 1 LINHA!! -->
<input type="hidden" name="c_RefProfNomeEmpresa" id="c_RefProfNomeEmpresa" value="">
<input type="hidden" name="c_RefProfCargo" id="c_RefProfCargo" value="">
<input type="hidden" name="c_RefProfDdd" id="c_RefProfDdd" value="">
<input type="hidden" name="c_RefProfTelefone" id="c_RefProfTelefone" value="">
<input type="hidden" name="c_RefProfPeriodoRegistro" id="c_RefProfPeriodoRegistro" value="">
<input type="hidden" name="c_RefProfRendimentos" id="c_RefProfRendimentos" value="">
<input type="hidden" name="c_RefProfCnpj" id="c_RefProfCnpj" value="">
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
<br>
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
							<input name="c_RefProfNomeEmpresa" id="c_RefProfNomeEmpresa" class="TA" maxlength="60"  style="width:450px;" value="<%=strRefProfNomeEmpresa%>" readonly>
						</p>
					</td>
					<td class="MC" align="left">
						<p class="R">CNPJ</p>
						<p class="C">
							<input name="c_RefProfCnpj" id="c_RefProfCnpj" class="TA" maxlength="18" size="24" value="<%=strRefProfCnpj%>" readonly>
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
							<input name="c_RefProfCargo" id="c_RefProfCargo" class="TA" maxlength="40" style="width:350px;" value="<%=strRefProfCargo%>" readonly>
						</p>
					</td>
					<td class="MD" align="left">
						<p class="R">DDD</p>
						<p class="C">
							<input name="c_RefProfDdd" id="c_RefProfDdd" class="TA" maxlength="2" size=4 value="<%=strRefProfDdd%>" readonly>
						</p>
					</td>
					<td align="left">
						<p class="R">TELEFONE</p>
						<p class="C">
							<input name="c_RefProfTelefone" id="c_RefProfTelefone" class="TA" maxlength="9" value="<%=telefone_formata(strRefProfTelefone)%>" readonly>
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
							<input name="c_RefProfPeriodoRegistro" id="c_RefProfPeriodoRegistro" class="TA" maxlength="10" value="<%=strRefProfPeriodoRegistro%>" readonly>
						</p>
					</td>
					<td width="50%" align="left">
						<p class="R">RENDIMENTOS (<%=SIMBOLO_MONETARIO%>)</p>
						<p class="C">
							<input name="c_RefProfRendimentos" id="c_RefProfRendimentos" class="TA" maxlength="18" value="<%=strRefProfRendimentos%>" readonly>
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
<input type="hidden" name="c_RefComercialNomeEmpresa" id="c_RefComercialNomeEmpresa" value="">
<input type="hidden" name="c_RefComercialContato" id="c_RefComercialContato" value="">
<input type="hidden" name="c_RefComercialDdd" id="c_RefComercialDdd" value="">
<input type="hidden" name="c_RefComercialTelefone" id="c_RefComercialTelefone" value="">
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
							<input name="c_RefComercialNomeEmpresa" id="c_RefComercialNomeEmpresa" class="TA" maxlength="60"  style="width:600px;" value="<%=strRefComercialNomeEmpresa%>" readonly>
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
							<input name="c_RefComercialContato" id="c_RefComercialContato" class="TA" maxlength="40" value="<%=strRefComercialContato%>" readonly>
						</p>
					</td>
					<td class="MD" align="left">
						<p class="R">DDD</p>
						<p class="C">
							<input name="c_RefComercialDdd" id="c_RefComercialDdd" class="TA" maxlength="2" size="4" value="<%=strRefComercialDdd%>" readonly>
						</p>
					</td>
					<td align="left">
						<p class="R">TELEFONE</p>
						<p class="C">
							<input name="c_RefComercialTelefone" id="c_RefComercialTelefone" class="TA" maxlength="9" value="<%=telefone_formata(strRefComercialTelefone)%>" readonly>
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


<!-- ************   PJ: DADOS DO SÓCIO MAJORITÁRIO   ************ -->
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
					<td class="MC MD" width="85%" align="left">
						<p class="R">NOME</p>
						<p class="C">
						<input id="c_SocioMajNome" name="c_SocioMajNome" class="TA" value="<%=Trim("" & rs("SocMaj_Nome"))%>" maxlength="60" size="61" readonly>
						</p>
					</td>
					<td class="MC" align="left">
						<p class="R">CPF</p>
						<p class="C">
						<input id="c_SocioMajCpf" name="c_SocioMajCpf" class="TA" value="<%=cnpj_cpf_formata(Trim("" & rs("SocMaj_CPF")))%>" maxlength="14" size="15" readonly>
						</p>
					</td>
				</tr>
			</table>
		</td>
	</tr>

<%	s_codigo = Trim("" & rs("SocMaj_banco"))
	s_descricao = ""
	s_codigo_e_descricao = ""
	if s_codigo <> "" then s_descricao = x_banco(s_codigo)
	if (s_codigo <> "") And (s_descricao <> "") then s_codigo_e_descricao = s_codigo & " - " & s_descricao
%>
	<tr>
		<td width="100%" align="left">
			<table width="100%" class="QS" cellspacing="0">
				<tr>
					<td width="100%" align="left">
						<p class="R">BANCO</p>
						<p class="C">
							<input id="c_SocioMajBanco" name="c_SocioMajBanco" class="TA" value="<%=s_codigo_e_descricao%>" style="width:620px;" readonly />
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
							<input name="c_SocioMajAgencia" id="c_SocioMajAgencia" class="TA" maxlength="8" value="<%=Trim("" & rs("SocMaj_agencia"))%>" readonly>
						</p>
					</td>
					<td class="MD" align="left">
						<p class="R">CONTA</p>
						<p class="C">
							<input name="c_SocioMajConta" id="c_SocioMajConta" class="TA" maxlength="12" value="<%=Trim("" & rs("SocMaj_conta"))%>" readonly>
						</p>
					</td>
					<td class="MD" align="left">
						<p class="R">DDD</p>
						<p class="C">
							<input name="c_SocioMajDdd" id="c_SocioMajDdd" class="TA" maxlength="2" size="4" value="<%=Trim("" & rs("SocMaj_ddd"))%>" readonly>
						</p>
					</td>
					<td align="left">
						<p class="R">TELEFONE</p>
						<p class="C">
							<input name="c_SocioMajTelefone" id="c_SocioMajTelefone" class="TA" maxlength="9" value="<%=telefone_formata(Trim("" & rs("SocMaj_telefone")))%>" readonly>
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
							<input name="c_SocioMajContato" id="c_SocioMajContato" class="TA" maxlength="40"  style="width:600px;" value="<%=Trim("" & rs("SocMaj_contato"))%>" readonly>
						</p>
					</td>
				</tr>
			</table>
		</td>
	</tr>
</table>
<%end if%>


<% if (ocultar_botoes = "S") then %>
<br />
<br />
<% elseif (pedido_selecionado_inicial = "") And (orcamento_selecionado_inicial = "") then %>
<!-- ************   SEPARADOR   ************ -->
<table width="698" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc" align="left">&nbsp;</td></tr>
</table>
<br />

<table class="notPrint" width="649" cellspacing="0">
<tr>
	<td align="center">
			<% s_url = "javascript:history.back();" %>
			<a name="bVOLTAR" id="bVOLTAR" href="<%=s_url%>" title="volta para a página anterior">
			<img src="../botao/voltar.gif" width="176" height="55" border="0"></a>
	</td>
</tr>
</table>
<% else %>
<!-- ************   SEPARADOR   ************ -->
<table width="698" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc" align="left">&nbsp;</td></tr>
</table>
<br />

<table class="notPrint" width="649" cellspacing="0">
<tr>
	<td align="center">
			<% if pagina_retorno <> "" then
				s_url = pagina_retorno
			else
				s_url = "PedidoConsultaView.asp" & "?pedido_selecionado=" & pedido_selecionado_inicial & "&pedido_selecionado_inicial=" & pedido_selecionado_inicial & "&usuario=" & usuario & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))
				end if %>
			<a name="bVOLTAR" id="bVOLTAR" href="<%=s_url%>" title="volta para a página anterior">
			<img src="../botao/voltar.gif" width="176" height="55" border="0"></a>
	</td>
</tr>
</table>
<% end if %>
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