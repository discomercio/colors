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
'	  FinCadContaCorrenteEdita.asp
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
	dim s, strSql, usuario, id_selecionado, operacao_selecionada, s_cor
	usuario = trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 
	
'	REGISTRO A EDITAR
	id_selecionado = trim(request("id_selecionado"))
	operacao_selecionada = trim(request("operacao_selecionada"))
	
	if operacao_selecionada=OP_INCLUI then
		id_selecionado=retorna_so_digitos(id_selecionado)
		end if

	if (id_selecionado="") Or (converte_numero(id_selecionado)=0) then Response.Redirect("aviso.asp?id=" & ERR_ID_NAO_INFORMADO) 
	if (operacao_selecionada<>OP_INCLUI) And (operacao_selecionada<>OP_CONSULTA) then Response.Redirect("aviso.asp?id=" & ERR_OPERACAO_NAO_ESPECIFICADA)

'	CONECTA COM O BANCO DE DADOS
	dim cn,rs
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	strSql = "SELECT " & _
				"*" & _
			" FROM t_FIN_CONTA_CORRENTE" & _
			" WHERE" & _
				" (id = " & id_selecionado & ")"
	set rs = cn.Execute(strSql)
	if Err <> 0 then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
	
	if operacao_selecionada=OP_INCLUI then
		if Not rs.EOF then Response.Redirect("aviso.asp?id=" & ERR_ID_JA_CADASTRADO)
	elseif operacao_selecionada=OP_CONSULTA then
		if rs.EOF then Response.Redirect("aviso.asp?id=" & ERR_ID_NAO_CADASTRADO)
		end if
	
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

<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
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

function AtualizaRegistro( f ) {
	if (trim(f.c_descricao.value)=="") {
		alert('Preencha a descrição!!');
		f.c_descricao.focus();
		return;
		}
	if (trim(f.c_banco.value)=="") {
		alert('Selecione o banco!!');
		f.c_banco.focus();
		return;
		}
	if (trim(f.c_agencia_sem_digito.value)=="") {
		alert('Informe a agência!!');
		f.c_agencia_sem_digito.focus();
		return;
		}
	if (trim(f.c_conta_sem_digito.value)=="") {
		alert('Informe o número da conta!!');
		f.c_conta_sem_digito.focus();
		return;
		}
	if (!isDate(f.c_dt_saldo_inicial)) {
		alert("Data inválida!!");
		f.c_dt_saldo_inicial.focus();
		return;
		}
	if (trim(f.c_vl_saldo_inicial.value)=="") {
		alert("Informe o valor!!");
		f.c_vl_saldo_inicial.focus();
		return;
		}
//  PARA O CASO DE TER CLICADO NO BOTÃO BACK APÓS TER CLICADO NA OPERAÇÃO EXCLUIR
	f.operacao_selecionada.value=f.operacao_selecionada_original.value;
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

<link href="<%=URL_FILE__E_CSS%>" rel="stylesheet" type="text/css">
<link href="<%=URL_FILE__EPRINTER_CSS%>" rel="stylesheet" type="text/css" media="print">

<style type="text/css">
#rb_st_ativo {
	margin: 0pt 4pt 1pt 6pt;
	}
</style>


<%	if operacao_selecionada=OP_INCLUI then
		s = "fCAD.c_descricao.focus()"
	else
		s = "focus()"
		end if
%>
<body onload="<%=s%>">
<center>



<!--  FORMULÁRIO DE CADASTRO  -->

<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
<%	if operacao_selecionada=OP_INCLUI then
		s = "Cadastro de Nova Conta Corrente"
	else
		s = "Consulta/Edição de Conta Corrente"
		end if
%>
	<td align="center" valign="bottom"><p class="PEDIDO"><%=s%><br><span class="C">&nbsp;</span></p></td>
</tr>
</table>
<br>


<!--  CAMPOS DO CADASTRO  -->
<form id="fCAD" name="fCAD" method="post" action="FinCadContaCorrenteAtualiza.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name='operacao_selecionada_original' id="operacao_selecionada_original" value='<%=operacao_selecionada%>'>
<input type="hidden" name='operacao_selecionada' id="operacao_selecionada" value='<%=operacao_selecionada%>'>

<!-- ************   ID / DESCRIÇÃO   ************ -->
<table width="649" class="Q" cellspacing="0">
	<tr>
		<td class="MD" align="center" width="15%">
			<p class="R">ID</p>
			<p class="C">
				<input id="id_selecionado" name="id_selecionado" class="TA" value="<%=id_selecionado%>" readonly size="6" style="text-align:center; color:#0000ff">
			</p>
		</td>
		<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("descricao")) else s=""%>
		<td width="85%">
			<p class="R">DESCRIÇÃO</p>
			<p class="C">
				<input id="c_descricao" name="c_descricao" class="TA" type="text" maxlength="30" size="60" value="<%=s%>" 
					onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCAD.c_agencia_sem_digito.focus(); filtra_nome_identificador();">
			</p>
		</td>
	</tr>
</table>

<table width="649" class="QS" cellspacing="0">
	<tr>
		<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("banco")) else s=""%>
		<td colspan="4">
			<p class="R">BANCO</p>
			<p class="C">
				<select name="c_banco" id="c_banco" style="margin:1px 10px 6px 10px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
				<%=banco_monta_itens_select(s) %>
				</select>
			</p>
		</td>
	</tr>
	<tr>
		<td class="MC MD">
			<p class="R">AGÊNCIA</p>
			<p class="C">
				<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("agencia_sem_digito")) else s=""%>
				<input name="c_agencia_sem_digito" id="c_agencia_sem_digito" class="TA" maxlength="7" value="<%=s%>" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fCAD.c_digito_agencia.focus(); filtra_agencia_bancaria_sem_digito();">
			</p>
		</td>
		<td class="MC MD">
			<p class="R">DÍGITO</p>
			<p class="C">
				<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("digito_agencia")) else s=""%>
				<input name="c_digito_agencia" id="c_digito_agencia" class="TA" maxlength="1" style="width:40px;" value="<%=s%>" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fCAD.c_conta_sem_digito.focus(); filtra_numerico();">
			</p>
		</td>
		<td class="MC MD">
			<p class="R">CONTA</p>
			<p class="C">
				<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("conta_sem_digito")) else s=""%>
				<input name="c_conta_sem_digito" id="c_conta_sem_digito" class="TA" maxlength="11" value="<%=s%>" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fCAD.c_digito_conta.focus(); filtra_conta_bancaria_sem_digito();">
			</p>
		</td>
		<td class="MC">
			<p class="R">DÍGITO</p>
			<p class="C">
				<%if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("digito_conta")) else s=""%>
				<input name="c_digito_conta" id="c_digito_conta" class="TA" maxlength="1" style="width:40px;" value="<%=s%>" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fCAD.c_dt_saldo_inicial.focus(); filtra_numerico();">
			</p>
		</td>
	</tr>
</table>


<!-- ************   STATUS ATIVO   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
<%
	dim st_ativo
	st_ativo=false
	if operacao_selecionada=OP_CONSULTA then
		if Cstr(rs("st_ativo")) = Cstr(COD_FIN_ST_ATIVO__ATIVO) then st_ativo=true
	elseif operacao_selecionada=OP_INCLUI then
		st_ativo=true
		end if
%>
		<td width="25%" class="MD">
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
		<td class="MD" valign="top">
			<p class="R">DATA SALDO INICIAL</p>
			<p class="C">
				<%if operacao_selecionada=OP_CONSULTA then s=formata_data("" & rs("dt_saldo_inicial")) else s=""%>
				<input class="TA" maxlength="10" style="width:110px;text-align:left;" 
					name="c_dt_saldo_inicial" id="c_dt_saldo_inicial" 
					onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fCAD.c_vl_saldo_inicial.focus(); filtra_data();" 
					onblur="if (!isDate(this)) {alert('Data inválida!!'); this.focus();}" 
					value="<%=s%>">
			</p>
		</td>
		<td valign="top">
			<p class="R">VL SALDO INICIAL (<%=SIMBOLO_MONETARIO%>)</p>
			<p class="C">
				<%
					if operacao_selecionada=OP_CONSULTA then s=formata_moeda("" & rs("vl_saldo_inicial")) else s=""
					if converte_numero(s) < 0 then s_cor="red" else s_cor="green"
				%>
				<input class="TA" maxlength="12" style="width:140px;text-align:left;color:<%=s_cor%>;" 
					name="c_vl_saldo_inicial" id="c_vl_saldo_inicial" 
					onkeypress="if (digitou_enter(true)&&tem_info(this.value)) bATUALIZA.focus(); filtra_moeda();" 
					onblur="this.value=formata_moeda(this.value); this.style.color=(converte_numero(this.value)<0?'red':'green');" 
					value="<%=s%>">
			</p>
		</td>
	</tr>
</table>


<!-- ************   SEPARADOR   ************ -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>
<br>


<table class="notPrint" width="649" cellspacing="0">
<tr>
	<td><a href="javascript:history.back()" title="cancela as alterações no cadastro">
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