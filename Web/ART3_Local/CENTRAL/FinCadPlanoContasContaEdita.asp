<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =====================================
'	  FinCadPlanoContasContaEdita.asp
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
	dim s, strSql, usuario, id_selecionado, operacao_selecionada, s_natureza, blnStSistema
	usuario = trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 
	
'	REGISTRO A EDITAR
	id_selecionado = trim(request("id_selecionado"))
	operacao_selecionada = trim(request("operacao_selecionada"))
	s_natureza = Trim(Request.Form("rb_natureza"))
	
	if operacao_selecionada=OP_INCLUI then
		id_selecionado=retorna_so_digitos(id_selecionado)
		end if

	id_selecionado=normaliza_codigo(id_selecionado, TAM_PLANO_CONTAS__CONTA)
	
	if (id_selecionado="") Or (converte_numero(id_selecionado)=0) then Response.Redirect("aviso.asp?id=" & ERR_ID_NAO_INFORMADO) 
	if (operacao_selecionada<>OP_INCLUI) And (operacao_selecionada<>OP_CONSULTA) then Response.Redirect("aviso.asp?id=" & ERR_OPERACAO_NAO_ESPECIFICADA)
	if (Cstr(s_natureza)<>Cstr(COD_FIN_NATUREZA__CREDITO)) And (Cstr(s_natureza)<>Cstr(COD_FIN_NATUREZA__DEBITO)) then Response.Redirect("aviso.asp?id=" & ERR_FIN_NATUREZA_OPERACAO_NAO_ESPECIFICADO)

'	CONECTA COM O BANCO DE DADOS
	dim cn,rs,rs2
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	strSql = "SELECT " & _
				"*" & _
			" FROM t_FIN_PLANO_CONTAS_CONTA" & _
			" WHERE" & _
				" (id = " & id_selecionado & ")" & _
				" AND (natureza = '" & s_natureza & "')"
	set rs = cn.Execute(strSql)
	if Err <> 0 then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)

	blnStSistema = False
	if operacao_selecionada=OP_INCLUI then
		if Not rs.EOF then Response.Redirect("aviso.asp?id=" & ERR_ID_JA_CADASTRADO)
	elseif operacao_selecionada=OP_CONSULTA then
		if rs.EOF then Response.Redirect("aviso.asp?id=" & ERR_ID_NAO_CADASTRADO)
	'	É UMA CONTA USADA PELO SISTEMA P/ EFETUAR LANÇAMENTOS AUTOMÁTICOS?
		if Cstr(rs("st_sistema")) = Cstr(COD_FIN_ST_SISTEMA__SIM) then blnStSistema = True
		end if

'	UMA CONTA PODE SER CADASTRADA PARA CRÉDITO E PARA DÉBITO
'	VERIFICA SE JÁ EXISTE UMA CADASTRADA E COMPARA SE O GRUPO DE CONTA DE AMBAS ESTÃO
'	DIFERENTES, CASO SIM, EXIBE UM AVISO.
	dim s_natureza_conta_espelhada, s_grupo_conta_espelhada
	s_natureza_conta_espelhada = ""
	s_grupo_conta_espelhada=""
	if Cstr(s_natureza) = Cstr(COD_FIN_NATUREZA__CREDITO) then
		s_natureza_conta_espelhada = Cstr(COD_FIN_NATUREZA__DEBITO)
	elseif Cstr(s_natureza) = Cstr(COD_FIN_NATUREZA__DEBITO) then
		s_natureza_conta_espelhada = Cstr(COD_FIN_NATUREZA__CREDITO)
		end if
	
	strSql = "SELECT " & _
				"*" & _
			" FROM t_FIN_PLANO_CONTAS_CONTA" &_
			" WHERE" & _
				" (id = " & id_selecionado & ")" & _
				" AND (natureza = '" & s_natureza_conta_espelhada & "')"
	set rs2 = cn.Execute(strSql)
	if Err <> 0 then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
	if Not rs2.Eof then
		s_grupo_conta_espelhada = Trim("" & rs2("id_plano_contas_grupo"))
		end if
	
	



' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________

function finPlanoContasGrupoTodosMontaItensSelect(byval id_default)
dim x, r, strResp, ha_default
	id_default = Trim("" & id_default)
	ha_default=False
	set r = cn.Execute("SELECT * FROM t_FIN_PLANO_CONTAS_GRUPO ORDER BY id")
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
		strResp = strResp & normaliza_codigo(x,TAM_PLANO_CONTAS__GRUPO) & " - " & Trim("" & r("descricao"))
		strResp = strResp & "</OPTION>" & chr(13)
		r.MoveNext
		loop

	if Not ha_default then
		strResp = "<OPTION SELECTED VALUE=''>&nbsp;</OPTION>" & chr(13) & strResp
	else
		strResp = "<OPTION VALUE=''>&nbsp;</OPTION>" & chr(13) & strResp
		end if
		
	finPlanoContasGrupoTodosMontaItensSelect = strResp
	r.close
	set r=nothing
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
var i,b;
	if (trim(f.c_descricao.value)=="") {
		alert('Preencha a descrição!!');
		f.c_descricao.focus();
		return;
		}
	b=false;
	for (i=0; i<f.rb_natureza.length; i++) {
		if (f.rb_natureza[i].checked) {
			b=true;
			break;
			}
		}
	if (!b) {
		alert('Informe a natureza da operação!!');
		return;
		}
	if (trim(f.c_grupo.value)=="") {
		alert('Selecione o grupo de contas ao qual esta conta deve ser vinculada!!');
		f.c_grupo.focus();
		return;
		}
	if (trim(f.c_grupo_conta_espelhada.value)!="") {
		if (trim(f.c_grupo_conta_espelhada.value)!=trim(f.c_grupo.value)) {
			if (!confirm("Existe uma conta simétrica a esta cadastrada para a natureza de operação " + f.c_descricao_natureza_conta_espelhada.value + "\nMas a conta simétrica está vinculada a um outro grupo de contas: " + f.c_grupo_conta_espelhada.value + "\nContinua mesmo assim?")) return;
			}
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
	margin: 0pt 2pt 1pt 5pt;
	}
#rb_natureza {
	margin: 0pt 2pt 1pt 5pt;
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

<table width="689" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
<%	if operacao_selecionada=OP_INCLUI then
		s = "Plano de Contas: Cadastro de Nova Conta"
	else
		s = "Plano de Contas: Consulta/Edição de Conta"
		end if
%>
	<td align="center" valign="bottom"><p class="PEDIDO"><%=s%><br><span class="C">&nbsp;</span></p></td>
</tr>
</table>
<br>


<!--  CAMPOS DO CADASTRO  -->
<form id="fCAD" name="fCAD" method="post" action="FinCadPlanoContasContaAtualiza.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name='rb_natureza' id="rb_natureza" value='<%=s_natureza%>'>
<input type="hidden" name='c_grupo_conta_espelhada' id="c_grupo_conta_espelhada" value='<%=normaliza_codigo(s_grupo_conta_espelhada,TAM_PLANO_CONTAS__GRUPO)%>'>
<input type="hidden" name='c_descricao_natureza_conta_espelhada' id="c_descricao_natureza_conta_espelhada" value='<%=finNaturezaDescricao(s_natureza_conta_espelhada)%>'>
<input type="hidden" name='operacao_selecionada_original' id="operacao_selecionada_original" value='<%=operacao_selecionada%>'>
<input type="hidden" name='operacao_selecionada' id="operacao_selecionada" value='<%=operacao_selecionada%>'>


<!-- ************   ID / DESCRIÇÃO   ************ -->
<table width="689" class="Q" cellspacing="0">
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
				<input id="c_descricao" name="c_descricao" class="TA" type="text" maxlength="60" size="85" value="<%=s%>" 
					onkeypress="if (digitou_enter(true) && tem_info(this.value)) bATUALIZA.focus(); filtra_nome_identificador();">
			</p>
		</td>
	</tr>
</table>

<!-- ************   STATUS ATIVO   ************ -->
<table width="689" class="QS" cellspacing="0">
	<tr>
<%
	dim st_ativo
	dim s_grupo
	st_ativo=false
	s_grupo=""
	if operacao_selecionada=OP_CONSULTA then
		if Cstr(rs("st_ativo")) = Cstr(COD_FIN_ST_ATIVO__ATIVO) then st_ativo=true
		s_grupo = Trim("" & rs("id_plano_contas_grupo"))
	elseif operacao_selecionada=OP_INCLUI then
		st_ativo=true
		end if
%>
		<td width="15%" class="MD">
			<p class="R">STATUS</p>
			<p class="C">
				<input type="radio" id="rb_st_ativo" name="rb_st_ativo" 
					value="<%=COD_FIN_ST_ATIVO__INATIVO%>" 
					class="TA" <%if Not st_ativo then Response.Write(" checked")%>
					><span onclick="fCAD.rb_st_ativo[0].click();" 
					style="cursor:default;color:<%=finStAtivoCor(COD_FIN_ST_ATIVO__INATIVO)%>;"
					><%=finStAtivoDescricao(COD_FIN_ST_ATIVO__INATIVO)%></span
					>&nbsp;</p>
			<p class="C">
				<input type="radio" id="rb_st_ativo" name="rb_st_ativo" 
					value="<%=COD_FIN_ST_ATIVO__ATIVO%>" 
					class="TA" <%if st_ativo then Response.Write(" checked")%>
					><span onclick="fCAD.rb_st_ativo[1].click();" 
					style="cursor:default;color:<%=finStAtivoCor(COD_FIN_ST_ATIVO__ATIVO)%>;"
					><%=finStAtivoDescricao(COD_FIN_ST_ATIVO__ATIVO)%></span
					>&nbsp;</p>
		</td>
		<td width="15%" class="MD">
			<p class="R">NATUREZA</p>
			<p class="C">
			<input type="radio" id="rb_natureza" name="rb_natureza" 
				value="<%=COD_FIN_NATUREZA__DEBITO%>" 
				DISABLED
				class="TA" <%if Cstr(s_natureza)=Cstr(COD_FIN_NATUREZA__DEBITO) then Response.Write(" checked")%>
				><span onclick="fCAD.rb_natureza[0].click();" 
				style="cursor:default;color:<%=finNaturezaCor(COD_FIN_NATUREZA__DEBITO)%>;"
				><%=finNaturezaDescricao(COD_FIN_NATUREZA__DEBITO)%></span
				>&nbsp;</p>
			<p class="C">
			<input type="radio" id="rb_natureza" name="rb_natureza" 
				value="<%=COD_FIN_NATUREZA__CREDITO%>" 
				DISABLED
				class="TA" <%if Cstr(s_natureza)=Cstr(COD_FIN_NATUREZA__CREDITO) then Response.Write(" checked")%>
				><span onclick="fCAD.rb_natureza[1].click();" 
				style="cursor:default;color:<%=finNaturezaCor(COD_FIN_NATUREZA__CREDITO)%>;"
				><%=finNaturezaDescricao(COD_FIN_NATUREZA__CREDITO)%></span
				>&nbsp;</p>
		</td>
		<td valign="top">
			<p class="R">GRUPO DE CONTAS</p>
			<select id="c_grupo" name="c_grupo" style="margin-left:4px;margin-top:4px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
			<%=finPlanoContasGrupoTodosMontaItensSelect(s_grupo)%>
			</select>
		</td>
	</tr>
</table>

<% if blnStSistema then %>
	<br>
	<span class="Expl">Obs: este grupo de contas é usado pelo sistema para efetuar lançamentos de modo automático</span>
<% end if %>

<!-- ************   SEPARADOR   ************ -->
<table width="689" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>
<br>


<table class="notPrint" width="689" cellspacing="0">
<tr>
	<td><a href="javascript:history.back()" title="cancela as alterações no cadastro">
		<img src="../botao/cancelar.gif" width="176" height="55" border="0"></a></td>
	<%
	s = ""
	if operacao_selecionada=OP_CONSULTA then
		if Not blnStSistema then
			s = "<td align='CENTER'>" & _
					"<div name='dREMOVE' id='dREMOVE'>" & _
						"<a href='javascript:RemoveRegistro(fCAD)' title='exclui do banco de dados'>" & _
							"<img src='../botao/remover.gif' width=176 height=55 border=0>" & _
						"</a>" & _
					"</div>" & _
				"</td>"
			end if
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