<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp"        -->
<!-- #include file = "../global/Global.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%

'     ===========================================================
'	  M E N U O R C A M E N T I S T A E I N D I C A D O R . A S P
'     ===========================================================
'
'
'	  S E R V E R   S I D E   S C R I P T I N G
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
'						I N I C I A L I Z A     P Á G I N A     A S P
' _____________________________________________________________________________________________


	On Error GoTo 0
	Err.Clear
	
'	OBTEM USUÁRIO
	dim s, usuario, usuario_nome, loja, loja_nome, intIdx, strDisabled
	usuario = Trim(Session("usuario_atual"))
	usuario_nome = Trim(Session("usuario_nome_atual"))
	loja = Trim(Session("loja_atual"))
	loja_nome = Session("loja_nome_atual")
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 
	If (loja = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))
	
	if (Not operacao_permitida(OP_LJA_CONS_CAD_ORCAMENTISTAS_E_INDICADORES, s_lista_operacoes_permitidas)) And _
	   (Not operacao_permitida(OP_LJA_EDITA_CAD_ORCAMENTISTAS_E_INDICADORES, s_lista_operacoes_permitidas)) And _
	   (Not operacao_permitida(OP_LJA_CADASTRAMENTO_INDICADOR, s_lista_operacoes_permitidas)) then 
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

'	CONECTA COM O BANCO DE DADOS
	dim cn
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	
	'RECUPERA OPÇÃO MEMORIZADA
	dim ordenacao_default
	ordenacao_default = get_default_valor_texto_bd(usuario, "MenuOrcamentistaEIndicador|ordenacao")
%>





<%
'	  C L I E N T   S I D E   S C R I P T I N G
'
'      CCCCCCC   LLL        IIIII  EEEEEEEEE  NNN   NNN  TTTTTTTTT EEEEEEEEE
'     CCC   CCC  LLL         III   EEE        NNNN  NNN     TTT    EEE
'     CCC        LLL         III   EEE        NNNNN NNN     TTT    EEE
'     CCC        LLL         III   EEEEEE     NNN NNNNN     TTT    EEEEEE
'     CCC        LLL         III   EEE        NNN  NNNN     TTT    EEE
'     CCC   CCC  LLL   LLL   III   EEE        NNN   NNN     TTT    EEE
'      CCCCCCC   LLLLLLLLL  IIIII  EEEEEEEEE  NNN   NNN     TTT    EEEEEEEEE
%>

<html>


<head>
	<title>LOJA</title>
	</head>

<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
function fOPConcluir( f ){
var s_dest, s_op, s_id_selecionado, intIdx;
	
	s_dest="";
	s_op="";
	s_id_selecionado="";
	intIdx = -1;
	
	intIdx++;
	if (f.rb_op[intIdx].checked) {
		s_dest="OrcamentistaEIndicadorNovo.asp";
		s_op=OP_INCLUI;
		s_id_selecionado=f.c_novo.value;
		if (trim(f.c_novo.value)=="") {
			alert("Forneça a identificação para o novo orçamentista!!");
			f.c_novo.focus();
			return false;
			}
		if ((!f.rb_tipo[0].checked)&&(!f.rb_tipo[1].checked)) {
			alert("Informe se o novo orçamentista / indicador é PF ou PJ!!");
			return false;
			}
		}
	
	intIdx++;
	if (f.rb_op[intIdx].checked) {
		s_dest="OrcamentistaEIndicadorConsulta.asp";
		s_op=OP_CONSULTA;
		s_id_selecionado=f.c_cons.value;
		if (trim(f.c_cons.value)=="") {
			alert("Forneça a identificação do orçamentista a ser consultado!!");
			f.c_cons.focus();
			return false;
			}
		}

	intIdx++;

	if (f.rb_op[intIdx].checked) {
	    s_dest="OrcamentistaEIndicadorEdita.asp";
	    s_op=OP_CONSULTA;
	    s_id_selecionado=f.c_edit.value;
	    if (f.c_edit.value=="") {
	        alert("Forneça a identificação do orçamentista a ser editado!!");
	        f.c_edit.focus();
	        return false;
	    }
	}


	intIdx++;
	if (f.rb_op[intIdx].checked) {
		s_dest="OrcamentistaEIndicadorLista.asp?op=A";
		}

	intIdx++;
	if (f.rb_op[intIdx].checked) {
		s_dest="OrcamentistaEIndicadorLista.asp?op=I";
		}

	intIdx++;
	if (f.rb_op[intIdx].checked) {
		s_dest="OrcamentistaEIndicadorLista.asp?op=T";
		}

	intIdx++;
	if (f.rb_op[intIdx].checked) {
		if (trim(f.vendedor.value)=='') {
			alert('Selecione o vendedor!!');
			f.vendedor.focus();
			return;
			}
		s_dest="OrcamentistaEIndicadorAssocAoVendedor.asp";
		}

	if (s_dest=="") {
		alert("Escolha uma das opções!!");
		return false;
		}
	
	f.id_selecionado.value=s_id_selecionado;
	f.operacao_selecionada.value=s_op;
	
	f.action=s_dest;
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



<body onload="focus();">

<!--  MENU SUPERIOR -->  
<table width="100%" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">

<tr>
	<td align="RIGHT" vAlign="BOTTOM"><p class="PEDIDO"><% = loja_nome & " (" & loja & ")" %><br>
	<%	s = usuario_nome
		if s = "" then s = usuario
		s = x_saudacao & ", " & s
		s = "<span class='Cd'>" & s & "</span><br>"
	%>
	<%=s%>
	<span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="senha.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="altera a senha atual do usuário" class="LAlteraSenha">altera senha</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></p></td>
	</tr>

</table>


<center>
<!--  ***********************************************************************************************  -->
<!--  F O R M U L Á R I O                         												       -->
<!--  ***********************************************************************************************  -->
<form METHOD="POST" id="fOP" name="fOP" OnSubmit="if (!fOPConcluir(fOP)) return false">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<INPUT type="hidden" name='id_selecionado' id="id_selecionado" value=''>
<INPUT type="hidden" name='operacao_selecionada' id="operacao_selecionada" value=''>
<input type="hidden" name="url_origem" id="url_origem" value="MenuOrcamentistaEIndicador.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" />

<br />
<span class="T">CADASTRO DE ORÇAMENTISTAS / INDICADORES</span>
<div class="QFn" align="CENTER">
<table class="TFn">
	<tr>
		<td NOWRAP>
			<% intIdx = 0 %>
			
			<%	strDisabled=""
				if Not operacao_permitida(OP_LJA_CADASTRAMENTO_INDICADOR, s_lista_operacoes_permitidas) then strDisabled = " disabled tabindex=-1"
			%>
			<% intIdx = intIdx+1 %>
			<input type="radio" id="rb_op" name="rb_op" value="<%=CStr(intIdx)%>" class="CBOX" onclick="fOP.c_novo.focus()"<%=strDisabled%>><span style="cursor:default;" onclick="fOP.rb_op[<%=CStr(intIdx-1)%>].click();fOP.c_novo.focus();"<%=strDisabled%>>Cadastrar Indicador</span>&nbsp;
				<input name="c_novo" id="c_novo" type="text" maxlength="<%=MAX_TAMANHO_ID_ORCAMENTISTA_E_INDICADOR%>"" size="18" onblur="this.value=trim(this.value);" onclick="fOP.rb_op[<%=CStr(intIdx-1)%>].click();" onkeypress="this.click(); if (digitou_enter(true) && tem_info(this.value)) if (fOPConcluir(fOP)) fOP.submit(); filtra_nome_identificador();"
				<%=strDisabled%>><span style='width:12px;'></span><input type="radio" id="rb_tipo" name="rb_tipo" value="<%=ID_PF%>" class="CBOX" onclick="fOP.rb_op[<%=CStr(intIdx-1)%>].click();"<%=strDisabled%>><span class="rbLink" onclick="fOP.rb_tipo[0].click(); fOP.rb_op[<%=CStr(intIdx-1)%>].click();"
				<%=strDisabled%>><%=ID_PF%></span>
				<input type="radio" id="rb_tipo" name="rb_tipo" value="<%=ID_PJ%>" class="CBOX" onclick="fOP.rb_op[<%=CStr(intIdx-1)%>].click();"<%=strDisabled%>><span class="rbLink" onclick="fOP.rb_tipo[1].click(); fOP.rb_op[<%=CStr(intIdx-1)%>].click();"
				<%=strDisabled%>><%=ID_PJ%></span>
				<br>
				
			<%	strDisabled=""
				if (Not operacao_permitida(OP_LJA_CONS_CAD_ORCAMENTISTAS_E_INDICADORES, s_lista_operacoes_permitidas)) And (Not operacao_permitida(OP_LJA_EDITA_CAD_ORCAMENTISTAS_E_INDICADORES, s_lista_operacoes_permitidas)) then strDisabled = " disabled tabindex=-1"
			%>
			<% intIdx = intIdx+1 %>
			<input type="radio" id="rb_op" name="rb_op" value="<%=CStr(intIdx)%>" class="CBOX" onclick="fOP.c_cons.focus()"<%=strDisabled%>><span style="cursor:default" onclick="fOP.rb_op[<%=CStr(intIdx-1)%>].click();fOP.c_cons.focus();">Consultar</span>&nbsp;
				<input name="c_cons" id="c_cons" type="text" maxlength="<%=MAX_TAMANHO_ID_ORCAMENTISTA_E_INDICADOR%>" size="18" onblur="this.value=trim(this.value);" onclick="fOP.rb_op[<%=CStr(intIdx-1)%>].click()" onkeypress="this.click(); if (digitou_enter(true) && tem_info(this.value)) if (fOPConcluir(fOP)) fOP.submit(); filtra_nome_identificador();"<%=strDisabled%>><br>
			
			<% strDisabled=""
			if not operacao_permitida(OP_LJA_EDITA_CAD_ORCAMENTISTAS_E_INDICADORES, s_lista_operacoes_permitidas) then strDisabled=" disabled tabindex=-1" %>
			<% intIdx = intIdx+1 %>
			<input type="radio" id="rb_op" name="rb_op" value="<%=CStr(intIdx)%>" class="CBOX" onclick="fOP.c_edit.focus()"<%=strDisabled%>><span style="cursor:default" onclick="fOP.rb_op[<%=CStr(intIdx-1)%>].click();fOP.c_edit.focus();">Editar</span>&nbsp;
				<input name="c_edit" id="c_edit" type="text" maxlength="<%=MAX_TAMANHO_ID_ORCAMENTISTA_E_INDICADOR%>" size="18" onblur="this.value=trim(this.value);" onclick="fOP.rb_op[<%=CStr(intIdx-1)%>].click()" onkeypress="this.click(); if (digitou_enter(true) && tem_info(this.value)) if (fOPConcluir(fOP)) fOP.submit(); filtra_nome_identificador();"<%=strDisabled%>><br>
			
			<% strDisabled=""
				if (Not operacao_permitida(OP_LJA_CONS_CAD_ORCAMENTISTAS_E_INDICADORES, s_lista_operacoes_permitidas)) And (Not operacao_permitida(OP_LJA_EDITA_CAD_ORCAMENTISTAS_E_INDICADORES, s_lista_operacoes_permitidas)) then strDisabled = " disabled tabindex=-1"
			%>
			<% intIdx = intIdx+1 %>
			<input type="radio" id="rb_op" name="rb_op" value="<%=CStr(intIdx)%>" class="CBOX"<%=strDisabled%>><span class="rbLink" onclick="fOP.rb_op[<%=CStr(intIdx-1)%>].click(); fOP.bEXECUTAR.click();"<%=strDisabled%>>Consultar Ativos</span><br>
			
			<% strDisabled=""
			 if Not operacao_permitida(OP_LJA_CONSULTA_UNIVERSAL_PEDIDO_ORCAMENTO, s_lista_operacoes_permitidas) then strDisabled=" disabled tabindex=-1" %>
			<% intIdx = intIdx+1 %>
			<input type="radio" id="rb_op" name="rb_op" value="<%=CStr(intIdx)%>" class="CBOX"<%=strDisabled%>><span class="rbLink" onclick="fOP.rb_op[<%=CStr(intIdx-1)%>].click(); fOP.bEXECUTAR.click();"<%=strDisabled%>>Consultar Inativos</span><br>

			<% strDisabled=""
			 if Not operacao_permitida(OP_LJA_CONSULTA_UNIVERSAL_PEDIDO_ORCAMENTO, s_lista_operacoes_permitidas) then strDisabled=" disabled tabindex=-1" %>
			<% intIdx = intIdx+1 %>
			<input type="radio" id="rb_op" name="rb_op" value="<%=CStr(intIdx)%>" class="CBOX"<%=strDisabled%>><span class="rbLink" onclick="fOP.rb_op[<%=CStr(intIdx-1)%>].click(); fOP.bEXECUTAR.click();"<%=strDisabled%>>Consultar Ativos e Inativos</span><br>
			
			<% intIdx = intIdx+1 %>
			<input type="radio" id="rb_op" name="rb_op" value="<%=CStr(intIdx)%>" class="CBOX"<%=strDisabled%>><span class="rbLink" onclick="fOP.rb_op[<%=CStr(intIdx-1)%>].click(); fOP.vendedor.focus();"<%=strDisabled%>>Associados ao Vendedor</span>
			<br>
				<select id="vendedor" name="vendedor" style="margin-left:25px;" onclick="fOP.rb_op[<%=CStr(intIdx-1)%>].click();"<%=strDisabled%>>
				  <% =vendedor_do_indicador_desta_loja_monta_itens_select(loja, Null) %>
				</select>
			</td>
		</tr>
	</table>

	<table width="100%" cellpadding="0" cellspacing="0"><tr><td class="MC" style="height:8px;"></td></tr></table>

	<table cellspacing="0" cellpadding="0" border="0">
		<tr>
			<td><span class="Lbl">Ordenação</span></td>
		</tr>
		<tr>
			<td>
				<select id="ordenacao" name="ordenacao" style="min-width:140px;">
					<% =ordenacao_lista_indicadores_monta_itens_select(ordenacao_default) %>
				</select>
			</td>
		</tr>
	</table>
	
	<table width="100%" cellpadding="0" cellspacing="0"><tr><td class="MB" style="height:12px;"></td></tr></table>

	<p class="R" style="margin: 0 10 0 10">&nbsp;</p>
	<input name="bEXECUTAR" id="bEXECUTAR" type="submit" class="Botao" value="EXECUTAR" title="executa">
	<p class="R" style="margin: 0 10 0 10">&nbsp;</p>
</div>
</form>

<p class="TracoBottom"></p>

<table cellSpacing="0">
<tr>
	<td align="center"><a href="MenuFuncoesAdministrativas.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="volta para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
</tr>
</table>

</center>

</body>
</html>


<%
'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>