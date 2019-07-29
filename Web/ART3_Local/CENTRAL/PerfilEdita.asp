<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/Global.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ===============================
'	  P E R F I L E D I T A . A S P
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
	
'	OBTEM O ID
	dim s, usuario, perfil_selecionado, operacao_selecionada
	usuario = trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 
		
'	PERFIL A EDITAR
	perfil_selecionado = Ucase(trim(request("perfil_selecionado")))
	operacao_selecionada = trim(request("operacao_selecionada"))
	
	if operacao_selecionada=OP_INCLUI then
		perfil_selecionado=filtra_nome_identificador(perfil_selecionado)
		end if
		
	if perfil_selecionado="" then Response.Redirect("aviso.asp?id=" & ERR_PERFIL_NAO_ESPECIFICADO) 
	if (operacao_selecionada<>OP_INCLUI) And (operacao_selecionada<>OP_CONSULTA) then Response.Redirect("aviso.asp?id=" & ERR_OPERACAO_NAO_ESPECIFICADA)
	
'	CONECTA COM O BANCO DE DADOS
	dim cn, rs, msg_erro
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

	s = "SELECT * FROM t_PERFIL WHERE (apelido='" & perfil_selecionado & "')"
	if rs.State <> 0 then rs.Close
	rs.open s, cn
	if Err <> 0 then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)

	dim intIndex, s_op_cadastradas, s_descricao, s_checked, s_cod_nivel_acesso_bloco_notas_pedido, s_cod_nivel_acesso_chamado_pedido, st_inativo, s_cor
	s_descricao = ""
	s_op_cadastradas = ""
	s_cod_nivel_acesso_bloco_notas_pedido = ""
    s_cod_nivel_acesso_chamado_pedido = ""
	st_inativo = ""
	
	if operacao_selecionada=OP_INCLUI then
		if Not rs.EOF then Response.Redirect("aviso.asp?id=" & ERR_PERFIL_JA_CADASTRADO)
	elseif operacao_selecionada=OP_CONSULTA then
		if rs.EOF then Response.Redirect("aviso.asp?id=" & ERR_PERFIL_NAO_CADASTRADO)
		s_descricao = Trim("" & rs("descricao"))
		s_cod_nivel_acesso_bloco_notas_pedido = Trim("" & rs("nivel_acesso_bloco_notas_pedido"))
        s_cod_nivel_acesso_chamado_pedido = Trim("" & rs("nivel_acesso_chamado"))
		st_inativo = Trim("" & rs("st_inativo"))
	'	OBTÉM A LISTA DE OPERAÇÕES JÁ CADASTRADAS P/ ESTE PERFIL
		s = "SELECT id_operacao FROM t_PERFIL_ITEM INNER JOIN t_PERFIL ON t_PERFIL_ITEM.id_perfil=t_PERFIL.id WHERE apelido='" & perfil_selecionado & "'"
		if rs.State <> 0 then rs.Close
		rs.open s, cn
		do while Not rs.Eof
			s = Cstr(rs("id_operacao"))
			if Right(s_op_cadastradas,1) <> "|" then s_op_cadastradas = s_op_cadastradas & "|"
			s_op_cadastradas = s_op_cadastradas & s
			rs.MoveNext
			loop
		if (s_op_cadastradas <> "") And (Right(s_op_cadastradas,1) <> "|") then s_op_cadastradas = s_op_cadastradas & "|"
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

<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
var keys = "";

function trataBodyKeypress() {
var i;
	keys += String.fromCharCode(window.event.keyCode);
	if (isSelectAllCheckBoxesKeywordOk(keys)) {
		for (i = 0; i < fCAD["ckb_op_central"].length; i++) {
			if (!fCAD["ckb_op_central"][i].disabled) fCAD["ckb_op_central"][i].checked = true;
		}
		for (i = 0; i < fCAD["ckb_op_loja"].length; i++) {
			if (!fCAD["ckb_op_loja"][i].disabled) fCAD["ckb_op_loja"][i].checked = true;
		}
	}
}

function RemovePerfil( f ) {
var b;
	b=window.confirm('Confirma a exclusão do perfil?');
	if (b){
		f.operacao_selecionada.value=OP_EXCLUI;
		dREMOVE.style.visibility="hidden";
		window.status = "Aguarde ...";
		f.submit();
		}
}

function AtualizaPerfil( f ) {
    var i, qtdeOpCentral, qtdeOpLoja, qtdeOpNivelAcessoBlocoNotas, qtdeOpNivelAcessoChamados;
	if (trim(f.c_descricao.value)=="") {
		alert('Preencha a descrição do perfil!!');
		f.c_descricao.focus();
		return;
		}

	qtdeOpNivelAcessoBlocoNotas = 0;
	qtdeOpNivelAcessoChamados = 0;
	qtdeOpCentral=0;
	for (i=0; i<f.ckb_op_central.length; i++) {
		if (f.ckb_op_central[i].checked) {
			qtdeOpCentral++;
			if ((f.ckb_op_central[i].value == OP_CEN_BLOCO_NOTAS_PEDIDO_LEITURA) || (f.ckb_op_central[i].value == OP_CEN_BLOCO_NOTAS_PEDIDO_CADASTRAMENTO)) qtdeOpNivelAcessoBlocoNotas++;
			if ((f.ckb_op_central[i].value == OP_CEN_PEDIDO_CHAMADO_LEITURA_QUALQUER_CHAMADO) || (f.ckb_op_central[i].value == OP_CEN_PEDIDO_CHAMADO_CADASTRAMENTO)) qtdeOpNivelAcessoChamados++;
			}
		}

	qtdeOpLoja=0;
	for (i=0; i<f.ckb_op_loja.length; i++) {
		if (f.ckb_op_loja[i].checked) {
			qtdeOpLoja++;
			if ((f.ckb_op_loja[i].value == OP_LJA_BLOCO_NOTAS_PEDIDO_LEITURA) || (f.ckb_op_loja[i].value == OP_LJA_BLOCO_NOTAS_PEDIDO_CADASTRAMENTO)) qtdeOpNivelAcessoBlocoNotas++;
			if ((f.ckb_op_loja[i].value == OP_LJA_PEDIDO_CHAMADO_LEITURA_QUALQUER_CHAMADO) || (f.ckb_op_loja[i].value == OP_LJA_PEDIDO_CHAMADO_CADASTRAMENTO)) qtdeOpNivelAcessoChamados++;
        }
		}
    // nivel de acesso ao bloco de notas do pedido
	if ((qtdeOpNivelAcessoBlocoNotas > 0) && (trim(f.c_nivel_acesso_bloco_notas.value) == "")) {
		alert('Selecione o nível de acesso para o bloco de notas do pedido!!');
		f.c_nivel_acesso_bloco_notas.focus();
		return;
		}

	if ((qtdeOpNivelAcessoBlocoNotas == 0) && (trim(f.c_nivel_acesso_bloco_notas.value) != "")) {
		alert('O nível de acesso para o bloco de notas do pedido foi definido, mas nenhuma operação referente ao bloco de notas foi habilitada!!');
		f.c_nivel_acesso_bloco_notas.focus();
		return;
	}
    // nivel de acesso aos chamados do pedido
	if ((qtdeOpNivelAcessoChamados > 0) && (trim(f.c_nivel_acesso_chamado.value) == "")) {
	    alert('Selecione o nível de acesso para os chamados do pedido!!');
	    f.c_nivel_acesso_chamado.focus();
	    return;
	}

	if ((qtdeOpNivelAcessoChamados == 0) && (trim(f.c_nivel_acesso_chamado.value) != "")) {
	    alert('O nível de acesso para os chamados do pedido foi definido, mas nenhuma operação de leitura ou cadastramento de chamados foi habilitada!!');
	    f.c_nivel_acesso_chamado.focus();
	    return;
	}
		
	if ((qtdeOpCentral==0)&&(qtdeOpLoja==0)) {
		alert('Nenhuma operação da lista foi selecionada!!');
		return;
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
#ckb_op_central {
	margin: 0pt 2pt 1pt 15pt;
	}
#ckb_op_loja {
	margin: 0pt 2pt 1pt 15pt;
	}
</style>


<%	if operacao_selecionada=OP_INCLUI then
		s = "fCAD.c_descricao.focus();"
	else
		s = "focus();"
		end if
%>
<body onload="<%=s%>" onkeypress="trataBodyKeypress();">
<center>



<!--  CADASTRO DO PERFIL -->

<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
<%	if operacao_selecionada=OP_INCLUI then
		s = "Cadastro de Novo Perfil"
	else
		s = "Consulta/Edição de Perfil Cadastrado"
		end if
%>
	<td align="center" valign="bottom"><p class="PEDIDO"><%=s%><br><span class="C">&nbsp;</span></p></td>
</tr>
</table>
<br>


<!--  CAMPOS DO CADASTRO  -->
<form id="fCAD" name="fCAD" method="post" action="PerfilAtualiza.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="operacao_selecionada" id="operacao_selecionada" value='<%=operacao_selecionada%>'>

<!-- ************   PERFIL   ************ -->
<table width="649" class="Q" cellSpacing="0">
	<tr>
		<td class="MD" width="25%"><p class="R">PERFIL</p><p class="C"><input id="perfil_selecionado" name="perfil_selecionado" class="TA" value="<%=perfil_selecionado%>" readonly size="20" style="text-align:left; color:#0000ff"></p></td>
		<td width="75%"><p class="R">DESCRIÇÃO</p><p class="C"><input id="c_descricao" name="c_descricao" class="TA" maxlength="40" size="70" value="<%=s_descricao%>" onkeypress="filtra_nome_identificador();"></p></td>
	</tr>
	<tr>
		<td colspan="2" class="MC"><p class="R">STATUS</p>
			<%	s_checked=""
				if (st_inativo = "0") Or (st_inativo = "") then s_checked = " CHECKED" %>
			<p class="C"><input type="radio" name="rb_st_inativo" value="0" style="color:green;margin-left:20px;" <%=s_checked%> /><span style="cursor:default;color:green;" onclick="fCAD.rb_st_inativo[0].click();">Ativo</span></p>
			<%	s_checked=""
				if st_inativo = "1" then s_checked = " CHECKED" %>
			<p class="C"><input type="radio" name="rb_st_inativo" value="1" style="color:red;margin-left:20px;" <%=s_checked%> /><span style="cursor:default;color:red;" onclick="fCAD.rb_st_inativo[1].click();">Inativo</span></p>
		</td>
	</tr>
	<tr>
		<td colspan="2" class="MC"><p class="R">NÍVEL DE ACESSO AO BLOCO DE NOTAS DO PEDIDO</p>
		<p class="C">
			<select id="c_nivel_acesso_bloco_notas" name="c_nivel_acesso_bloco_notas" style="margin-top:6px;margin-bottom:4px;margin-left:10px;width:180px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) {this.options[0].selected=true;}">
			<% =nivel_acesso_bloco_notas_pedido_monta_itens_select(s_cod_nivel_acesso_bloco_notas_pedido, COD_NIVEL_ACESSO_BLOCO_NOTAS_PEDIDO__ILIMITADO) %>
			</select>
		</p>
		</td>
	</tr>
    <tr>
		<td colspan="2" class="MC"><p class="R">NÍVEL DE ACESSO AOS CHAMADOS DO PEDIDO</p>
		<p class="C">
			<select id="c_nivel_acesso_chamado" name="c_nivel_acesso_chamado" style="margin-top:6px;margin-bottom:4px;margin-left:10px;width:180px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) {this.options[0].selected=true;}">
			<% =nivel_acesso_chamado_pedido_monta_itens_select(s_cod_nivel_acesso_chamado_pedido, COD_NIVEL_ACESSO_CHAMADO_PEDIDO__ILIMITADO, False) %>
			</select>
		</p>
		</td>
	</tr>
</table>

<!-- ************   OPERAÇÕES DA CENTRAL   ************ -->
<table width="649" class="QS" cellSpacing="0">
	<tr>
		<td width="100%">
		<p class="R">OPERAÇÕES DA CENTRAL</p>
<%
	s = "SELECT * FROM t_OPERACAO WHERE modulo='" & COD_OP_MODULO_CENTRAL & "' ORDER BY ordenacao"
	if rs.State <> 0 then rs.Close
	rs.open s, cn
	
	intIndex = -1
	do while Not rs.Eof
		intIndex = intIndex + 1
		s = "|" & Cstr(rs("id")) & "|"
		s_checked = ""
		if Instr(s_op_cadastradas, s) > 0 then s_checked = " checked"
		s_cor = "black"
		if Trim("" & rs("st_inativo")) = "1" then s_cor="#A9A9A9"
%>
		<p class="C"><input type="checkbox" id="ckb_op_central" name="ckb_op_central" value="<%=Cstr(rs("id"))%>" class="TA"<%=s_checked%>><span style="cursor:default;color:<%=s_cor%>;" onclick="fCAD.ckb_op_central[<%=Cstr(intIndex)%>].click();"><%=Trim("" & rs("descricao"))%></span>&nbsp;</p>
<%
		rs.MoveNext
		loop
%>
		</td>
	</tr>
</table>

<!-- ************   OPERAÇÕES DA LOJA   ************ -->
<table width="649" class="QS" cellSpacing="0">
	<tr>
		<td width="100%">
		<p class="R">OPERAÇÕES DA LOJA</p>
<%
	s = "SELECT * FROM t_OPERACAO WHERE modulo='" & COD_OP_MODULO_LOJA & "' ORDER BY ordenacao"
	if rs.State <> 0 then rs.Close
	rs.open s, cn
	
	intIndex = -1
	do while Not rs.Eof
		intIndex = intIndex + 1
		s = "|" & Cstr(rs("id")) & "|"
		s_checked = ""
		if Instr(s_op_cadastradas, s) > 0 then s_checked = " checked"
		s_cor = "black"
		if Trim("" & rs("st_inativo")) = "1" then s_cor="#A9A9A9"
%>
		<p class="C"><input type="checkbox" id="ckb_op_loja" name="ckb_op_loja" value="<%=Cstr(rs("id"))%>" class="TA"<%=s_checked%>><span style="cursor:default;color:<%=s_cor%>;" onclick="fCAD.ckb_op_loja[<%=Cstr(intIndex)%>].click();"><%=Trim("" & rs("descricao"))%></span>&nbsp;</p>
<%
		rs.MoveNext
		loop
%>
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
	<td><a href="javascript:history.back()" title="cancela as alterações do perfil">
		<img src="../botao/cancelar.gif" width="176" height="55" border="0"></a></td>
	<%
	s = ""
	if operacao_selecionada=OP_CONSULTA then
		s = "<td align='CENTER'><div name='dREMOVE' id='dREMOVE'><a href='javascript:RemovePerfil(fCAD)' "
		s =s + "title='remove o perfil cadastrado'><img src='../botao/remover.gif' width=176 height=55 border=0></a></div>"
		end if
	%><%=s%>
	<td align="right"><div name="dATUALIZA" id="dATUALIZA">
		<a name="bATUALIZA" id="bATUALIZA" href="javascript:AtualizaPerfil(fCAD)" title="atualiza o cadastro do perfil">
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
	if rs.State <> 0 then rs.Close
	set rs = nothing
	
	cn.Close
	set cn = nothing
%>