<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ===============================
'	  U S U A R I O E D I T A . A S P
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
'			I N I C I A L I Z A     P � G I N A     A S P    N O    S E R V I D O R
' _____________________________________________________________________________________________

	On Error GoTo 0
	Err.Clear
	
'	OBTEM O ID
	dim s, s_loja, usuario, usuario_selecionado, operacao_selecionada, usuario_bloqueado, vendedor_externo, s_cor
	usuario = trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 
	
	dim alerta
	alerta = ""
	
'	USU�RIO A EDITAR
	dim senha_descripto, chave
	usuario_selecionado = Ucase(trim(request("usuario_selecionado")))
	operacao_selecionada = trim(request("operacao_selecionada"))
	
	if operacao_selecionada=OP_INCLUI then
		if usuario_selecionado <> filtra_nome_identificador(usuario_selecionado) then
			alerta=texto_add_br(alerta)
			alerta=alerta & "Identificador cont�m caractere(s) inv�lido(s)!"
			end if
		end if
		
	if usuario_selecionado="" then Response.Redirect("aviso.asp?id=" & ERR_USUARIO_NAO_ESPECIFICADO) 
	if (operacao_selecionada<>OP_INCLUI) And (operacao_selecionada<>OP_CONSULTA) then Response.Redirect("aviso.asp?id=" & ERR_OPERACAO_NAO_ESPECIFICADA)
	
'	CONECTA COM O BANCO DE DADOS
	dim cn, rs, r
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	set rs = cn.Execute("select * from t_USUARIO where (usuario='" & QuotedStr(usuario_selecionado) & "')")
	if Err <> 0 then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
	
	dim intIndex, s_perfil_cadastrado, s_checked
	dim s_lista_lojas_vendedor
	dim s_cd_cadastrado
	
	s_perfil_cadastrado = ""
	s_lista_lojas_vendedor = ""
	s_cd_cadastrado = ""
	
	if operacao_selecionada=OP_INCLUI then
		if Not rs.EOF then Response.Redirect("aviso.asp?id=" & ERR_USUARIO_JA_CADASTRADO)
		set r = cn.Execute("SELECT * FROM t_ORCAMENTISTA_E_INDICADOR WHERE (apelido = '" & QuotedStr(usuario_selecionado) & "')")
		if Not r.Eof then Response.Redirect("aviso.asp?id=" & ERR_ID_JA_EM_USO_POR_ORCAMENTISTA)
	elseif operacao_selecionada=OP_CONSULTA then
		if rs.EOF then Response.Redirect("aviso.asp?id=" & ERR_USUARIO_NAO_CADASTRADO)
		
	'	OBT�M A LISTA DE OPERA��ES J� CADASTRADAS P/ ESTE PERFIL
		s = "SELECT id_perfil FROM t_PERFIL_X_USUARIO WHERE usuario='" & QuotedStr(usuario_selecionado) & "'"
		set r = cn.Execute(s)
		if Err <> 0 then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)

		do while Not r.Eof
			s = Trim("" & r("id_perfil"))
			if Right(s_perfil_cadastrado,1) <> "|" then s_perfil_cadastrado = s_perfil_cadastrado & "|"
			s_perfil_cadastrado = s_perfil_cadastrado & s
			r.MoveNext
			loop
		if (s_perfil_cadastrado <> "") And (Right(s_perfil_cadastrado,1) <> "|") then s_perfil_cadastrado = s_perfil_cadastrado & "|"

		if r.State <> 0 then r.Close
		set r = nothing
		
	'	SE � VENDEDOR DA LOJA, OBT�M A LISTA DE LOJAS LIBERADAS
		s = "SELECT loja FROM t_USUARIO_X_LOJA WHERE usuario='" & QuotedStr(usuario_selecionado) & "' ORDER BY CONVERT(smallint, loja)"
		set r = cn.Execute(s)
		if Err <> 0 then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)

		do while Not r.Eof
			s = Trim("" & r("loja"))
			if Right(s_lista_lojas_vendedor,1) <> "|" then s_lista_lojas_vendedor = s_lista_lojas_vendedor & "|"
			s_lista_lojas_vendedor = s_lista_lojas_vendedor & s
			r.MoveNext
			loop
		if (s_lista_lojas_vendedor <> "") And (Right(s_lista_lojas_vendedor,1) <> "|") then s_lista_lojas_vendedor = s_lista_lojas_vendedor & "|"
		
		if r.State <> 0 then r.Close
		set r = nothing

	'	OBT�M A LISTA DE CD'S J� CADASTRADOS P/ ESTE USU�RIO
		dim s_sql_lista_cd_cadastrado
		s_sql_lista_cd_cadastrado = ""

		s = "SELECT id_nfe_emitente FROM t_USUARIO_X_NFe_EMITENTE WHERE usuario='" & QuotedStr(usuario_selecionado) & "'"
		set r = cn.Execute(s)
		if Err <> 0 then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)

		do while Not r.Eof
			s = Trim("" & r("id_nfe_emitente"))
			if Right(s_cd_cadastrado,1) <> "|" then s_cd_cadastrado = s_cd_cadastrado & "|"
			s_cd_cadastrado = s_cd_cadastrado & s
			if s_sql_lista_cd_cadastrado <> "" then s_sql_lista_cd_cadastrado = s_sql_lista_cd_cadastrado & ","
			s_sql_lista_cd_cadastrado = s_sql_lista_cd_cadastrado & s
			r.MoveNext
			loop
		if (s_cd_cadastrado <> "") And (Right(s_cd_cadastrado,1) <> "|") then s_cd_cadastrado = s_cd_cadastrado & "|"

		if r.State <> 0 then r.Close
		set r = nothing
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

<% if False then 'APENAS P/ HABILITAR O INTELLISENSE DURANTE O DESENVOLVIMENTO!! %>
<script src="../Global/jquery.js" language="JavaScript" type="text/javascript"></script>
<% end if %>

<script src="<%=URL_FILE__JQUERY%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_I18N%>" Language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
var TAM_MIN_SENHA = <%=TAM_MIN_SENHA%>;

	$(function () {
		$(".CKBLOJAVEND").change(function () {
			if ($(this).is(":checked")) {
				$("#ckb_vendedor").prop("checked", true);
			}
		});
	});

function RemoveUsuario( f ) {
var b;
	b=window.confirm('Confirma a exclus�o do usu�rio?');
	if (b){
		f.operacao_selecionada.value=OP_EXCLUI;
		dREMOVE.style.visibility="hidden";
		window.status = "Aguarde ...";
		f.submit();
		}
}

function AtualizaUsuario( f ) {
var i, s_senha, blnTemLoja;
	s_senha=trim(f.senha.value);
	if (s_senha=="") {
		alert('Preencha a senha!!');
		f.senha.focus();
		return;
		}
		
	if (s_senha != trim(f.senha2.value)) {
		alert('A confirma��o da senha n�o est� correta!!');
		f.senha2.focus();
		return;
		}

	// Valida��es realizadas somente p/ inclus�o de novo usu�rio ou se alterou a senha
	if ((f.operacao_selecionada.value == OP_INCLUI) || (f.senha.value != f.senha_original.value)) {
		if (s_senha.length < TAM_MIN_SENHA) {
			alert('A senha deve possuir no m�nimo ' + TAM_MIN_SENHA + ' caracteres!!');
			f.senha.focus();
			return;
		}

		if (!(tem_digito(s_senha) && tem_letra(s_senha))) {
			alert("A senha deve conter no m�nimo 1 letra e 1 d�gito num�rico");
			f.senha.focus();
			return;
		}
	}

	if (trim(f.nome.value)=="") {
		alert('Preencha o nome!!');
		f.nome.focus();
		return;
	}

	if (trim(f.email.value) != "") {
	    if (!email_ok(trim(f.email.value))) {
	        alert("Endere�o de e-mail inv�lido!!");
	        f.email.focus();
	        return;
	    }
	}

	if ((trim(f.ddd.value) != "") || (trim(f.telefone.value) != "")) {
		if (trim(f.ddd.value) == "") {
			alert("Informe o DDD!")
			f.ddd.focus();
			return;
		}
		if (!ddd_ok(f.ddd.value)) {
			alert("DDD inv�lido!")
			f.ddd.focus();
			return;
		}
		if (trim(f.telefone.value) == "") {
			alert("Informe o telefone!")
			f.telefone.focus();
			return;
		}
		if (!telefone_ok(f.telefone.value)) {
			alert("Telefone inv�lido!")
			f.telefone.focus();
			return;
		}
	}

	if (f.ckb_vendedor.checked) {
		blnTemLoja=false;
		for (i=0; i<f.ckb_loja_vendedor.length; i++) {
			if (f.ckb_loja_vendedor[i].checked) {
				blnTemLoja=true;
				break;
				}
			}
		if (!blnTemLoja) {
			alert('Indique a(s) loja(s) que este vendedor pode acessar!!');
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
<link href="<%=URL_FILE__E_CSS%>" rel="stylesheet" type="text/css">
<link href="<%=URL_FILE__EPRINTER_CSS%>" rel="stylesheet" type="text/css" media="print">

<style type="text/css">
#ckb_vendedor, #ckb_vendedor_ext, #ckb_perfil, #ckb_usuario_x_cd {
	margin: 0pt 2pt 1pt 15pt;
	}

#ckb_loja_vendedor {
	margin: 0pt 2pt 1pt 45pt;
	}

#rb_estado {
	margin: 0pt 2pt 1pt 15pt;
	}
</style>



<% if alerta <> "" then %>
<!-- ************************************************************ -->
<!-- **********  P�GINA PARA EXIBIR MENSAGENS DE ERRO  ********** -->
<!-- ************************************************************ -->
<body onload="bVOLTAR.focus();">
<center>
<br>
<!--  T E L A  -->
<p class="T">A V I S O</p>
<div class="MtAlerta" style="width:600px;font-weight:bold;" align="center"><p style='margin:5px 2px 5px 2px;'><%=alerta%></p></div>
<br><br>
<p class="TracoBottom"></p>
<table cellSpacing="0">
<tr>
	<td align="center"><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()"><img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
</tr>
</table>
</center>
</body>



<% else %>

<%	if operacao_selecionada=OP_INCLUI then
		s = "fCAD.senha.focus();"
	else
		s = "focus();"
		end if
%>
<body onload="<%=s%>">
<center>



<!--  CADASTRO DO USU�RIO -->

<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
<%	if operacao_selecionada=OP_INCLUI then
		s = "Cadastro de Novo Usu�rio"
	else
		s = "Consulta/Edi��o de Usu�rio Cadastrado"
		end if
%>
	<td align="center" valign="bottom"><span class="PEDIDO"><%=s%></span><br /></td>
</tr>
</table>
<br>


<!--  CAMPOS DO CADASTRO  -->
<form id="fCAD" name="fCAD" method="post" action="usuarioAtualiza.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="operacao_selecionada" id="operacao_selecionada" value='<%=operacao_selecionada%>' />

<!-- ************   USU�RIO/SENHA   ************ -->
<table width="649" class="Q" cellspacing="0">
	<tr>
		<td class="MD" width="50%" align="left"><p class="R">USU�RIO</p><p class="C"><input id="usuario_selecionado" name="usuario_selecionado" class="TA" value="<%=usuario_selecionado%>" readonly size="30" style="text-align:left; color:#0000ff"></p></td>
<%
	senha_descripto= ""
	if operacao_selecionada=OP_CONSULTA then
		s = Trim("" & rs("datastamp"))
		chave = gera_chave(FATOR_BD)
		decodifica_dado s, senha_descripto, chave
		end if
%>
		<td class="MD" width="25%" align="left"><p class="R">SENHA</p><p class="C"><input id="senha" name="senha" class="TA" type="password" maxlength="15" size="18" value="<%=senha_descripto%>" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCAD.senha2.focus();"></p></td>
		<td width="25%" align="left"><p class="R">SENHA (CONFIRMA��O)</p><p class="C"><input id="senha2" name="senha2" class="TA" type="password" maxlength="15" size="18" value="<%=senha_descripto%>" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fCAD.nome.focus();"></p></td>
		<input type="hidden" name="senha_original" id="senha_original" value="<%=senha_descripto%>" />
	</tr>
</table>

<!-- ************   NOME   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
<%
	s=""
	if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("nome"))
%>
		<td width="100%" align="left"><p class="R">NOME</p><p class="C"><input id="nome" name="nome" class="TA" value="<%=s%>" maxlength="40" size="85" onkeypress="filtra_nome_identificador();"></p></td>
	</tr>
</table>

<!-- ************   E-MAIL   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
<%
	s=""
	if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("email"))
%>
		<td width="100%" align="left"><p class="R">E-MAIL</p><p class="C"><input id="email" name="email" class="TA" value="<%=s%>" maxlength="60" size="85" onkeypress="filtra_email();"></p></td>
	</tr>
</table>

<!-- ************   TELEFONE   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
<%
	s=""
	if operacao_selecionada=OP_CONSULTA then s=Trim("" & rs("ddd"))
%>
		<td width="25%" align="left" class="MD"><p class="R">DDD</p><p class="C"><input id="ddd" name="ddd" class="TA" value="<%=s%>" maxlength="2" size="4" onkeypress="filtra_numerico();" onblur="this.value=retorna_so_digitos(this.value);"></p></td>
<%
	s=""
	if operacao_selecionada=OP_CONSULTA then s=telefone_formata(Trim("" & rs("telefone")))
%>
		<td width="75%" align="left"><p class="R">TELEFONE</p><p class="C"><input id="telefone" name="telefone" class="TA" value="<%=s%>" maxlength="10" size="12" onkeypress="filtra_numerico();" onblur="this.value=telefone_formata(retorna_so_digitos(this.value));"></p></td>
	</tr>
</table>

<!-- ************   ESTADO BLOQUEADO?   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
<%
	usuario_bloqueado=false
	if operacao_selecionada=OP_CONSULTA then
		if rs("bloqueado") <> 0 then usuario_bloqueado=true
		end if
%>
		<td width="100%" align="left">
		<p class="R">ESTADO</p>
		<p class="C"><input type="radio" id="rb_estado" name="rb_estado" value="0" class="TA"<%if not usuario_bloqueado then Response.Write(" checked")%>><span onclick="fCAD.rb_estado[0].click();" style="cursor:default; color:#006600">Acesso permitido</span>&nbsp;</p>
		<p class="C"><input type="radio" id="rb_estado" name="rb_estado" value="1" class="TA"<%if usuario_bloqueado then Response.Write(" checked")%>><span onclick="fCAD.rb_estado[1].click();" style="cursor:default; color:#ff0000">Acesso bloqueado</span>&nbsp;</p>
		</td>
	</tr>
</table>

<!-- ************   LOGIN BLOQUEADO AUTOMATICAMENTE?   ************ -->
<%
if operacao_selecionada=OP_CONSULTA then
%>
<table width="649" class="QS" cellspacing="0">
	<tr>
<%
	s = "&nbsp;"
	s_cor = "black"
	if rs("StLoginBloqueadoAutomatico") <> 0 then
		s = "Bloqueado em " & formata_data_hora_sem_seg(rs("DataHoraBloqueadoAutomatico")) & " (" & Trim("" & rs("QtdeConsecutivaFalhaLogin")) & " tentativas consecutivas com senha errada)"
		s_cor = "red"
		end if
%>
		<td width="100%" align="left">
		<p class="R">LOGIN BLOQUEADO AUTOMATICAMENTE</p>
		<p class="C" id="pMsgStLoginBloqueadoAutomatico" style="color:<%=s_cor%>;"><%=s%>
		<% if rs("StLoginBloqueadoAutomatico") <> 0 then %>
		<input type="checkbox" id="ckb_desbloquear_bloqueio_automatico" name="ckb_desbloquear_bloqueio_automatico" value="ON" class="TA" style="margin-left:15px;" /><span class="C" onclick="fCAD.ckb_desbloquear_bloqueio_automatico.click();" style="cursor:default;">Desbloquear</span>
		<% end if %>
		</p>
		</td>
	</tr>
</table>
<%
	end if
%>

<!-- ************   PERFIS DE ACESSO   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
		<td width="100%" align="left">
		<p class="R">PERFIL DE ACESSO</p>
<%
	s = "SELECT * FROM t_PERFIL ORDER BY apelido"
	set r = cn.Execute(s)
	
	intIndex = -1
	do while Not r.Eof
		if Trim("" & r("st_oculto")) = "0" then
			intIndex = intIndex + 1
			s = "|" & Cstr(r("id")) & "|"
			s_checked = ""
			if Instr(s_perfil_cadastrado, s) > 0 then s_checked = " checked"
			s_cor = "black"
			if Trim("" & r("st_inativo")) = "1" then s_cor = "#A9A9A9"
	%>
			<p class="C"><input type="checkbox" id="ckb_perfil" name="ckb_perfil" value="<%=Cstr(r("id"))%>" class="TA"<%=s_checked%>><span style="cursor:default;color:<%=s_cor%>;" onclick="fCAD.ckb_perfil[<%=Cstr(intIndex)%>].click();"><%=Trim("" & r("apelido")) & " - " & Trim("" & r("descricao"))%></span>&nbsp;</p>
	<%
		else 'if-then-else Trim("" & rs("st_oculto")) = "0"
			'Se o perfil est� c/ o flag oculto ativo, n�o exibe na tela, mas cria um campo hidden p/ mant�-lo no cadastro do usu�rio
			s = "|" & Cstr(r("id")) & "|"
			if Instr(s_perfil_cadastrado, s) > 0 then
				intIndex = intIndex + 1
	%>
				<input type="hidden" id="ckb_perfil" name="ckb_perfil" value="<%=Cstr(r("id"))%>" />
	<%
				end if
			end if 'if Trim("" & rs("st_oculto")) = "0"

		r.MoveNext
		loop

	if r.State <> 0 then r.Close
	set r = nothing
%>
		</td>
	</tr>
</table>

<!-- ************   VENDEDOR EXTERNO?   ************ -->
<table width="649" class="QS" cellspacing="0">
	<tr>
<%
	vendedor_externo=false
	if operacao_selecionada=OP_CONSULTA then
		if rs("vendedor_externo") <> 0 then vendedor_externo=true
		end if
%>
		<td width="100%" align="left">
		<p class="R">VENDEDOR EXTERNO</p>
		<p class="C"><input type="checkbox" id="ckb_vendedor_ext" name="ckb_vendedor_ext" value="CKB_VENDEDOR_EXT_ON" class="TA"<%if vendedor_externo then Response.Write(" checked")%>><span onclick="fCAD.ckb_vendedor_ext.click();" style="cursor:default;">Vendedor Externo</span>&nbsp;</p>
		</td>
	</tr>
</table>

<!-- ************   VENDEDOR DA LOJA?   ************ -->
<!-- FOR�A A CRIA��O DE UM ARRAY MESMO QUANDO EXISTE SOMENTE 1 ITEM -->
<input type="hidden" class="CBOX" name="ckb_loja_vendedor" id="ckb_loja_vendedor" value="" />

<table width="649" class="QS" cellspacing="0">
	<tr>
		<td width="100%" align="left">
		<p class="R">VENDEDOR DA LOJA</p>
<%  s=""
	s_loja=""
	if operacao_selecionada=OP_CONSULTA then 
		s=Cstr(rs("vendedor_loja"))
		s_loja=Trim("" & rs("loja"))
		end if
%>
		<p class="C"><input type="checkbox" id="ckb_vendedor" name="ckb_vendedor" value="<%=ID_VENDEDOR%>" class="TA" <%if s="1" then Response.Write(" checked")%>><span style="cursor:default" onclick="fCAD.ckb_vendedor.click();">Vendedor da loja</span></p>
		
<%
	s = "SELECT loja, nome FROM t_LOJA ORDER BY CONVERT(smallint, loja)"
	set r = cn.Execute(s)
	
	intIndex = -1
	do while Not r.Eof
		intIndex = intIndex + 1
		s = "|" & Cstr(r("loja")) & "|"
		s_checked = ""
		if Instr(s_lista_lojas_vendedor, s) > 0 then s_checked = " checked"
%>
		<p class="C"><input type="checkbox" id="ckb_loja_vendedor" name="ckb_loja_vendedor" value="<%=Cstr(r("loja"))%>" class="TA CKBLOJAVEND"<%=s_checked%>><span style="cursor:default" onclick="fCAD.ckb_loja_vendedor[<%=Cstr(intIndex+1)%>].click();"><%=Trim("" & r("loja"))%> - <%=Trim("" & r("nome"))%></span>&nbsp;</p>
<%
		r.MoveNext
		loop
		
	if r.State <> 0 then r.Close
	set r = nothing
%>
		</td>
	</tr>
</table>

<!-- ************   ACESSA OPERA��ES QUE ENVOLVEM CENTRO DE DISTRIBUI��O?   ************ -->
<!-- FOR�A A CRIA��O DE UM ARRAY MESMO QUANDO EXISTE SOMENTE 1 ITEM -->
<input type="hidden" class="CBOX" name="ckb_usuario_x_cd" id="ckb_usuario_x_cd" value="" />

<table width="649" class="QS" cellspacing="0">
	<tr>
		<td width="100%" align="left">
		<p class="R">CD</p>
<%
	s = "SELECT" & _
			" id," & _
			" apelido" & _
		" FROM t_NFe_EMITENTE" & _
		" WHERE" & _
			"((st_ativo=1) AND (st_habilitado_ctrl_estoque=1))"
	if s_sql_lista_cd_cadastrado <> "" then
		s = s & _
			" OR " &_ 
			"(id IN (" & s_sql_lista_cd_cadastrado & "))"
		end if
	s =s & _
		" ORDER BY" & _
			" ordem"
	set r = cn.Execute(s)
	
	intIndex = -1
	do while Not r.Eof
		intIndex = intIndex + 1
		s = "|" & Cstr(r("id")) & "|"
		s_checked = ""
		if Instr(s_cd_cadastrado, s) > 0 then s_checked = " checked"
%>
		<p class="C"><input type="checkbox" id="ckb_usuario_x_cd" name="ckb_usuario_x_cd" value="<%=Cstr(r("id"))%>" class="TA"<%=s_checked%>><span style="cursor:default" onclick="fCAD.ckb_usuario_x_cd[<%=Cstr(intIndex+1)%>].click();"><%=Ucase(Trim("" & r("apelido")))%></span>&nbsp;</p>
<%
		r.MoveNext
		loop
		
	if r.State <> 0 then r.Close
	set r = nothing
%>
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
	<td align="left"><a href="javascript:history.back()" title="cancela as altera��es do usu�rio">
		<img src="../botao/cancelar.gif" width="176" height="55" border="0"></a></td>
	<%
	s = ""
	if operacao_selecionada=OP_CONSULTA then
		s = "<td align='center'><div name='dREMOVE' id='dREMOVE'><a href='javascript:RemoveUsuario(fCAD)' "
		s =s + "title='remove o usu�rio cadastrado'><img src='../botao/remover.gif' width=176 height=55 border=0></a></div>"
		end if
	%><%=s%>
	<td align="right"><div name="dATUALIZA" id="dATUALIZA">
		<a name="bATUALIZA" id="bATUALIZA" href="javascript:AtualizaUsuario(fCAD)" title="atualiza o cadastro do usu�rio">
		<img src="../botao/confirmar.gif" width="176" height="55" border="0"></a></div>
	</td>
</tr>
</table>
</form>

</center>
</body>

<% end if %>

</html>


<%
'	FECHA CONEXAO COM O BANCO DE DADOS
	rs.Close
	set rs = nothing
	
	cn.Close
	set cn = nothing
%>