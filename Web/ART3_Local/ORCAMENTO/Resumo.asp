<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp"        -->
<!-- #include file = "../global/Global.asp"        -->
<%

'     ===================
'	  R E S U M O . A S P
'     ===================
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
	
	Server.ScriptTimeout = MAX_SERVER_SCRIPT_TIMEOUT_EM_SEG
	
'	EXIBIÇÃO DE BOTÕES DE PESQUISA DE CEP
	dim blnPesquisaCEPAntiga, blnPesquisaCEPNova
	
	blnPesquisaCEPAntiga = False
	blnPesquisaCEPNova = True
	
	Const blnQuadroAvisosHabilitado = False
	
'	VERIFICA ID
	dim s, loja, loja_nome, usuario, usuario_nome, senha, senha_real, chave
	dim dt_ult_alteracao_senha, usuario_bloqueado, usuario_bloqueado_automatico, confere_login_no_bd, eh_primeira_execucao
	dim idUsuario, qtdeConsecutivaFalhaLogin, max_tentativas_login, blnUsuarioCadastrado, blnSenhaConfereOk, dtHrBloqueioAutomatico
	dim id_email, assunto_mensagem, corpo_mensagem, remetente_mensagem, msg_erro_grava_email, rEmailDestinatario, ambiente_execucao
	
	confere_login_no_bd = (Trim(Session("usuario_a_checar")) <> "")	
	usuario = Trim(Session("usuario_a_checar")): Session("usuario_a_checar") = " "
	senha = Trim(Session("senha_a_checar")): Session("senha_a_checar") = " "
	
	if usuario = "" then usuario = Session("usuario_atual")
	if senha = "" then senha = Session("senha_atual")
	usuario_nome = Session("usuario_nome_atual")
	loja_nome = Session("loja_nome_atual")

	if (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO)

	if isHorarioManutencaoSistema then Response.Redirect("aviso.asp?id=" & ERR_HORARIO_MANUTENCAO_SISTEMA)

	dim strSessionCtrlTicket

'	CONECTA COM O BANCO DE DADOS
	dim cn,rs
	If Not bdd_conecta(cn) then Response.Redirect("Aviso.asp?id=" & ERR_CONEXAO)

'	VERIFICA LOJA NO BD
	if confere_login_no_bd then
		eh_primeira_execucao = true

	'	VERIFICA USUARIO E SENHA NO BD
		dt_ult_alteracao_senha = null
		usuario_bloqueado=false
		usuario_bloqueado_automatico=False
		blnUsuarioCadastrado=False
		blnSenhaConfereOk=False

		s = "SELECT" & _
				" Id" & _
				", loja" & _
				", razao_social_nome" & _
				", senha" & _
				", datastamp" & _
				", dt_ult_alteracao_senha" & _
				", hab_acesso_sistema" & _
				", status" & _
				", StLoginBloqueadoAutomatico" & _
				", QtdeConsecutivaFalhaLogin" & _
			" FROM t_ORCAMENTISTA_E_INDICADOR" & _
			" WHERE" & _
				" (apelido = '" & QuotedStr(usuario) & "')"
		set rs = cn.Execute(s)
		if Err <> 0 then Response.Redirect("Aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
		
		if Not rs.eof then
			blnUsuarioCadastrado=True

		'	TEM SENHA CADASTRADA?
			if Trim("" & rs("datastamp")) = "" then usuario_bloqueado=true
		'	TEM ACESSO AO SISTEMA?
			if rs("hab_acesso_sistema")<>1 then usuario_bloqueado=true
		'	ATIVO?
			if rs("status") <> "A" then usuario_bloqueado=true
			
		'	ACESSO BLOQUEADO AUTOMATICAMENTE POR EXCESSO DE TENTATIVAS C/ SENHA ERRADA?
			if rs("StLoginBloqueadoAutomatico")<>0 then usuario_bloqueado_automatico=true
			qtdeConsecutivaFalhaLogin = rs("QtdeConsecutivaFalhaLogin")
			max_tentativas_login = obtem_parametro_max_tentativas_login

			idUsuario = rs("Id")
			dt_ult_alteracao_senha = rs("dt_ult_alteracao_senha")
			usuario_nome = Trim("" & rs("razao_social_nome"))

			loja = Trim("" & rs("loja"))
			if loja = "" then Response.Redirect("Aviso.asp?id=" & ERR_IDENTIFICACAO_LOJA)
			if converte_numero(loja) = 0 then Response.Redirect("Aviso.asp?id=" & ERR_IDENTIFICACAO_LOJA)
			
			loja_nome = trim(x_loja(loja))
			If loja_nome = "" then Response.Redirect("Aviso.asp?id=" & ERR_IDENTIFICACAO_LOJA)
			
			senha_real = ""
			s = Trim("" & rs("datastamp"))
			chave = gera_chave(FATOR_BD)
			decodifica_dado s, senha_real, chave
			if UCase(trim(senha_real)) = UCase(trim(senha)) then
				'SENHA CONFERE OK
				blnSenhaConfereOk = True
			else
				if senha_real <> "" then senha = ""
				end if
			end if

		rs.close
		set rs = nothing
		
		'REGISTRA HISTÓRICO DE LOGIN (NA SEQUÊNCIA DE PRIORIDADE)
		'MOTIVO: USUÁRIO NÃO CADASTRADO
		if Not blnUsuarioCadastrado then
			'USUÁRIO NÃO CADASTRADO
			s = "INSERT INTO t_LOGIN_HISTORICO (" & _
					"StSucesso" & _
					", IP" & _
					", sistema_responsavel" & _
					", Login" & _
					", Motivo" & _
					", IdCfgModulo" & _
				") VALUES (" & _
					"0" & _
					", '" & QuotedStr(Trim("" & Request.ServerVariables("REMOTE_ADDR"))) & "'" & _
					", " & CStr(COD_SISTEMA_RESPONSAVEL_CADASTRO__ERP) & _
					", '" & QuotedStr(usuario) & "'" & _
					", '" & COD_CONTROLE_LOGIN_FALHA__USUARIO_NAO_CADASTRADO & "'" & _
					", " & CStr(ID_MODULO__PRE_PEDIDO) & _
				")"
			cn.Execute(s)

		'MOTIVO: SENHA NÃO CONFERE
		elseif Not blnSenhaConfereOk then
			qtdeConsecutivaFalhaLogin = qtdeConsecutivaFalhaLogin + 1

			'SENHA NÃO CONFERE
			s = "INSERT INTO t_LOGIN_HISTORICO (" & _
					"IdTipoUsuarioContexto" & _
					", IdUsuario" & _
					", StSucesso" & _
					", IP" & _
					", sistema_responsavel" & _
					", Login" & _
					", Motivo" & _
					", IdCfgModulo" & _
				") VALUES (" & _
					COD_USUARIO_CONTEXTO__PARCEIRO & _
					", " & CStr(idUsuario) & _
					", 0" & _
					", '" & QuotedStr(Trim("" & Request.ServerVariables("REMOTE_ADDR"))) & "'" & _
					", " & CStr(COD_SISTEMA_RESPONSAVEL_CADASTRO__ERP) & _
					", '" & QuotedStr(usuario) & "'" & _
					", '" & COD_CONTROLE_LOGIN_FALHA__SENHA_INVALIDA & "'" & _
					", " & CStr(ID_MODULO__PRE_PEDIDO) & _
				")"
			cn.Execute(s)

			'Incrementa quantidade de tentativas consecutivas com falha
			s = "UPDATE t_ORCAMENTISTA_E_INDICADOR SET" & _
					" QtdeConsecutivaFalhaLogin = QtdeConsecutivaFalhaLogin + 1" & _
				" WHERE" & _
					" (apelido = '" & QuotedStr(usuario) & "')"
			cn.Execute(s)

			if (Not usuario_bloqueado_automatico) And (qtdeConsecutivaFalhaLogin >= max_tentativas_login) then
				'Usuário será bloqueado automaticamente no próximo login
				dtHrBloqueioAutomatico = Now
				s = "UPDATE t_ORCAMENTISTA_E_INDICADOR SET" & _
						" StLoginBloqueadoAutomatico = 1" & _
						", DataHoraBloqueadoAutomatico = getdate()" & _
						", EnderecoIpBloqueadoAutomatico = '" & QuotedStr(Trim("" & Request.ServerVariables("REMOTE_ADDR"))) & "'" & _
					" WHERE" & _
						" (Id = " & CStr(idUsuario) & ")"
				cn.Execute(s)

				'Envia e-mail de alerta sobre o bloqueio automático
				set rEmailDestinatario = get_registro_t_parametro(ID_PARAMETRO_EmailDestinatarioAlertaLoginBloqueadoAutomatico)
				if Trim("" & rEmailDestinatario.campo_texto) <> "" then
					ambiente_execucao = getParametroFromCampoTexto(ID_PARAMETRO_AMBIENTE_EXECUCAO_OWNER) & "/" & getParametroFromCampoTexto(ID_PARAMETRO_AMBIENTE_EXECUCAO_CONTEXTO)
					assunto_mensagem = getParametroFromCampoTexto(ID_PARAMETRO_SubjectEmailAlertaLoginBloqueadoAutomatico)
					corpo_mensagem = getParametroFromCampoTexto(ID_PARAMETRO_BodyEmailAlertaLoginBloqueadoAutomatico)
					remetente_mensagem = getParametroFromCampoTexto(ID_PARAMETRO_EmailRemetenteAlertaLoginBloqueadoAutomatico)
					
					assunto_mensagem = Replace(assunto_mensagem, "[AMBIENTE]", ambiente_execucao)
					assunto_mensagem = Replace(assunto_mensagem, "[LOGIN_USUARIO]", usuario)
					assunto_mensagem = Replace(assunto_mensagem, "[DATA_HORA_BLOQUEIO]", formata_data_hora_sem_seg(dtHrBloqueioAutomatico))

					corpo_mensagem = Replace(corpo_mensagem, "[AMBIENTE]", ambiente_execucao)
					corpo_mensagem = Replace(corpo_mensagem, "[LOGIN_USUARIO]", usuario)
					corpo_mensagem = Replace(corpo_mensagem, "[IdTipoUsuarioContexto]", CStr(COD_USUARIO_CONTEXTO__PARCEIRO))
					corpo_mensagem = Replace(corpo_mensagem, "[IdUsuario]", CStr(idUsuario))
					corpo_mensagem = Replace(corpo_mensagem, "[IP]", QuotedStr(Trim("" & Request.ServerVariables("REMOTE_ADDR"))))
					corpo_mensagem = Replace(corpo_mensagem, "[DATA_HORA_BLOQUEIO]", formata_data_hora_sem_seg(dtHrBloqueioAutomatico))
					corpo_mensagem = Replace(corpo_mensagem, "[MAX_TENTATIVAS_LOGIN]", CStr(max_tentativas_login))

					EmailSndSvcGravaMensagemParaEnvio remetente_mensagem, _
													"", _
													rEmailDestinatario.campo_texto, _
													"", _
													"", _
													assunto_mensagem, _
													corpo_mensagem, _
													Now, _
													id_email, _
													msg_erro_grava_email
					end if 'if Trim("" & rEmailDestinatario.campo_texto) <> ""
				end if 'if (Not usuario_bloqueado_automatico) And (qtdeConsecutivaFalhaLogin >= max_tentativas_login)

		'MOTIVO: USUÁRIO ESTÁ BLOQUEADO AUTOMATICAMENTE
		elseif usuario_bloqueado_automatico then
			'USUÁRIO ENCONTRA-SE BLOQUEADO AUTOMATICAMENTE
			s = "INSERT INTO t_LOGIN_HISTORICO (" & _
					"IdTipoUsuarioContexto" & _
					", IdUsuario" & _
					", StSucesso" & _
					", IP" & _
					", sistema_responsavel" & _
					", Login" & _
					", Motivo" & _
					", IdCfgModulo" & _
				") VALUES (" & _
					COD_USUARIO_CONTEXTO__PARCEIRO & _
					", " & CStr(idUsuario) & _
					", 0" & _
					", '" & QuotedStr(Trim("" & Request.ServerVariables("REMOTE_ADDR"))) & "'" & _
					", " & CStr(COD_SISTEMA_RESPONSAVEL_CADASTRO__ERP) & _
					", '" & QuotedStr(usuario) & "'" & _
					", '" & COD_CONTROLE_LOGIN_FALHA__BLOQUEADO_AUTOMATICO & "'" & _
					", " & CStr(ID_MODULO__PRE_PEDIDO) & _
				")"
			cn.Execute(s)

		'MOTIVO: USUÁRIO ESTÁ BLOQUEADO MANUALMENTE
		elseif usuario_bloqueado then
			'USUÁRIO BLOQUEADO MANUALMENTE
			s = "INSERT INTO t_LOGIN_HISTORICO (" & _
					"IdTipoUsuarioContexto" & _
					", IdUsuario" & _
					", StSucesso" & _
					", IP" & _
					", sistema_responsavel" & _
					", Login" & _
					", Motivo" & _
					", IdCfgModulo" & _
				") VALUES (" & _
					COD_USUARIO_CONTEXTO__PARCEIRO & _
					", " & CStr(idUsuario) & _
					", 0" & _
					", '" & QuotedStr(Trim("" & Request.ServerVariables("REMOTE_ADDR"))) & "'" & _
					", " & CStr(COD_SISTEMA_RESPONSAVEL_CADASTRO__ERP) & _
					", '" & QuotedStr(usuario) & "'" & _
					", '" & COD_CONTROLE_LOGIN_FALHA__BLOQUEADO_MANUAL & "'" & _
					", " & CStr(ID_MODULO__PRE_PEDIDO) & _
				")"
			cn.Execute(s)
			end if 'if - elseif (falhas de login)


		if Not blnUsuarioCadastrado then
			cn.Close
			Response.Redirect("aviso.asp?id=" & ERR_IDENTIFICACAO)
			end if

		if Not blnSenhaConfereOk then
			cn.Close
			Response.Redirect("aviso.asp?id=" & ERR_IDENTIFICACAO)
			end if

		if usuario_bloqueado_automatico then
			cn.Close
			Response.Redirect("aviso.asp?id=" & ERR_USUARIO_BLOQUEADO_AUTOMATICO)
			end if
		
		if usuario_bloqueado then
			cn.Close
			Response.Redirect("aviso.asp?id=" & ERR_USUARIO_BLOQUEADO)
			end if
		
		
		Session("loja_atual") = loja
		Session("usuario_atual") = usuario
		Session("senha_atual") = senha
		Session("usuario_nome_atual") = usuario_nome
		Session("loja_nome_atual") = loja_nome
		
		strSessionCtrlTicket = ""

		s = "UPDATE t_ORCAMENTISTA_E_INDICADOR SET" & _
				" dt_ult_acesso = " & bd_formata_data_hora(Now) & _
				", QtdeConsecutivaFalhaLogin = 0" & _
			" WHERE" & _
				" (apelido = '" & QuotedStr(usuario) & "')"
		cn.Execute(s)

		s = "INSERT INTO t_SESSAO_HISTORICO (" & _
				"Usuario, " & _
				"SessionCtrlTicket, " & _
				"DtHrInicio, " & _
				"Loja, " & _
				"Modulo, " & _
				"IP, " & _
				"UserAgent" & _
			") VALUES (" & _
				"'" & QuotedStr(usuario) & "'," & _
				"'" & strSessionCtrlTicket & "'," & _
				bd_formata_data_hora(Session("DataHoraLogon")) & "," & _
				"'" & loja & "'," & _
				"'" & SESSION_CTRL_MODULO_ORCAMENTO & "'," & _
				"'" & QuotedStr(Trim("" & Request.ServerVariables("REMOTE_ADDR"))) & "'," & _
				"'" & QuotedStr(Trim("" & Request.ServerVariables("HTTP_USER_AGENT"))) & "'" & _
			")"
		cn.Execute(s)

		'Login bem sucedido
		s = "INSERT INTO t_LOGIN_HISTORICO (" & _
				"IdTipoUsuarioContexto" & _
				", IdUsuario" & _
				", StSucesso" & _
				", IP" & _
				", sistema_responsavel" & _
				", Login" & _
				", IdCfgModulo" & _
			") VALUES (" & _
				COD_USUARIO_CONTEXTO__PARCEIRO & _
				", " & CStr(idUsuario) & _
				", 1" & _
				", '" & QuotedStr(Trim("" & Request.ServerVariables("REMOTE_ADDR"))) & "'" & _
				", " & CStr(COD_SISTEMA_RESPONSAVEL_CADASTRO__ERP) & _
				", '" & QuotedStr(usuario) & "'" & _
				", " & CStr(ID_MODULO__PRE_PEDIDO) & _
			")"
		cn.Execute(s)

		if IsNull(dt_ult_alteracao_senha) then Response.Redirect("Senha.asp")
		end if  'if (confere_login_no_bd)
	
	Dim vMsg()
	if blnQuadroAvisosHabilitado then
		if Trim(Session("verificar_quadro_avisos")) <> "" then
			Session("verificar_quadro_avisos") = " "
			if recupera_avisos_nao_lidos(loja, usuario, vMsg) then Response.Redirect("QuadroAvisoMostra.asp")
			end if
		end if

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


<%=DOCTYPE_LEGADO%>

<html>


<head>
	<title><%=TITULO_JANELA_MODULO_ORCAMENTO%></title>
	</head>

<script language="JavaScript" type="text/javascript">
window.focus();
</script>

<script src="<%=URL_FILE__JQUERY%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_I18N%>" Language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__AJAX_JS%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JANELACEP_JS%>" language="JavaScript" type="text/javascript"></script>

<% if eh_primeira_execucao then %>
<script language="JavaScript" type="text/javascript">
configura_painel();
</script>
<% end if %>

<script language="JavaScript" type="text/javascript">
var fCepPopup;

function ProcessaSelecaoCEP(){};

function AbrePesquisaCep(){
var strUrl;
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
	ProcessaSelecaoCEP=null;
	strUrl="../Global/AjaxCepPesqPopup.asp?ModoApenasConsulta=S";
	fCepPopup=window.open(strUrl, "AjaxCepPesqPopup", "status=1,toolbar=0,location=0,menubar=0,directories=0,resizable=1,scrollbars=1,width=980,height=650,left=0,top=0");
	fCepPopup.focus();
}

function fCPConcluir( f ){
var s;
	s=f.cnpj_cpf_selecionado.value;
	s=retorna_so_digitos(s);
	if (s.length == 0) {
		alert("Informe o CNPJ/CPF do cliente!!");
		f.cnpj_cpf_selecionado.focus();
		return false;
		}
		
	if (!cnpj_cpf_ok(s)) {
		alert("CNPJ/CPF inválido!!");
		f.cnpj_cpf_selecionado.focus();
		return false;
		}

	window.status = "Aguarde ...";
	f.submit();
}

function fOFConcluir( f ){
var s, iop;

	iop=-1;
	s="";

 // LEITURA DO QUADRO DE AVISOS (SOMENTE NÃO LIDOS)
	iop++;
	if (f.rb_op[iop].checked) {
		s="QuadroAvisoMostra.asp";
		f.opcao_selecionada.value="";
		}

 // LEITURA DO QUADRO DE AVISOS (TODOS OS AVISOS)
	iop++;
	if (f.rb_op[iop].checked) {
		s="QuadroAvisoMostra.asp";
		f.opcao_selecionada.value="S";
		}

	if (s=="") {
		alert("Escolha uma das funções!!");
		return false;
		}

	window.status = "Aguarde ...";
	f.action=s;
	f.submit();
}
</script>

<script type="text/javascript">
	function exibeJanelaCEP_Consulta() {
		$.mostraJanelaCEP(null);
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



<body id="corpoPagina" link="navy" alink="navy" vlink="navy">

<!--  I D E N T I F I C A Ç Ã O  -->
<table width="100%" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">

<tr>
	<td align="right" valign="bottom">
	<%	s = usuario_nome
		if s = "" then s = usuario
		s = x_saudacao & ", " & s
		s = "<span class='Cd'>" & s & "</span>"
	%>
	<%=s%>
	</td>
</tr>
<tr>
	<td align="right" valign="bottom">
	<span class="Rc">
	<% if blnPesquisaCEPAntiga then %>
		<span name="bPesqCep" id="bPesqCep" class="LPesqCep" onclick="AbrePesquisaCep();">Pesquisar CEP</span>&nbsp;&nbsp;&nbsp;
	<% end if %>
	<% if blnPesquisaCEPAntiga and blnPesquisaCEPNova then Response.Write "&nbsp;nbsp;nbsp;" %>
	<% if blnPesquisaCEPNova then %>
		<span name="bPesqCep" id="bPesqCep" class="LPesqCep" onclick="exibeJanelaCEP_Consulta();">Pesquisar CEP</span>&nbsp;&nbsp;&nbsp;
	<% end if %>
		<a href="senha.asp" title="altera a senha atual do usuário" class="LAlteraSenha">altera senha</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span>
	</td>
</tr>

</table>
<center>

<!-- #include file = "../global/JanelaBuscaCEP.htm"    -->

<br>


<!--  CADASTRA NOVO  -->
<form action="ClientePesquisa.asp" method="post" id="fCP" name="fCP" onsubmit="if (!fCPConcluir(fCP)) return false;">
<span id="sNOVOPED" class="T">PRÉ-PEDIDO</span>
<div class="QFn" align="center" style="width:600px">
	<p class="C" style="margin: 2 10 2 10">&nbsp;</p>
	<table cellpadding="0" cellspacing="0">
		<tr>
			<td nowrap class="R" align="right">
				<p class="C" style="margin-top:5px;">CNPJ/CPF DO CLIENTE&nbsp;</p>
			</td>
			<td align="left">
				<input name="cnpj_cpf_selecionado" id="cnpj_cpf_selecionado" type="text" maxlength="18" size="20" onblur="if (!cnpj_cpf_ok(this.value)) {alert('CNPJ/CPF inválido!!');this.focus();} else this.value=cnpj_cpf_formata(this.value);" onkeypress="if (digitou_enter(true) && tem_info(this.value) && cnpj_cpf_ok(this.value)) {this.value=cnpj_cpf_formata(this.value); fCPConcluir(fCP);} filtra_cnpj_cpf();">
			</td>
		</tr>
	</table>
	<p class="R" style="margin: 4 10 0 10">&nbsp;</p>
	<table>
		<tr>
			<td align="center">
				<input name="bEXECUTAR" id="bEXECUTAR" type="submit" class="Botao" 
					   value="EXECUTAR CONSULTA" title="executa a pesquisa no cadastro de clientes">
			</td>
		</tr>
	</table>
	<p class="R" style="margin: 0 10 0 10">&nbsp;</p>
</div>
</form>


<!--  C O N S U L T A   O R Ç A M E N T O S  -->
<br />
<form action="orcamento.asp" method="post" id="fORC" name="fORC">
<span class="T">CONSULTA PRÉ-PEDIDOS</span>
<table width="600" class="Q" cellspacing="0">
  <tr class="DefaultBkg">
	<td align="center" class='MB'><p class="R"><a href='OrcamentosCadastrados.asp' onclick="javascript:window.status='Aguarde ...';">Pré-Pedidos Cadastrados</a></p></td>
  </tr>
  <tr class="DefaultBkg">
	<td align="center"><p class="R"><a href='OrcamentosQueViraramPedidos.asp' onclick="javascript:window.status='Aguarde ...';">Pré-Pedidos Que Viraram Pedidos</a></p></td>
  </tr>
</table>

<table width="600" class="QS" cellspacing="0">
  <tr class="DefaultBkg">
	<td valign="middle" align="center">
		<p class="C" style="margin: 12px 0px 12px 0px;">Nº PRÉ-PEDIDO&nbsp;&nbsp;<input size="10" maxlength="10" name="orcamento_selecionado" id="orcamento_selecionado" onkeypress="if (digitou_enter(true) && tem_info(this.value)) {if (normaliza_num_orcamento(this.value)!='') this.value=normaliza_num_orcamento(this.value); fORC.submit();} filtra_orcamento();" onblur="if (normaliza_num_orcamento(this.value)!='') this.value=normaliza_num_orcamento(this.value);">
				&nbsp;&nbsp;&nbsp;<input name="CONSULTAR" id="CONSULTAR" type="submit" class="Botao" 
										 value="CONSULTAR" title="consulta um pré-pedido específico desta loja">
		</p>
	</td>
  </tr>
</table>
</form>

<!--  C O N S U L T A   P E D I D O S  -->
<br />
<form action="pedido.asp" method="post" id="fPED" name="fPED">
<span class="T">CONSULTA PEDIDOS</span>
<table width="600" class="Q" cellspacing="0">
  <tr class="DefaultBkg">
	<td align="center" class='MB'><p class="R"><a href='PedidosEncerrados.asp' onclick="javascript:window.status='Aguarde ...';">Pedidos Encerrados</a></p></td>
  </tr>
  <tr class="DefaultBkg">
	<td align="center"><p class="R"><a href='PedidosEmAndamento.asp' onclick="javascript:window.status='Aguarde ...';">Pedidos Em Andamento</a></p></td>
  </tr>
</table>

<table width="600" class="QS" cellspacing="0">
  <tr class="DefaultBkg">
	<td valign="middle" align="center">
		<p class="C" style="margin: 12px 0px 12px 0px;">Nº PEDIDO&nbsp;&nbsp;<input size="10" maxlength="10" name="pedido_selecionado" id="pedido_selecionado" onkeypress="if (digitou_enter(true) && tem_info(this.value)) {if (normaliza_num_pedido(this.value)!='') this.value=normaliza_num_pedido(this.value); fPED.submit();} filtra_pedido();" onblur="if (normaliza_num_pedido(this.value)!='') this.value=normaliza_num_pedido(this.value);">
				&nbsp;&nbsp;&nbsp;<input name="CONSULTARPED" id="CONSULTARPED" type="submit" class="Botao" 
										 value="CONSULTAR" title="consulta um pedido específico desta loja">
		</p>
	</td>
  </tr>
</table>
</form>


<% if blnQuadroAvisosHabilitado then %>
<!--  ***********************************************************************************************  -->
<!--  O U T R A S   F U N Ç Õ E S                 												       -->
<!--  ***********************************************************************************************  -->
<br />
<form method="post" id="fOF" name="fOF" onsubmit="if (!fOFConcluir(fOF)) return false">
<input type="hidden" name="opcao_selecionada" id="opcao_selecionada" value=''>
<input type="hidden" name="opcao_alerta_se_nao_ha_aviso" id="opcao_alerta_se_nao_ha_aviso" value='S'>
<span class="T">OUTRAS FUNÇÕES</span>
<div class="QFn" align="center" style="width:600px;">
<table class="TFn">
	<tr>
		<td align="left" nowrap>
			<input type="radio" name="rb_op" id="rb_op" value="1" class="CBOX"><span class="rbLink" onclick="fOF.rb_op[0].click(); fOF.bEXECUTAR.click();"
				>Ler Quadro de Avisos (somente não lidos)</span><br>
			<input type="radio" name="rb_op" id="rb_op" value="2" class="CBOX"><span class="rbLink" onclick="fOF.rb_op[1].click(); fOF.bEXECUTAR.click();"
				>Ler Quadro de Avisos (todos os avisos)</span>
			</td>
		</tr>
	</table>

	<span class="R" style="margin: 4 10 0 10">&nbsp;</span>
	<input name="bEXECUTAR" id="bEXECUTAR" type="submit" class="Botao" value="EXECUTAR" title="executa">
	<p class="R" style="margin: 0 10 0 10">&nbsp;</p>

</div>
</form>
<% end if %>

</center>

</body>
</html>

<%
'	FECHA CONEXAO COM O BANCO DE DADOS
'	Obs.: Para que o fechamento seja imediato é necessário acertar
'		  o registro do IIS 4.0, desabilitando o "connection pooling".
'		  Ver artigo no MSDN (ID: Q189410)
	cn.Close
	set cn = nothing
%>
