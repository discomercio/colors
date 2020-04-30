<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/Global.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ===============================================
'	  P E D I D O B L O C O N O T A S N O V O . A S P
'     ===============================================
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


	On Error GoTo 0
	Err.Clear

	dim usuario, pedido_selecionado
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO)

	pedido_selecionado = Trim(request("pedido_selecionado"))
	if (pedido_selecionado = "") then Response.Redirect("aviso.asp?id=" & ERR_PEDIDO_NAO_ESPECIFICADO)

    dim url_origem
    url_origem = Trim(Request("url_origem"))
	
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

	if Not operacao_permitida(OP_CEN_BLOCO_NOTAS_PEDIDO_CADASTRAMENTO, s_lista_operacoes_permitidas) then
		Response.Redirect("aviso.asp?id=" & ERR_NIVEL_ACESSO_INSUFICIENTE)
		end if
		
	dim nivel_acesso_bloco_notas
	nivel_acesso_bloco_notas = Session("nivel_acesso_bloco_notas")
	if Trim(nivel_acesso_bloco_notas) = "" then
		nivel_acesso_bloco_notas = obtem_nivel_acesso_bloco_notas_pedido(cn, usuario)
		Session("nivel_acesso_bloco_notas") = nivel_acesso_bloco_notas
		end if
		
	dim cn, rs
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

	dim r_pedido, r_vendedor, alerta, msg_erro
	dim s, s_ckb_notificar_vendedor_status, s_ckb_notificar_vendedor_msg, s_ckb_notificar_demais_particip_status, s_ckb_notificar_demais_particip_msg
	dim qtde_demais_particip, qtde_demais_particip_com_email, qtde_demais_particip_sem_email
	alerta=""
	if Not le_pedido(pedido_selecionado, r_pedido, msg_erro) then 
		alerta = msg_erro
		end if

	if alerta = "" then
		if Not le_usuario(r_pedido.vendedor, r_vendedor, msg_erro) then
			alerta = msg_erro
			end if
		end if

	s_ckb_notificar_vendedor_status = ""
	s_ckb_notificar_vendedor_msg = ""
	if alerta = "" then
		if Ucase(usuario) = Ucase(r_pedido.vendedor) then
			s_ckb_notificar_vendedor_status = " disabled"
			s_ckb_notificar_vendedor_msg = ""
		else
			if Trim(r_vendedor.email) = "" then
				s_ckb_notificar_vendedor_status = " disabled"
				s_ckb_notificar_vendedor_msg = " (endereço de e-mail não cadastrado)"
				end if
			end if
		end if

	dim i, v_demais_particip
	redim v_demais_particip(0)
	set v_demais_particip(ubound(v_demais_particip)) = new cl_QUATRO_COLUNAS
	v_demais_particip(ubound(v_demais_particip)).c1 = ""

	qtde_demais_particip = 0
	qtde_demais_particip_com_email = 0
	qtde_demais_particip_sem_email = 0
	s_ckb_notificar_demais_particip_status = ""
	s_ckb_notificar_demais_particip_msg = ""
	if alerta = "" then
		s = "SELECT DISTINCT" & _
				" tPBN.usuario," & _
				" tU.email," & _
				" (SELECT Coalesce(Max(nivel_acesso_bloco_notas_pedido), " & Cstr(COD_NIVEL_ACESSO_BLOCO_NOTAS_PEDIDO__NAO_DEFINIDO) & ") AS max_nivel_acesso_bloco_notas_pedido FROM t_PERFIL INNER JOIN t_PERFIL_X_USUARIO ON t_PERFIL.id=t_PERFIL_X_USUARIO.id_perfil WHERE (t_PERFIL_X_USUARIO.usuario = tPBN.usuario)) AS nivel_acesso_bloco_notas_pedido" & _
			" FROM t_PEDIDO_BLOCO_NOTAS tPBN" & _
				" LEFT JOIN t_USUARIO tU ON (tU.usuario = tPBN.usuario)" & _
			" WHERE" & _
				" (pedido = '" & pedido_selecionado & "')" & _
				" AND (tPBN.usuario NOT IN ('" & usuario & "','" & r_pedido.vendedor & "'))"
		if rs.State <> 0 then rs.Close
		rs.open s, cn
		if rs.Eof then
			s_ckb_notificar_demais_particip_status = " disabled"
			s_ckb_notificar_demais_particip_msg = " (não há outros participantes)"
		else
			do while Not rs.Eof
				if v_demais_particip(ubound(v_demais_particip)).c1 <> "" then
					redim preserve v_demais_particip(ubound(v_demais_particip)+1)
					set v_demais_particip(ubound(v_demais_particip)) = new cl_QUATRO_COLUNAS
					end if
				qtde_demais_particip = qtde_demais_particip + 1
				v_demais_particip(ubound(v_demais_particip)).c1 = Trim("" & rs("usuario"))
				v_demais_particip(ubound(v_demais_particip)).c2 = rs("nivel_acesso_bloco_notas_pedido")
				v_demais_particip(ubound(v_demais_particip)).c3 = Trim("" & rs("email"))
				if Trim("" & rs("email")) = "" then
					v_demais_particip(ubound(v_demais_particip)).c4 = "Usuário '" & Trim("" & rs("usuario")) & "' participante do bloco de notas não possui e-mail cadastrado!"
					qtde_demais_particip_sem_email = qtde_demais_particip_sem_email + 1
				else
					v_demais_particip(ubound(v_demais_particip)).c4 = ""
					qtde_demais_particip_com_email = qtde_demais_particip_com_email + 1
					end if
				rs.MoveNext
				loop
			
			if qtde_demais_particip_com_email = 0 then
				s_ckb_notificar_demais_particip_status = " disabled"
				s_ckb_notificar_demais_particip_msg = " (demais participantes não possuem e-mail cadastrado)"
				end if
			end if
		end if

	dim strJscript
	strJscript = "<script language='JavaScript' type='text/javascript'>" & vbCrLf & _
					"	var vendedor_nivel_acesso_bloco_notas = " & r_vendedor.nivel_acesso_bloco_notas_pedido & ";" & vbCrLf & _
					"	var vDemaisParticip = new Array();" & vbCrLf & _
					"	vDemaisParticip[0] = new oUsuario(" & chr(34) & chr(34) & "," & chr(34) & chr(34) & "," & chr(34) & chr(34) & "," & chr(34) & chr(34) & ");" & vbCrLf

	if qtde_demais_particip > 0 then
		for i=LBound(v_demais_particip) to UBound(v_demais_particip)
			if Trim("" & v_demais_particip(i).c1) <> "" then
				strJscript = strJscript & _
								"	vDemaisParticip[vDemaisParticip.length] = new oUsuario(" & _
										chr(34) & v_demais_particip(i).c1 & chr(34) & _
										"," & chr(34) & v_demais_particip(i).c2 & chr(34) & _
										"," & chr(34) & v_demais_particip(i).c3 & chr(34) & _
										"," & chr(34) & v_demais_particip(i).c4 & chr(34) & _
										");" & vbCrLf
				end if
			next
		end if

	strJscript = strJscript & _
				"</script>" & vbCrLf
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
	<title>CENTRAL<%=MontaNumPedidoExibicaoTitleBrowser(pedido_selecionado)%></title>
	</head>



<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>
<script src="<%=URL_FILE__AJAX_JS%>" Language="JavaScript" Type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
	function oUsuario(usuario, nivel_acesso_bloco_notas_pedido, email, msg_alerta) {
		this.usuario = usuario;
		this.nivel_acesso_bloco_notas_pedido = nivel_acesso_bloco_notas_pedido;
		this.email = email;
		this.msg_alerta = msg_alerta;
	}
</script>

<% =strJscript %>

<script language="JavaScript" type="text/javascript">
function calcula_tamanho_restante() {
	var f, s;
	f = fPED;
	s = "" + fPED.c_mensagem.value;
	f.c_tamanho_restante.value = MAX_TAM_MENSAGEM_BLOCO_NOTAS - s.length;
}

function isNivelAcessoOk(f) {
var msg_demais_particip, msg_demais_particip_sem_email, msg_demais_particip_sem_acesso;
	if (f.ckb_notificar_vendedor.checked) {
		if (converte_numero(f.c_nivel_acesso_bloco_notas.value) > vendedor_nivel_acesso_bloco_notas) {
			alert("Não é possível enviar a notificação por e-mail para o vendedor porque ele não possui o nível de acesso necessário!");
			return false;
		}
	}

	if (f.ckb_notificar_demais_particip.checked) {
		msg_demais_particip = "";
		msg_demais_particip_sem_email = "";
		msg_demais_particip_sem_acesso = "";
		for (var i = 0; i < vDemaisParticip.length; i++) {
			if (vDemaisParticip[i].usuario != "") {
				if (vDemaisParticip[i].email == "") {
					if (msg_demais_particip_sem_email != "") msg_demais_particip_sem_email += "\n";
					msg_demais_particip_sem_email += "Usuário '" + vDemaisParticip[i].usuario + "' não possui e-mail cadastrado!";
				}
				if (converte_numero(f.c_nivel_acesso_bloco_notas.value) > converte_numero(vDemaisParticip[i].nivel_acesso_bloco_notas_pedido)) {
					if (msg_demais_particip_sem_acesso != "") msg_demais_particip_sem_acesso += "\n";
					msg_demais_particip_sem_acesso += "Usuário '" + vDemaisParticip[i].usuario + "' não irá receber o e-mail por não possuir nível de acesso!";
				}
			}

			if ((msg_demais_particip_sem_email != "") && (msg_demais_particip_sem_acesso != "")) {
				msg_demais_particip = msg_demais_particip_sem_email + "\n" + msg_demais_particip_sem_acesso;
			}
			else {
				msg_demais_particip = msg_demais_particip_sem_email + msg_demais_particip_sem_acesso;
			}

			if (msg_demais_particip != "") {
				msg_demais_particip = "Foram detectadas as seguintes pendências:" + "\n\n" + msg_demais_particip + "\n\nDeseja continuar mesmo assim?";
				if (!confirm(msg_demais_particip)) return false;
			}
		}
	}

	return true;
}

function fPEDBlocoNotasNovoConfirma(f) {
var s;

	s = "" + f.c_mensagem.value;
	if (s.length == 0) {
		alert('É necessário escrever o texto da mensagem!!');
		f.c_mensagem.focus();
		return;
		}
	if (s.length > MAX_TAM_MENSAGEM_BLOCO_NOTAS) {
		alert('Conteúdo da mensagem excede em ' + (s.length - MAX_TAM_MENSAGEM_BLOCO_NOTAS) + ' caracteres o tamanho máximo de ' + MAX_TAM_MENSAGEM_BLOCO_NOTAS + '!!');
		f.c_mensagem.focus();
		return;
		}

	if (trim(f.c_nivel_acesso_bloco_notas.value) == "") {
		alert('É necessário definir o nível de acesso para a leitura da mensagem!!');
		f.c_nivel_acesso_bloco_notas.focus();
		return;
		}

	if (!isNivelAcessoOk(f)) return;

	dCONFIRMA.style.visibility="hidden";
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

<% if alerta <> "" then %>
<!-- ************************************************************ -->
<!-- **********  PÁGINA PARA EXIBIR MENSAGENS DE ERRO  ********** -->
<!-- ************************************************************ -->
<body onload="bVOLTAR.focus();">
<center>
<br>
<!--  T E L A  -->
<p class="T">A V I S O</p>
<div class="MtAlerta" style="width:600px;font-weight:bold;" align="center"><p style='margin:5px 2px 5px 2px;'><%=alerta%></p></div>
<br><br>
<p class="TracoBottom"></p>
<table cellspacing="0">
<tr>
	<td align="center"><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()"><img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
</tr>
</table>
</center>
</body>




<% else %>
<body onload="fPED.c_mensagem.focus();">
<center>

<form id="fPED" name="fPED" method="post" action="PedidoBlocoNotasNovoConfirma.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="pedido_selecionado" id="pedido_selecionado" value='<%=pedido_selecionado%>'>
<input type="hidden" name="url_origem" id="url_origem" value="<%=url_origem%>" />

<!--  I D E N T I F I C A Ç Ã O   D O   P E D I D O -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">

<tr>
	<td align="left" valign="bottom"><p class="PEDIDO">Bloco de Notas</p></td>
	<td align="right" valign="bottom"><p class="PEDIDO" style="font-size:14pt;">Pedido <%=pedido_selecionado%></p></td>
</tr>
</table>
<br>

<table>
<tr>
	<td align="right" valign="bottom">
		<span class="PLLd">Tamanho restante:</span><input name="c_tamanho_restante" id="c_tamanho_restante" tabindex=-1 readonly class="TA" style="width:35px;text-align:right;" value='<%=Cstr(MAX_TAM_MENSAGEM_BLOCO_NOTAS)%>' />
	</td>
</tr>
<tr>
	<td>
	<table class="Q" style="width:649px;" cellSpacing="0">
		<tr>
			<td><p class="Rf">MENSAGEM</p>
				<textarea name="c_mensagem" id="c_mensagem" class="PLLe" rows="<%=Cstr(MAX_LINHAS_MENSAGEM_BLOCO_NOTAS)%>" 
					style="width:642px;margin-left:2pt;" onkeypress="limita_tamanho(this,MAX_TAM_MENSAGEM_BLOCO_NOTAS);" onblur="this.value=trim(this.value);calcula_tamanho_restante();"
					onkeyup="calcula_tamanho_restante();"
					></textarea>
			</td>
		</tr>
	</table>
	</td>
</tr>
<% if converte_numero(nivel_acesso_bloco_notas) = converte_numero(COD_NIVEL_ACESSO_BLOCO_NOTAS_PEDIDO__PUBLICO) then %>
<input type="hidden" name="c_nivel_acesso_bloco_notas" id="c_nivel_acesso_bloco_notas" value='<%=Cstr(COD_NIVEL_ACESSO_BLOCO_NOTAS_PEDIDO__PUBLICO)%>'>
<% else %>
<tr>
	<td>
		<br />
		<p class="Rf">NÍVEL DE ACESSO PARA LEITURA</p>
		<select id="c_nivel_acesso_bloco_notas" name="c_nivel_acesso_bloco_notas" style="margin-top:3px;margin-bottom:4px;width:180px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) {this.options[0].selected=true;}">
		<% =nivel_acesso_bloco_notas_pedido_monta_itens_select(Null, nivel_acesso_bloco_notas) %>
		</select>
	</td>
</tr>
<% end if %>

<tr>
	<td>
		<br />
		<p class="Rf">ENVIAR NOTIFICAÇÃO POR E-MAIL</p>
		<input type="checkbox" name="ckb_notificar_vendedor" id="ckb_notificar_vendedor" value="ON" <%=s_ckb_notificar_vendedor_status%> /><span class="C" style="cursor:default;" onclick="fPED.ckb_notificar_vendedor.click();">Notificar o vendedor</span>
			<span style="font-size:8pt;font-style:italic;color:red;" name="spnVendedorMsg" id="spnVendedorMsg"><%=s_ckb_notificar_vendedor_msg%></span>
		<!-- FUNCIONALIDADE DESABILITADA -->
		<div style="display:none;">
		<br />
		<input type="checkbox" name="ckb_notificar_demais_particip" id="ckb_notificar_demais_particip" value="ON" <%=s_ckb_notificar_demais_particip_status%> /><span class="C" style="cursor:default;" onclick="fPED.ckb_notificar_demais_particip.click();">Notificar demais participantes</span>
			<span style="font-size:8pt;font-style:italic;color:red;" name="spnDemaisParticipMsg" id="spnDemaisParticipMsg"><%=s_ckb_notificar_demais_particip_msg%></span>
		</div>
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
	<td><a name="bCANCELA" id="bCANCELA" href="javascript:history.back()" title="cancela o cadastramento de nova mensagem no bloco de notas">
		<img src="../botao/cancelar.gif" width="176" height="55" border="0"></a></td>
	<td align="right"><div name="dCONFIRMA" id="dCONFIRMA">
		<a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fPEDBlocoNotasNovoConfirma(fPED)" title="grava a mensagem no bloco de notas">
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
	if rs.State <> 0 then rs.Close
	set rs = nothing

'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>