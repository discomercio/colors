<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =========================================
'	  Q U A D R O A V I S O M O S T R A . A S P
'     =========================================
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
	dim s, usuario
	usuario = trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 
	
'	CONECTA COM O BANCO DE DADOS
	dim cn, r, msg_erro
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	If Not cria_recordset_otimista(r, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

	dim opcao_selecionada, opcao_alerta_se_nao_ha_aviso, alerta
	opcao_selecionada=Trim(request("opcao_selecionada"))
	opcao_alerta_se_nao_ha_aviso=Trim(request("opcao_alerta_se_nao_ha_aviso"))

	Dim iMsg, nMsg
	Dim vMsg()

	alerta = ""
	if opcao_selecionada = "" then
		if Not recupera_avisos_nao_lidos("", usuario, vMsg) then 
			if opcao_alerta_se_nao_ha_aviso="" then
				Response.Redirect("resumo.asp" & "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")))
			else
				alerta = "NÃO HÁ AVISOS."
				end if
			end if
	else
		if Not recupera_avisos("", usuario, vMsg) then 
			if opcao_alerta_se_nao_ha_aviso="" then
				Response.Redirect("resumo.asp" & "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")))
			else
				alerta = "NÃO HÁ AVISOS."
				end if
			end if
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

configura_painel();

function trataBodyKeypress() {
var i;
	keys += String.fromCharCode(window.event.keyCode);
	if (isSelectAllCheckBoxesKeywordOk(keys)) {
		for (i = 1; i < fCAD["xMsg"].length; i++) {
			if (!fCAD["xMsg"][i].disabled) fCAD["xMsg"][i].checked = true;
		}
	}
}

function RemoveAviso( f ) {
var i, max;
	max=f["xMsg"].length;
	f.aviso_selecionado.value="";
	for (i=0; i < max; i++) {
		if (f["xMsg"][i].checked) {
			if (f["xMsg"][i].value!="") {
				if (f.aviso_selecionado.value!="") f.aviso_selecionado.value=f.aviso_selecionado.value + "|";
				f.aviso_selecionado.value=f.aviso_selecionado.value+f["xMsg"][i].value;
				}
			}
		}

	if (f.aviso_selecionado.value=="") {
		alert("Nenhum aviso selecionado!!");
		return;
		}
		
	dREMOVE.style.visibility="hidden";
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


<body onload="focus();" onkeypress="trataBodyKeypress();">
<center>



<!--  QUADRO DE AVISOS -->

<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
<%
	s = "Quadro de Avisos"
%>
	<td align="center" valign="bottom"><p class="T" style="font-size:20pt;"><%=s%><br><span class="C">&nbsp;</span></p></td>
</tr>
</table>
<br>


<!--  CAMPOS  -->
<form id="fCAD" name="fCAD" method="post" action="QuadroAvisoLido.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="aviso_selecionado" id="aviso_selecionado" value=''>

<!-- FORÇA A CRIAÇÃO DE UM ARRAY DE CHECK BOXES MESMO QUANDO EXISTE SOMENTE 1 MENSAGEM -->
<input type="hidden" class="CBOX" name="xMsg" id="xMsg" value="">

<% if alerta = "" then %>
	<%	nMsg = 0
		for iMsg=Lbound(vMsg) to Ubound(vMsg)
		nMsg = nMsg + 1 %>

    <%  '   MARCAR COMO EXIBIDO

		if Err = 0 then
		'	VERIFICA SE O AVISO AINDA EXISTE E OBTÉM INFO P/ LOG
			r.Open "SELECT * FROM t_AVISO WHERE (id='" & Trim("" & vMsg(iMsg).id_aviso) & "')", cn
			if Err = 0 then
				if Not r.EOF then
					r.Close 
				'	VERIFICA SE JÁ NÃO ESTÁ MARCADO COMO EXIBIDO
					r.Open "SELECT * FROM t_AVISO_EXIBIDO WHERE (id='" & Trim("" & vMsg(iMsg).id_aviso) & "') AND (usuario='" & usuario & "')", cn
					if Err = 0 then
						if r.EOF then
						'	MARCA COMO LIDO
							r.AddNew
							if Err = 0 then
								r("id") = vMsg(iMsg).id_aviso
								r("usuario") = usuario
								r("dt_hr_ult_exibicao") = Now
								r.Update 
							end if
						else
							if Err = 0 then
								r("dt_hr_ult_exibicao") = Now
								r.Update 
							end if
                        end if
					end if
				r.Close 
				end if
			end if
		end if

    %>




	<!-- ************   MENSAGEM   ************ -->
	<table width="649" cellSpacing="0">
		<tr><td width="100%">
		<span class="Lbl">Divulgado em:&nbsp;&nbsp;<%=formata_data_hora(vMsg(iMsg).dt_ult_atualizacao)%>
		<%if (vMsg(iMsg).lido<>"") And (Not IsNull(vMsg(iMsg).dt_lido)) then %>
		&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
		<span class="Lbl" style="color:#6699FF;">Lido em:&nbsp;&nbsp;<%=formata_data_hora(vMsg(iMsg).dt_lido)%></span>
		<%end if%>
		</span>
		<textarea id="mensagem" name="mensagem" class="QuadroAviso" readonly onkeypress="filtra_nome_identificador();"><%=vMsg(iMsg).mensagem%></textarea>
		<br><input type="checkbox" class="CBOX" name="xMsg" id="xMsg" value="<%=vMsg(iMsg).id_aviso%>" <%if vMsg(iMsg).lido<>"" then Response.Write " disabled"%>><span class="CBOX" style="cursor:default;<%if vMsg(iMsg).lido<>"" then Response.Write " color:#808080;"%>" onclick="fCAD.xMsg[<%=Cstr(nMsg)%>].click();">Não exibir mais este aviso</span>
		</td></tr>
	</table>

	<%if iMsg<>Ubound(vMsg) then Response.Write "<br><br><br>"%>

    


	<% next %>

<% else %>

	<div class='MtAviso' style="width:400px;font-weight:bold;" align="center">
	<p style='margin:5px 2px 5px 2px;'><%=alerta%></p></div>
	
<% end if %>


<!-- ************   SEPARADOR   ************ -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>
<br>

<table class="notPrint" width="649" cellSpacing="0">
<tr>
	<% if alerta = "" then s="'left'" else s="'center'" %>
	<td align=<%=s%>><a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="cancela as alterações">
		<img src="../botao/cancelar.gif" width="176" height="55" border="0"></a></td>

<% if alerta = "" then %>
	<td align="right"><div name='dREMOVE' id='dREMOVE'>
		<a name="bREMOVE" id="bREMOVE" href="javascript:RemoveAviso(fCAD)" title="apaga as mensagens selecionadas">
		<img src="../botao/remover.gif" width="176" height="55" border="0"></a></div>
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
	cn.Close
	set cn = nothing
%>