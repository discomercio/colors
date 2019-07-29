<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/Global.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =======================================================
'	  RelAcompanhamentoChamadosFiltro.asp
'     =======================================================
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

	dim usuario
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs, msg_erro
    dim s
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

	if Not operacao_permitida(OP_CEN_PEDIDO_CHAMADO_CADASTRAMENTO, s_lista_operacoes_permitidas) then
		Response.Redirect("aviso.asp?id=" & ERR_NIVEL_ACESSO_INSUFICIENTE)
		end if
		
	dim intIdx, blnHasDepto, blnMostraOpcaoTodos
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
	<title>CENTRAL</title>
	</head>



<script src="<%=URL_FILE__JQUERY%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_I18N%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>

<script type="text/javascript">
	$(function() {
	    $("#c_dt_cad_chamado_inicio").hUtilUI('datepicker_filtro_inicial');
		$("#c_dt_cad_chamado_termino").hUtilUI('datepicker_filtro_final');
	});
</script>

<script language="JavaScript" type="text/javascript">
function fFILTROConfirma(f) {
var i, blnFlag;
var s_de, s_ate;

    if (trim(f.c_dt_cad_chamado_inicio.value) != "") {
        if (!isDate(f.c_dt_cad_chamado_inicio)) {
            alert("Data de início inválida!!");
            f.c_dt_cad_chamado_inicio.focus();
            return;
        }
    }

    if (trim(f.c_dt_cad_chamado_termino.value) != "") {
        if (!isDate(f.c_dt_cad_chamado_termino)) {
            alert("Data de término inválida!!");
            f.c_dt_cad_chamado_termino.focus();
            return;
        }
    }

    s_de = trim(f.c_dt_cad_chamado_inicio.value);
    s_ate = trim(f.c_dt_cad_chamado_termino.value);
    if ((s_de != "") && (s_ate != "")) {
        s_de = retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(s_de));
        s_ate = retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(s_ate));
        if (s_de > s_ate) {
            alert("Data de término é menor que a data de início!!");
            f.c_dt_cad_chamado_termino.focus();
            return;
        }
    }

	blnFlag = false;
	for (i = 0; i < f.rb_status.length; i++) {
		if (f.rb_status[i].checked) {
			blnFlag = true;
			break;
		}
	}
	if (!blnFlag) {
		alert("Selecione o status do chamado!!");
		return;
	}
	
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
<link href="<%=URL_FILE__JQUERY_UI_CSS%>" rel="stylesheet" type="text/css">

<body onload="focus();">
<center>

<form id="fFILTRO" name="fFILTRO" method="post" action="RelAcompanhamentoChamadosExec.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Acompanhamento de Chamados</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>
<br>

<table class="Qx" cellSpacing="0" style="width:340px;">
<!--  PERÍODO: DATA DO CHAMADO  -->
	<tr>
		<td class="MT PLTe" NOWRAP style="background-color:whitesmoke;" valign="bottom"><span class="PLTe">PERÍODO DE ABERTURA DO CHAMADO</span></td>
	</tr>
	<tr bgColor="#FFFFFF" NOWRAP>
		<td class="ME MD">
			<table cellSpacing="0" cellPadding="0" style="margin:0px 20px 6px 10px;">
			<tr bgColor="#FFFFFF">
				<td>
					<input class="Cc" maxlength="10" style="width:70px;" name="c_dt_cad_chamado_inicio" id="c_dt_cad_chamado_inicio" onblur="if (!isDate(this)) {alert('Data de início inválida!'); this.focus();}" onkeypress="if (digitou_enter(true)) fFILTRO.c_dt_cad_chamado_termino.focus(); filtra_data();"
						>&nbsp;
					<span class="C">&nbsp;até&nbsp;</span>&nbsp;
					<input class="Cc" maxlength="10" style="width:70px;" name="c_dt_cad_chamado_termino" id="c_dt_cad_chamado_termino" onblur="if (!isDate(this)) {alert('Data de término inválida!'); this.focus();}" onkeypress="if (digitou_enter(true)) bCONFIRMA.focus(); filtra_data();">
				</td>
			</tr>
			</table>
		</td>
	</tr>

<!--  STATUS DO CHAMADO  -->
	<tr>
		<td class="MT PLTe" NOWRAP style="background-color:whitesmoke;" valign="bottom">&nbsp;STATUS DO CHAMADO</td>
	</tr>
	<tr bgColor="#FFFFFF" NOWRAP>
		<td class="ME MB MD">
			<% intIdx=-1 %>
			<input type="radio" id="rb_status" name="rb_status" value="ABERTO" class="CBOX" style="margin-left:20px;">
			<% intIdx=intIdx+1 %>
			<span style="cursor:default;" class="rbLink" onclick="fFILTRO.rb_status[<%=Cstr(intIdx)%>].click();">Aberto</span>
			<br>
			<input type="radio" id="rb_status" name="rb_status" value="EM_ANDAMENTO" class="CBOX" style="margin-left:20px;">
			<% intIdx=intIdx+1 %>
			<span style="cursor:default;" class="rbLink" onclick="fFILTRO.rb_status[<%=Cstr(intIdx)%>].click();">Em Andamento</span>
			<br>
            <input type="radio" id="rb_status" name="rb_status" value="FINALIZADO" class="CBOX" style="margin-left:20px;">
			<% intIdx=intIdx+1 %>
			<span style="cursor:default;" class="rbLink" onclick="fFILTRO.rb_status[<%=Cstr(intIdx)%>].click();">Finalizado</span>
			<br>
			<input type="radio" id="rb_status" name="rb_status" value="" checked class="CBOX" style="margin-left:20px;">
			<% intIdx=intIdx+1 %>
			<span style="cursor:default;" class="rbLink" onclick="fFILTRO.rb_status[<%=Cstr(intIdx)%>].click();">Todos</span>
		</td>
	</tr>
<!--  POSIÇÃO  -->
    <%  blnHasDepto = False
        blnMostraOpcaoTodos = False
        s = "select * from t_PEDIDO_CHAMADO_DEPTO where usuario_responsavel='" & usuario & "' or usuario_gestor='" & usuario & "'"  
        rs.Open s, cn
        if Not rs.Eof then blnHasDepto = True%>
	<tr>
		<td class="MDBE PLTe" NOWRAP style="background-color:whitesmoke;" valign="bottom">&nbsp;SELECIONAR CHAMADOS</td>
	</tr>
	<tr bgColor="#FFFFFF" NOWRAP>
		<td class="ME MB MD">
			<% intIdx=-1 %>
			<input type="radio" id="rb_posicao" name="rb_posicao" value="USUARIO_TX" class="CBOX" style="margin-left:20px;" <%if Not blnHasDepto then Response.Write "checked" %> />
			<% intIdx=intIdx+1 %>
			<span style="cursor:default;" class="rbLink" onclick="fFILTRO.rb_posicao[<%=Cstr(intIdx)%>].click();">Abertos por mim</span>
			<br>  

    <% if blnHasDepto then %>
			<input type="radio" id="rb_posicao" name="rb_posicao" value="USUARIO_RX" class="CBOX" style="margin-left:20px;">
			<% intIdx=intIdx+1
               blnMostraOpcaoTodos = True %>
			<span style="cursor:default;" class="rbLink" onclick="fFILTRO.rb_posicao[<%=Cstr(intIdx)%>].click();">Destinados ao meu departamento</span>	
            <br>		
    <% end if %>

    <% if blnHasDepto Or operacao_permitida(OP_CEN_PEDIDO_CHAMADO_ESCREVER_MSG_QUALQUER_CHAMADO, s_lista_operacoes_permitidas) then %>
            <input type="radio" id="rb_posicao" name="rb_posicao" value="USUARIO_MSG" class="CBOX" style="margin-left:20px;" />
			<% intIdx=intIdx+1
               blnMostraOpcaoTodos = True %>
			<span style="cursor:default;" class="rbLink" onclick="fFILTRO.rb_posicao[<%=Cstr(intIdx)%>].click();">Em que interagi com mensagens</span>
			<br>
    <% end if %>
            
    <% if blnMostraOpcaoTodos then %>
			<input type="radio" id="rb_posicao" name="rb_posicao" value="" checked class="CBOX" style="margin-left:20px;">
			<% intIdx=intIdx+1 %>
			<span style="cursor:default;" class="rbLink" onclick="fFILTRO.rb_posicao[<%=Cstr(intIdx)%>].click();">Todos</span>

    <% end if %>
		</td>
	</tr>
<!--  MOTIVO ABERTURA  -->
	<tr>
		<td class="ME MD MB PLTe" NOWRAP style="background-color:whitesmoke;" valign="bottom">&nbsp;MOTIVO DA ABERTURA DO CHAMADO</td>
	</tr>
	<tr bgColor="#FFFFFF" NOWRAP>
		<td class="ME MB MD">
			<select id="c_motivo_abertura" name="c_motivo_abertura" style="width:450px;margin:3px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
			<%=codigo_descricao_monta_itens_select(GRUPO_T_CODIGO_DESCRICAO__CHAMADOS_EM_PEDIDOS__MOTIVO_ABERTURA, "")%>
			</select>
		</td>
	</tr>
</table>

<!-- ************   SEPARADOR   ************ -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>
<br>


<table width="649" cellSpacing="0">
<tr>
	<td><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()" title="volta para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
	<td align="right"><div name="dCONFIRMA" id="dCONFIRMA"><a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fFILTROConfirma(fFILTRO)" title="executa a consulta">
		<img src="../botao/confirmar.gif" width="176" height="55" border="0"></a></div>
	</td>
</tr>
</table>
</form>

</center>
</body>
</html>

<%
    if rs.State <> 0 then rs.Close
	set rs = nothing

'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>
