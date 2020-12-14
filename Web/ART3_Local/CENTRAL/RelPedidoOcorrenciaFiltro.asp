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
'	  RelPedidoOcorrenciaFiltro.asp
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
	dim cn
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

	if Not operacao_permitida(OP_CEN_REL_OCORRENCIAS_EM_PEDIDOS, s_lista_operacoes_permitidas) then
		Response.Redirect("aviso.asp?id=" & ERR_NIVEL_ACESSO_INSUFICIENTE)
		end if
		
	dim intIdx
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
	<title>CENTRAL</title>
	</head>



<% if False then 'APENAS P/ HABILITAR O INTELLISENSE DURANTE O DESENVOLVIMENTO!! %>
<script src="../Global/jquery.js" language="JavaScript" type="text/javascript"></script>
<% end if %>

<script src="<%=URL_FILE__JQUERY%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_MY_GLOBAL%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>
<script src="<%=URL_FILE__AJAX_JS%>" language="JavaScript" type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
function fFILTROConfirma(f) {
var i, blnFlag;
	blnFlag = false;
	for (i = 0; i < f.rb_status.length; i++) {
		if (f.rb_status[i].checked) {
			blnFlag = true;
			break;
		}
	}
	if (!blnFlag) {
		alert("Selecione o status da ocorrência!!");
		return;
	}
	
	dCONFIRMA.style.visibility="hidden";
	window.status = "Aguarde ...";
	f.submit();
}
</script>

<script type="text/javascript">
    function geraArquivoXLS(f) {
        var serverVariableUrl, strUrl, xmlHttp;
        var i, oc_status, transportadora, loja;
        var s_de, s_ate, s_hoje;
        loja = $("#c_loja").val();

        var idxStatus = $("input[name=rb_status]:checked").index();
        if (idxStatus == 0) {
            oc_status = "ABERTA";
        }
        else if (idxStatus == 3) {
            oc_status = "EM_ANDAMENTO";
        }
        else {
            oc_status = "AMBOS";
        }

        serverVariableUrl = '<%=Request.ServerVariables("URL")%>';
        serverVariableUrl = serverVariableUrl.toUpperCase();
        serverVariableUrl = serverVariableUrl.substring(0, serverVariableUrl.indexOf("CENTRAL"));

        xmlhttp = GetXmlHttpObject();
        if (xmlhttp == null) {
            alert("O browser NÃO possui suporte ao AJAX!!");
            return;
        }

        window.status = "Aguarde, gerando arquivo ...";
        divMsgAguardeObtendoDados.style.visibility = "";

		strUrl = '<%=getProtocoloEmUsoHttpOrHttps%>://<%=Request.ServerVariables("SERVER_NAME")%>:<%=Request.ServerVariables("SERVER_PORT")%>' + serverVariableUrl + 'WebAPI/api/Relatorios/GetXLSReport2/';
        strUrl = strUrl + '?usuario=<%=usuario%>';
        strUrl = strUrl + '&oc_status=' + oc_status;
        strUrl = strUrl + '&loja=' + loja;
        strUrl = strUrl + '&transportadora=' + $("#c_transportadora").val();


        xmlhttp.onreadystatechange = function () {
            var xmlResp;

            if (xmlhttp.readyState == AJAX_REQUEST_IS_COMPLETE) {
                xmlResp = JSON.parse(xmlhttp.responseText);

                if (xmlResp.Status == "OK") {

					gerarRelatorio.action = '<%=getProtocoloEmUsoHttpOrHttps%>://<%=Request.ServerVariables("SERVER_NAME")%>:<%=Request.ServerVariables("SERVER_PORT")%>' + serverVariableUrl + 'WebAPI/api/Relatorios/downloadXLS2/?fileName=' + xmlResp.fileName;
                    gerarRelatorio.submit();

                    window.status = "Concluído";
                    divMsgAguardeObtendoDados.style.visibility = "hidden";
                }
                else if (xmlResp.Status == "Falha") {
                    window.status = "Concluído";
                    divMsgAguardeObtendoDados.style.visibility = "hidden";

                    alert("Falha ao gerar o arquivo XLS\n" + xmlResp.Exception);
                    return;
                }
                else if (xmlResp.Status == "Vazio") {
                    window.status = "Concluído";
                    divMsgAguardeObtendoDados.style.visibility = "hidden";

                    alert(xmlResp.Exception);
                    return;
                }
            }
        }

        xmlhttp.open("POST", strUrl, true);
        xmlhttp.send();

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

<!-- MENSAGEM: "Aguarde, obtendo dados" -->

	<div id="divMsgAguardeObtendoDados" name="divMsgAguardeObtendoDados" style="background-image: url('../Imagem/ajax_loader_gray_256.gif');background-repeat:no-repeat;background-position: center center;position:absolute;bottom:0px;left:0px;width:100%;height:100%;z-index:9;border: 1pt solid #C0C0C0;background-color: black;opacity:0.6;visibility:hidden;vertical-align: middle">

	</div>

    <form name="gerarRelatorio" id="gerarRelatorio" method="POST">
    <input type="hidden" name="idRel" id="idRel" value="" />
    </form>
<form id="fFILTRO" name="fFILTRO" method="post" action="RelPedidoOcorrencia.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Relatório de Ocorrências</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>
<br>

<table class="Qx" cellSpacing="0" style="width:240px;">
<!--  STATUS DA OCORRÊNCIA  -->
	<tr>
		<td class="MT PLTe" NOWRAP style="background-color:whitesmoke;" valign="bottom">&nbsp;STATUS DA OCORRÊNCIA</td>
	</tr>
	<tr bgColor="#FFFFFF" NOWRAP>
		<td class="ME MB MD" style="padding:5px;">
			<% intIdx=-1 %>
			<input type="radio" id="rb_status" name="rb_status" value="ABERTA" class="CBOX" style="margin-left:20px;">
			<% intIdx=intIdx+1 %>
			<span style="cursor:default;" class="rbLink" onclick="fFILTRO.rb_status[<%=Cstr(intIdx)%>].click();">Aberta</span>
			<br>
			<input type="radio" id="rb_status" name="rb_status" value="EM_ANDAMENTO" class="CBOX" style="margin-left:20px;">
			<% intIdx=intIdx+1 %>
			<span style="cursor:default;" class="rbLink" onclick="fFILTRO.rb_status[<%=Cstr(intIdx)%>].click();">Em Andamento</span>
			<br>
			<input type="radio" id="rb_status" name="rb_status" value="" checked class="CBOX" style="margin-left:20px;">
			<% intIdx=intIdx+1 %>
			<span style="cursor:default;" class="rbLink" onclick="fFILTRO.rb_status[<%=Cstr(intIdx)%>].click();">Ambos</span>
		</td>
	</tr>
<!--  LOJA  -->
	<tr>
		<td class="ME MB MD PLTe" NOWRAP style="background-color:whitesmoke;" valign="bottom">&nbsp;LOJA</td>
	</tr>
	<tr bgColor="#FFFFFF" NOWRAP>
		<td class="ME MB MD" style="padding:5px;">
		<input name="c_loja" id="c_loja" class="PLLe" maxlength="3" style="width:100px;margin-left:4px;font-size:12pt;"
				onkeypress="if (digitou_enter(true)&&tem_info(this.value)) $(this).hUtil('focusNext'); filtra_numerico();"
				onblur="this.value=normaliza_codigo(this.value,TAM_MIN_LOJA);">
		</td>
	</tr>
<!--  TRANSPORTADORA  -->
	<tr>
		<td class="ME MB MD PLTe" NOWRAP style="background-color:whitesmoke;" valign="bottom">&nbsp;TRANSPORTADORA</td>
	</tr>
	<tr bgColor="#FFFFFF" NOWRAP>
		<td class="ME MB MD" style="padding:5px;">
			<select id="c_transportadora" name="c_transportadora" style="margin:1px 10px 6px 10px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
			<% =transportadora_monta_itens_select(Null) %>
			</select>
		</td>
	</tr>
<!-- ArquivoXLS -->
    <tr bgColor="#FFFFFF" NOWRAP>
        <td class="ME MB MD" >
            <a href="javascript:geraArquivoXLS(gerarRelatorio);" class="C" style="margin-left:0px;"><div class="Button" style="width:150px;margin-left:10px;margin-top:10px; padding:3px;color:black;text-align:center;">Arquivo XLS</div></a></td>
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
	<td><a name="bVOLTAR" id="bVOLTAR" href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="volta para a página anterior">
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
'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>
