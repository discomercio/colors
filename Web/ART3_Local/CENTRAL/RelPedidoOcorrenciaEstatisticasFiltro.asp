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
'	  RelPedidoOcorrenciaEstatisticasFiltro.asp
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

	if Not operacao_permitida(OP_CEN_REL_ESTATISTICAS_OCORRENCIAS_EM_PEDIDOS, s_lista_operacoes_permitidas) then
		Response.Redirect("aviso.asp?id=" & ERR_NIVEL_ACESSO_INSUFICIENTE)
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



<html>


<head>
	<title>CENTRAL</title>
	</head>

<script src="<%=URL_FILE__JQUERY%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__AJAX_JS%>" language="JavaScript" type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
function fFILTROConfirma(f) {
var s_de, s_ate;

	if (trim(f.c_dt_cad_ocorrencia_inicio.value) == "") {
		alert("Informe a data de início do período!!");
		f.c_dt_cad_ocorrencia_inicio.focus();
		return;
	}

	if (trim(f.c_dt_cad_ocorrencia_termino.value) == "") {
		alert("Informe a data de término do período!!");
		f.c_dt_cad_ocorrencia_termino.focus();
		return;
	}

	if (trim(f.c_dt_cad_ocorrencia_inicio.value) != "") {
		if (!isDate(f.c_dt_cad_ocorrencia_inicio)) {
			alert("Data de início inválida!!");
			f.c_dt_cad_ocorrencia_inicio.focus();
			return;
		}
	}

	if (trim(f.c_dt_cad_ocorrencia_termino.value) != "") {
		if (!isDate(f.c_dt_cad_ocorrencia_termino)) {
			alert("Data de término inválida!!");
			f.c_dt_cad_ocorrencia_termino.focus();
			return;
		}
	}

	s_de = trim(f.c_dt_cad_ocorrencia_inicio.value);
	s_ate = trim(f.c_dt_cad_ocorrencia_termino.value);
	if ((s_de != "") && (s_ate != "")) {
		s_de = retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(s_de));
		s_ate = retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(s_ate));
		if (s_de > s_ate) {
			alert("Data de término é menor que a data de início!!");
			f.c_dt_cad_ocorrencia_termino.focus();
			return;
		}
	}
	
	dCONFIRMA.style.visibility="hidden";
	window.status = "Aguarde ...";
	f.submit();
}
</script>
<script type="text/javascript">
    function geraArquivoXLS(f) {
        var serverVariableUrl, strUrl, xmlHttp;
        var i, dt_inicio, dt_termino, motivo_ocorrencia, tp_ocorrencia, transportadora, vendedor, indicador, uf, loja;
        var s_de, s_ate, s_hoje;
        loja = substitui_caracteres($("#c_loja").val(),"\n","_")   ;


        if (trim($("#c_dt_cad_ocorrencia_inicio").val()) == "") {
            alert("Informe a data de início do período!!");
            $("#c_dt_cad_ocorrencia_inicio").focus();
            return;
        }

        if (trim($("#c_dt_cad_ocorrencia_termino").val()) == "") {
            alert("Informe a data de término do período!!");
            $("#c_dt_cad_ocorrencia_termino").focus();
            return;
        }


        s_de = trim($("#c_dt_cad_ocorrencia_inicio").val());
        s_ate = trim($("#c_dt_cad_ocorrencia_termino").val());
        if ((s_de != "") && (s_ate != "")) {
            s_de = retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(s_de));
            s_ate = retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(s_ate));
            if (s_de > s_ate) {
                alert("Data de término é menor que a data de início!!");
                $("#c_dt_cad_ocorrencia_termino").focus();
                return;
            }
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

        strUrl = 'http://<%=Request.ServerVariables("SERVER_NAME")%>:<%=Request.ServerVariables("SERVER_PORT")%>' + serverVariableUrl + 'WebAPI/api/Relatorios/GetXLSReport/';
        strUrl = strUrl + '?usuario=<%=usuario%>';
        strUrl = strUrl + '&dt_inicio=' + $("#c_dt_cad_ocorrencia_inicio").val();
        strUrl = strUrl + '&dt_termino=' + $("#c_dt_cad_ocorrencia_termino").val();
        strUrl = strUrl + '&motivo_ocorrencia=' + $("#c_motivo_abertura").val();
        strUrl = strUrl + '&tp_ocorrencia=' + $("#c_tipo_ocorrencia").val();
        strUrl = strUrl + '&transportadora=' + $("#c_transportadora").val();
        strUrl = strUrl + '&vendedor=' + $("#c_vendedor").val();
        strUrl = strUrl + '&indicador=' + $("#c_indicador").val();
        strUrl = strUrl + '&uf=' + $("#c_uf").val();
        strUrl = strUrl + '&loja=' + loja;
        

        xmlhttp.onreadystatechange = function () {
            var xmlResp;

            if (xmlhttp.readyState == AJAX_REQUEST_IS_COMPLETE) {
                xmlResp = JSON.parse(xmlhttp.responseText);

                if (xmlResp.Status == "OK") {

                	gerarRelatorio.action = 'http://<%=Request.ServerVariables("SERVER_NAME")%>:<%=Request.ServerVariables("SERVER_PORT")%>' + serverVariableUrl + 'WebAPI/api/Relatorios/downloadXLS/?fileName=' + xmlResp.fileName;
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

<body onload="fFILTRO.c_dt_cad_ocorrencia_inicio.focus();">
<center>
<!-- MENSAGEM: "Aguarde, obtendo dados" -->

	<div id="divMsgAguardeObtendoDados" name="divMsgAguardeObtendoDados" style="background-image: url('../Imagem/ajax_loader_gray_256.gif');background-repeat:no-repeat;background-position: center center;position:absolute;bottom:0px;left:0px;width:100%;height:100%;z-index:9;border: 1pt solid #C0C0C0;background-color: black;opacity:0.6;visibility:hidden;vertical-align: middle">

	</div>

    <form name="gerarRelatorio" id="gerarRelatorio" method="POST">
    <input type="hidden" name="idRel" id="idRel" value="" />
    </form>
<form id="fFILTRO" name="fFILTRO" method="post" action="RelPedidoOcorrenciaEstatisticasExec.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Estatísticas de Ocorrências</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>
<br>

<table class="Qx" cellSpacing="0" style="width:240px;">
<!--  PERÍODO: DATA DA OCORRÊNCIA  -->
	<tr>
		<td class="ME MD MC" NOWRAP><span class="PLTe">PERÍODO DE CADASTRAMENTO DA OCORRÊNCIA</span></td>
	</tr>
	<tr bgColor="#FFFFFF" NOWRAP>
		<td class="ME MD MB">
			<table cellSpacing="0" cellPadding="0" style="margin:0px 20px 6px 10px;">
			<tr bgColor="#FFFFFF">
				<td>
					<input class="Cc" maxlength="10" style="width:70px;" name="c_dt_cad_ocorrencia_inicio" id="c_dt_cad_ocorrencia_inicio" onblur="if (!isDate(this)) {alert('Data de início inválida!'); this.focus();}" onkeypress="if (digitou_enter(true)) fFILTRO.c_dt_cad_ocorrencia_termino.focus(); filtra_data();"
						>&nbsp;
					<span class="C">&nbsp;até&nbsp;</span>&nbsp;
					<input class="Cc" maxlength="10" style="width:70px;" name="c_dt_cad_ocorrencia_termino" id="c_dt_cad_ocorrencia_termino" onblur="if (!isDate(this)) {alert('Data de término inválida!'); this.focus();}" onkeypress="if (digitou_enter(true)) bCONFIRMA.focus(); filtra_data();">
				</td>
			</tr>
			</table>
		</td>
	</tr>

<!--  MOTIVO ABERTURA  -->
	<tr>
		<td class="ME MD PLTe" NOWRAP valign="bottom">&nbsp;MOTIVO DA ABERTURA DA OCORRÊNCIA</td>
	</tr>
	<tr bgColor="#FFFFFF" NOWRAP>
		<td class="ME MB MD">
			<select id="c_motivo_abertura" name="c_motivo_abertura" style="width:450px;margin:1px 10px 6px 10px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
			<%=codigo_descricao_monta_itens_select(GRUPO_T_CODIGO_DESCRICAO__OCORRENCIAS_EM_PEDIDOS__MOTIVO_ABERTURA, "")%>
			</select>
		</td>
	</tr>

<!--  TIPO DE OCORRÊNCIA  -->
	<tr>
		<td class="ME MD PLTe" NOWRAP valign="bottom">&nbsp;TIPO DE OCORRÊNCIA</td>
	</tr>
	<tr bgColor="#FFFFFF" NOWRAP>
		<td class="ME MB MD">
			<select id="c_tipo_ocorrencia" name="c_tipo_ocorrencia" style="width:450px;margin:1px 10px 6px 10px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
			<%=codigo_descricao_monta_itens_select(GRUPO_T_CODIGO_DESCRICAO__OCORRENCIAS_EM_PEDIDOS__TIPO_OCORRENCIA, "")%>
			</select>
		</td>
	</tr>

<!--  TRANSPORTADORA  -->
	<tr>
		<td class="ME MD PLTe" NOWRAP valign="bottom">&nbsp;TRANSPORTADORA</td>
	</tr>
	<tr bgColor="#FFFFFF" NOWRAP>
		<td class="ME MB MD">
			<select id="c_transportadora" name="c_transportadora" style="margin:1px 10px 6px 10px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
			<% =transportadora_monta_itens_select(Null) %>
			</select>
		</td>
	</tr>

<!--  VENDEDOR  -->
	<tr>
		<td class="ME MD PLTe" NOWRAP valign="bottom">&nbsp;VENDEDOR</td>
	</tr>
	<tr bgColor="#FFFFFF" NOWRAP>
		<td class="ME MB MD">
			<input maxlength="10" class="PLLe" style="width:220px;" name="c_vendedor" id="c_vendedor" onblur="this.value=trim(this.value);" onkeypress="if (digitou_enter(true)) fFILTRO.c_indicador.focus(); filtra_nome_identificador();">
		</td>
	</tr>

<!--  INDICADOR  -->
	<tr>
		<td class="ME MD PLTe" NOWRAP valign="bottom">&nbsp;INDICADOR</td>
	</tr>
	<tr bgColor="#FFFFFF" NOWRAP>
		<td class="ME MB MD">
			<select id="c_indicador" name="c_indicador" style="margin:1px 10px 6px 10px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
			<% =indicadores_monta_itens_select(Null) %>
			</select>
		</td>
	</tr>

<!--  UF  -->
	<tr>
		<td class="ME MD PLTe" NOWRAP valign="bottom">&nbsp;UF</td>
	</tr>
	<tr bgColor="#FFFFFF" NOWRAP>
		<td class="ME MB MD">
			<select id="c_uf" name="c_uf" style="margin:1px 10px 6px 10px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
			<% =uf_monta_itens_select(Null) %>
			</select>
		</td>
	</tr>
	
<!--  LOJA(S)  -->
	<tr>
		<td class="ME MD PLTe" NOWRAP valign="bottom">&nbsp;LOJA(S)</td>
	</tr>
	<tr bgColor="#FFFFFF" NOWRAP>
		<td class="ME MB MD">
			<table cellSpacing="0" cellPadding="0" style="margin:0px 20px 6px 10px;">
			<tr bgColor="#FFFFFF">
				<td>
					<textarea class="PLBe" style="width:100px;font-size:9pt;margin-bottom:4px;" rows="4" name="c_loja" id="c_loja" onkeypress="if (!digitou_enter(false) && !digitou_char('-')) filtra_numerico();" onblur="this.value=normaliza_lista_lojas(this.value);"></textarea>
				</td>
			</tr>
			</table>
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
'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>
