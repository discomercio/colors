<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ===============================================
'	  RelDevolucao2Filtro.asp
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

	dim usuario
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	CONECTA COM O BANCO DE DADOS
	dim cn
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	if Not operacao_permitida(OP_CEN_REL_DEVOLUCAO_PRODUTOS2, s_lista_operacoes_permitidas) then 
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if
	
	dim dtMinDtInicialFiltroPeriodo, intMaxDiasDtInicialFiltroPeriodo
	dim strMinDtInicialFiltroPeriodoYYYYMMDD, strMinDtInicialFiltroPeriodoDDMMYYYY
	if operacao_permitida(OP_CEN_RESTRINGE_DT_INICIAL_FILTRO_PERIODO, s_lista_operacoes_permitidas) then
		intMaxDiasDtInicialFiltroPeriodo = obtem_max_dias_dt_inicial_filtro_periodo()
		dtMinDtInicialFiltroPeriodo = Date - intMaxDiasDtInicialFiltroPeriodo
		strMinDtInicialFiltroPeriodoYYYYMMDD = formata_data_yyyymmdd(dtMinDtInicialFiltroPeriodo)
		strMinDtInicialFiltroPeriodoDDMMYYYY = formata_data(dtMinDtInicialFiltroPeriodo)
	else
		strMinDtInicialFiltroPeriodoYYYYMMDD = ""
		strMinDtInicialFiltroPeriodoDDMMYYYY = ""
		end if





' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________

' ____________________________________________________________________________
' INDICADORES MONTA ITENS SELECT
' LEMBRE-SE: O ORÇAMENTISTA É CONSIDERADO AUTOMATICAMENTE UM INDICADOR!!
function indicadores_monta_itens_select(byval id_default)
dim x, r, strResp, ha_default
	id_default = Trim("" & id_default)
	ha_default=False
	set r = cn.Execute("SELECT apelido, razao_social_nome_iniciais_em_maiusculas FROM t_ORCAMENTISTA_E_INDICADOR ORDER BY apelido")
	strResp = ""
	do while Not r.eof 
		x = UCase(Trim("" & r("apelido")))
		if (id_default<>"") And (id_default=x) then
			strResp = strResp & "<OPTION SELECTED"
			ha_default=True
		else
			strResp = strResp & "<OPTION"
			end if
		strResp = strResp & " VALUE='" & x & "'>"
		strResp = strResp & x & " - " & Trim("" & r("razao_social_nome_iniciais_em_maiusculas"))
		strResp = strResp & "</OPTION>" & chr(13)
		r.MoveNext
		loop

	if Not ha_default then
		strResp = "<OPTION SELECTED VALUE=''>&nbsp;</OPTION>" & chr(13) & strResp
		end if
		
	indicadores_monta_itens_select = strResp
	r.close
	set r=nothing
end function


' ____________________________________________________________________________
' CAPTADORES MONTA ITENS SELECT
function captadores_monta_itens_select(byval id_default)
dim x, r, strSql, strResp, ha_default
	id_default = Trim("" & id_default)
	ha_default=False
	strSql = "SELECT" & _
				" usuario," & _
				" nome_iniciais_em_maiusculas" & _
			" FROM t_USUARIO" & _
			" WHERE" & _
				" usuario IN " & _
					"(" & _
						"SELECT DISTINCT" & _
							" captador" & _
						" FROM t_ORCAMENTISTA_E_INDICADOR" & _
						" WHERE" & _
							" (captador IS NOT NULL)" & _
					")"
	set r = cn.Execute(strSql)
	strResp = ""
	do while Not r.eof 
		x = UCase(Trim("" & r("usuario")))
		if (id_default<>"") And (id_default=x) then
			strResp = strResp & "<OPTION SELECTED"
			ha_default=True
		else
			strResp = strResp & "<OPTION"
			end if
		strResp = strResp & " VALUE='" & x & "'>"
		strResp = strResp & x & " - " & Trim("" & r("nome_iniciais_em_maiusculas"))
		strResp = strResp & "</OPTION>" & chr(13)
		r.MoveNext
		loop

	if Not ha_default then
		strResp = "<OPTION SELECTED VALUE=''>&nbsp;</OPTION>" & chr(13) & strResp
		end if
		
	captadores_monta_itens_select = strResp
	r.close
	set r=nothing
end function

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



<% if False then 'APENAS P/ HABILITAR O INTELLISENSE DURANTE O DESENVOLVIMENTO!! %>
<script src="../Global/jquery.js" language="JavaScript" type="text/javascript"></script>
<% end if %>

<script src="<%=URL_FILE__JQUERY%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__AJAX_JS%>" language="JavaScript" type="text/javascript"></script>

<script type="text/javascript">
    $(function () {
        $("#divMsgAguardeObtendoDados").css('filter', 'alpha(opacity=50)');
    });
</script>

<script language="JavaScript" type="text/javascript">
function fFILTROConfirma(f) {
var strDtRefYYYYMMDD, strDtRefDDMMYYYY;

	if (!f.ckb_periodo_devolucao.checked) {
		alert("É obrigatório informar um período para a consulta!!");
		f.c_dt_devolucao_inicio.focus();
		return;
	}

	if (f.ckb_periodo_devolucao.checked) {
		if (trim(f.c_dt_devolucao_inicio.value) == "") {
			alert("Preencha a data de início!!");
			f.c_dt_devolucao_inicio.focus();
			return;
		}
		if (trim(f.c_dt_devolucao_termino.value) == "") {
			alert("Preencha a data de término!!");
			f.c_dt_devolucao_termino.focus();
			return;
		}
	
		if (!consiste_periodo(f.c_dt_devolucao_inicio, f.c_dt_devolucao_termino)) return;
		}

	if ((trim(f.c_produto.value)!="")&&(trim(f.c_fabricante.value)=="")) {
		if (!isEAN(f.c_produto.value)) {
			alert("Preencha o código do fabricante do produto " + f.c_produto.value + "!!");
			f.c_fabricante.focus();
			return;
			}
		}

//  Período de consulta está restrito por perfil de acesso?
	if (trim(f.c_MinDtInicialFiltroPeriodoYYYYMMDD.value) != "") {
		strDtRefDDMMYYYY = trim(f.c_dt_devolucao_inicio.value);
		if (trim(strDtRefDDMMYYYY) != "") {
			strDtRefYYYYMMDD = retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(strDtRefDDMMYYYY));
			if (strDtRefYYYYMMDD < f.c_MinDtInicialFiltroPeriodoYYYYMMDD.value) {
				alert("Data inválida para consulta: " + strDtRefDDMMYYYY + "\nO período de consulta não pode compreender datas anteriores a " + f.c_MinDtInicialFiltroPeriodoDDMMYYYY.value + "!!");
				return;
				}
			}
		strDtRefDDMMYYYY = trim(f.c_dt_devolucao_termino.value);
		if (trim(strDtRefDDMMYYYY) != "") {
			strDtRefYYYYMMDD = retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(strDtRefDDMMYYYY));
			if (strDtRefYYYYMMDD < f.c_MinDtInicialFiltroPeriodoYYYYMMDD.value) {
				alert("Data inválida para consulta: " + strDtRefDDMMYYYY + "\nO período de consulta não pode compreender datas anteriores a " + f.c_MinDtInicialFiltroPeriodoDDMMYYYY.value + "!!");
				return;
				}
			}
		}

	if (f.rb_saida[1].checked) {
	    gera_relatorio_excel(f);
	}
	else {
	    dCONFIRMA.style.visibility = "hidden";
	    window.status = "Aguarde ...";
	    f.action = "RelDevolucao2Exec.asp";
	    f.submit();
	}
}

function gera_relatorio_excel(f) {
var serverVariableUrl, strUrl, xmlHttp, s_loja;
serverVariableUrl = '<%=Request.ServerVariables("URL")%>';
serverVariableUrl = serverVariableUrl.toUpperCase();
serverVariableUrl = serverVariableUrl.substring(0, serverVariableUrl.indexOf("CENTRAL"));

s_loja = f.c_lista_loja.value;
//s_loja = s_loja.replace(String.fromCharCode(10), "");
s_loja = s_loja.replace(/\n/g, "_");

xmlhttp = GetXmlHttpObject();
if (xmlhttp == null) {
    alert("O browser NÃO possui suporte ao AJAX!!");
    return;
}

window.status = "Aguarde, gerando arquivo ...";
divMsgAguardeObtendoDados.style.visibility = "";

strUrl = '<%=getProtocoloEmUsoHttpOrHttps%>://<%=Request.ServerVariables("SERVER_NAME")%>:<%=Request.ServerVariables("SERVER_PORT")%>' + serverVariableUrl + 'WebAPI/api/Relatorios/GeraDevolucaoProdutos2XLS/';
strUrl = strUrl + '?usuario=<%=usuario%>';
strUrl = strUrl + '&dt_devolucao_inicio=' + f.c_dt_devolucao_inicio.value;
strUrl = strUrl + '&dt_devolucao_termino=' + f.c_dt_devolucao_termino.value;
strUrl = strUrl + '&fabricante=' + f.c_fabricante.value;
strUrl = strUrl + '&produto=' + f.c_produto.value;
strUrl = strUrl + '&pedido=' + f.c_pedido.value;
strUrl = strUrl + '&vendedor=' + f.c_vendedor.value;
strUrl = strUrl + '&indicador=' + f.c_indicador.value;
strUrl = strUrl + '&captador=' + f.c_captador.value;
strUrl = strUrl + '&lojas=' + s_loja;

xmlhttp.onreadystatechange = function () {
var xmlResp;

    if (xmlhttp.readyState == AJAX_REQUEST_IS_COMPLETE) {
        xmlResp = JSON.parse(xmlhttp.responseText);

        if (xmlResp.Status == "OK") {

			fFILTRO.action = '<%=getProtocoloEmUsoHttpOrHttps%>://<%=Request.ServerVariables("SERVER_NAME")%>:<%=Request.ServerVariables("SERVER_PORT")%>' + serverVariableUrl + 'WebAPI/api/Relatorios/downloadDevolucaoProdutos2XLS/?fileName=' + xmlResp.fileName;
            fFILTRO.submit();

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
<style type="text/css">
#rb_saida {
margin: 0pt 2pt 0pt 15pt;
vertical-align: top;
}
</style>

<body>
<center>

<form id="fFILTRO" name="fFILTRO" method="post" action="RelDevolucao2Exec.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="c_MinDtInicialFiltroPeriodoYYYYMMDD" id="c_MinDtInicialFiltroPeriodoYYYYMMDD" value='<%=strMinDtInicialFiltroPeriodoYYYYMMDD%>'>
<input type="hidden" name="c_MinDtInicialFiltroPeriodoDDMMYYYY" id="c_MinDtInicialFiltroPeriodoDDMMYYYY" value='<%=strMinDtInicialFiltroPeriodoDDMMYYYY%>'>

	<div id="divMsgAguardeObtendoDados" name="divMsgAguardeObtendoDados" style="background-image: url('../Imagem/ajax_loader_gray_256.gif');background-repeat:no-repeat;background-position: center center;position:absolute;bottom:0px;left:0px;width:100%;height:100%;z-index:9;border: 1pt solid #C0C0C0;background-color: black;opacity:0.6;visibility:hidden;vertical-align: middle">

	</div>

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Devolução de Produtos II</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>
<br>

<table class="Qx" style="width:240px;" cellSpacing="0">
<!--  PERÍODO DE DEVOLUÇÃO  -->
	<tr bgColor="#FFFFFF">
	<td class="MT" NOWRAP><span class="PLTe">PERÍODO DE DEVOLUÇÃO</span>
		<br>
		<table cellSpacing="0" cellPadding="0" style="margin-bottom:10px;">
		<tr bgColor="#FFFFFF"><td>
			<input type="checkbox" tabindex="-1" id="ckb_periodo_devolucao" name="ckb_periodo_devolucao"
				value="PERIODO_ON" onclick="if (fFILTRO.ckb_periodo_devolucao.checked) fFILTRO.c_dt_devolucao_inicio.focus();"><span class="C" style="cursor:default" 
				onclick="fFILTRO.ckb_periodo_devolucao.click();">Devolvido entre</span
				><input class="Cc" maxlength="10" style="width:70px;" name="c_dt_devolucao_inicio" id="c_dt_devolucao_inicio" onblur="if (!isDate(this)) {alert('Data inválida!'); this.focus();}" onkeypress="if (digitou_enter(true)) fFILTRO.c_dt_devolucao_termino.focus(); else fFILTRO.ckb_periodo_devolucao.checked=true; filtra_data();" onclick="fFILTRO.ckb_periodo_devolucao.checked=true;"
				>&nbsp;<span class="C">e</span>&nbsp;<input class="Cc" maxlength="10" style="width:70px;" name="c_dt_devolucao_termino" id="c_dt_devolucao_termino" onblur="if (!isDate(this)) {alert('Data inválida!'); this.focus();}" onkeypress="if (digitou_enter(true)) fFILTRO.c_fabricante.focus(); else fFILTRO.ckb_periodo_devolucao.checked=true; filtra_data();" onclick="fFILTRO.ckb_periodo_devolucao.checked=true;">
			</td></tr>
		</table>
	</td></tr>

<!--  FABRICANTE  -->
	<tr bgColor="#FFFFFF">
	<td class="MDBE" NOWRAP><span class="PLTe">FABRICANTE</span>
	<br>
		<input maxlength="4" class="PLLe" style="width:50px;" name="c_fabricante" id="c_fabricante" onblur="this.value=normaliza_codigo(this.value,TAM_MIN_FABRICANTE);" onkeypress="if (digitou_enter(true)) fFILTRO.c_produto.focus(); filtra_fabricante();">
		</td></tr>

<!--  PRODUTO  -->
	<tr bgColor="#FFFFFF">
	<td class="MDBE" NOWRAP><span class="PLTe">PRODUTO</span>
	<br>
		<input maxlength="13" class="PLLe" style="width:100px;" name="c_produto" id="c_produto" onblur="this.value=ucase(normaliza_codigo(this.value,TAM_MIN_PRODUTO));" onkeypress="if (digitou_enter(true)) fFILTRO.c_pedido.focus(); filtra_produto();">
		</td></tr>

<!--  PEDIDO  -->
	<tr bgColor="#FFFFFF">
	<td class="MDBE" NOWRAP><span class="PLTe">PEDIDO</span>
	<br>
		<input maxlength="10" class="PLLe" style="width:100px;" name="c_pedido" id="c_pedido" onblur="if (normaliza_num_pedido(this.value)!='') this.value=normaliza_num_pedido(this.value);" onkeypress="if (digitou_enter(true)) fFILTRO.c_vendedor.focus(); filtra_pedido();">
		</td></tr>
		
<!--  VENDEDOR  -->
	<tr bgColor="#FFFFFF">
	<td class="MDBE" NOWRAP><span class="PLTe">VENDEDOR</span>
	<br>
		<input maxlength="10" class="PLLe" style="width:100px;" name="c_vendedor" id="c_vendedor" onblur="this.value=trim(this.value);" onkeypress="if (digitou_enter(true)) fFILTRO.c_lista_loja.focus(); filtra_nome_identificador();">
		</td></tr>

<!--  INDICADOR  -->
	<tr bgColor="#FFFFFF">
		<td class="MDBE" NOWRAP><span class="PLTe">INDICADOR</span>
		<br>
			<select id="c_indicador" name="c_indicador" style="margin:1px 10px 6px 10px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
			<% =indicadores_monta_itens_select(Null) %>
			</select>
			</td></tr>

<!--  CAPTADOR  -->
	<tr bgColor="#FFFFFF">
		<td class="MDBE" NOWRAP><span class="PLTe">CAPTADOR</span>
		<br>
			<select id="c_captador" name="c_captador" style="margin:1px 10px 6px 10px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
			<% =captadores_monta_itens_select(Null) %>
			</select>
			</td></tr>
			
<!-- ************   LOJAS   ************ -->
	<tr bgColor="#FFFFFF">
	<td class="MDBE" NOWRAP><span class="PLTe">LOJA(S)</span>
		<br>
		<textarea class="PLBe" style="font-size:9pt;width:110px;margin-bottom:4px;margin-left:10px;" rows="6" name="c_lista_loja" id="c_lista_loja" onkeypress="if (!digitou_enter(false) && !digitou_char('-')) filtra_numerico();" onblur="this.value=normaliza_lista_lojas(this.value);"></textarea>
	</td></tr>
<!--  SAÍDA DO RELATÓRIO  -->
	<tr bgcolor="#FFFFFF">
	<td colspan="2" class="MDBE" nowrap><span class="PLTe">SAÍDA</span>
		<br><input type="radio" class="rbOpt" tabindex="-1" id="rb_saida" name="rb_saida" value="Html" onclick="dCONFIRMA.style.visibility='';" checked><span class="C lblOpt" style="cursor:default" onclick="fFILTRO.rb_saida[0].click();dCONFIRMA.style.visibility='';"
			>Html</span>

		<br><input type="radio" class="rbOpt" tabindex="-1" id="rb_saida" name="rb_saida" value="XLS" onclick="dCONFIRMA.style.visibility='';"><span class="C lblOpt" style="cursor:default" onclick="fFILTRO.rb_saida[1].click();dCONFIRMA.style.visibility='';"
			>Excel</span>
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
