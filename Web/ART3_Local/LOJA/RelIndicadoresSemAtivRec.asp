<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/Global.asp"    -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =========================================================
'	  I N D I C A D O R E S S E M A T I V R E C . A S P
'     =========================================================
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

	dim usuario,loja
	usuario = Trim(Session("usuario_atual"))
    loja = Session("loja_atual")
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

	dim s_script, strSql
	
'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

    if Not operacao_permitida(OP_LJA_REL_INDICADORES_SEM_ATIVIDADE_RECENTE, s_lista_operacoes_permitidas) then
        Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
    end if

'	FILTROS	
	dim c_vendedor, c_indicador
	dim c_loja
	


	c_vendedor = Trim(Request.Form("c_vendedor"))
	c_indicador = Trim(Request.Form("c_indicador"))
    


	c_loja = Trim(Request.Form("c_loja"))



' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________

' ____________________________________________________________________________
' VENDEDORES MONTA ITENS SELECT
'
function vendedores_monta_itens_select(byval id_default)
dim x, r, strResp, ha_default, strSql
	id_default = Trim("" & id_default)
	ha_default=False
	strSql = "SELECT DISTINCT" & _
				" usuario, nome_iniciais_em_maiusculas" & _
			" FROM t_USUARIO" & _
			" WHERE" & _
				" (vendedor_loja <> 0)" & _
                " AND (bloqueado = 0)" & _
				" AND (" & _
					"usuario IN (" & _
						"SELECT DISTINCT" & _
							" usuario" & _
						" FROM t_USUARIO_X_LOJA" & _
						" WHERE" & _
							" (loja = '" & loja & "')" & _
						")" & _
					")" & _
			" ORDER BY" & _
				" usuario"
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

	vendedores_monta_itens_select = strResp
	r.close
	set r=nothing
end function

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
	<title>LOJA</title>
    <meta charset="utf-8" />
	</head>

<script src="<%=URL_FILE__JQUERY%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_I18N%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__AJAX_JS%>" language="JavaScript" type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
var s_ult_vendedor_selecionado = "--XX--XX--XX--XX--XX--";

function fFILTROConfirma( f ) {
var strDtRefYYYYMMDD, strDtRefDDMMYYYY;
var data;
     
	dCONFIRMA.style.visibility = "hidden";

	$("#c_vendedor").children().prop('selected', true);
		
	window.status = "Aguarde ...";
	f.submit();
}

function ind_new(vendedor, apelido, nome) {
	this.vendedor = vendedor;
	this.apelido = apelido;
	this.nome = nome;
	return this;
}

</script>

<script type="text/javascript">
    $(function () {
        var data, ano, i, opt;

        $("#divMsgAguardeObtendoDados").css('filter', 'alpha(opacity=50)');

        $("#btnAdiciona").click(function () {
            var x = $("#c_vendedor_escolher option:selected");
            $("#c_vendedor").append(x);
            reOrdenarEscolhidos();
        });

        $("#btnRemove").click(function () {
            var x = $("#c_vendedor option:selected");
            $("#c_vendedor_escolher").append(x);
            reOrdenarAEscolher();
        });

        $("#c_vendedor_escolher").dblclick(function () {
            var x = $("#c_vendedor_escolher option:selected");
            $("#c_vendedor").append(x);
            reOrdenarEscolhidos();
        });

        $("#c_vendedor").dblclick(function () {
            var x = $("#c_vendedor option:selected");
            $("#c_vendedor_escolher").append(x);
            reOrdenarAEscolher();
        });
                    
    });

    function reOrdenarAEscolher() {
        $("#c_vendedor_escolher").html($("#c_vendedor_escolher option").sort(function (a, b) {
            return a.text.toUpperCase() == b.text.toUpperCase() ? 0 : a.text.toUpperCase() < b.text.toUpperCase() ? -1 : 1
        }))
    }

    function reOrdenarEscolhidos() {
        $("#c_vendedorr").html($("#c_vendedor option").sort(function (a, b) {
            return a.text.toUpperCase() == b.text.toUpperCase() ? 0 : a.text.toUpperCase() < b.text.toUpperCase() ? -1 : 1
        }))
    }

</script>
<script type="text/javascript">
    function CarregaListaVendedores(a, m) {
        var strUrl, xmlhttp;
        xmlhttp = GetXmlHttpObject();
        if (xmlhttp == null) {
            alert("O browser NÃO possui suporte ao AJAX!!");
            return;
        }

        window.status = "Aguarde, pesquisando vendedores de  " + m + "/" + a + " ...";
        divMsgAguardeObtendoDados.style.visibility = "";

        strUrl = "../Global/AjaxRelComissaoIndicadoresListaVendedores.asp";
        strUrl = strUrl + "?ano=" + a;
        strUrl = strUrl + "&mes=" + m;
        strUrl = strUrl + "&id=" + Math.random();
        xmlhttp.onreadystatechange = function () {
            var strResp;

            if (xmlhttp.readyState == 4) {
                strResp = xmlhttp.responseText;
                if (strResp == "") {
                    $('#spn_aviso').css('display', 'block');
                    $("#c_vendedor_escolher").children().empty();
                    divMsgAguardeObtendoDados.style.visibility = "hidden";
                }
                if (strResp != "") {
                    try {
                        $('#c_vendedor_escolher').html(xmlhttp.responseText);
                        $('#spn_aviso').css('display', 'none');
                        window.status = "Concluído"
                        divMsgAguardeObtendoDados.style.visibility = "hidden";
                    }
                    catch (e) {
                        alert("Falha na consulta!!");
                    }
                }
            }
        }
        xmlhttp.open("GET", strUrl, true);
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

<link href="<%=URL_FILE__E_CSS%>" rel="stylesheet" type="text/css">
<link href="<%=URL_FILE__JQUERY_UI_CSS%>" rel="stylesheet" type="text/css">

<style type="text/css">
 
 .aviso {
    font-family: Arial, Helvetica, sans-serif;
	font-size: 8pt;
	font-weight: bold;
	font-style: normal;
	margin: 0pt 0pt 0pt 0pt;
	color: #f00;
    display: none;
 }
 
</style>
<style type="text/css">
 
 #spn_aviso {
    display: none;
 }

</style>

<body onload="if (trim(fFILTRO.c_dt_entregue_mes.value)!='' && trim(fFILTRO.c_dt_entregue_ano.value)!='') { CarregaListaVendedores(fFILTRO.c_dt_entregue_ano.value, fFILTRO.c_dt_entregue_mes.value);}">
<center>
<!-- MENSAGEM: "Aguarde, obtendo dados" -->

	<div id="divMsgAguardeObtendoDados" name="divMsgAguardeObtendoDados" style="background-image: url('../Imagem/ajax_loader_gray_256.gif');background-repeat:no-repeat;background-position: center center;position:absolute;bottom:0px;left:0px;width:100%;height:100%;z-index:9;border: 1pt solid #C0C0C0;background-color: black;opacity:0.6;visibility:hidden;vertical-align: middle">

	</div>
<form id="fFILTRO" name="fFILTRO" method="post" action="RelIndicadoresSemAtivRecExec.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="690" cellpadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Relatório de Indicadores sem Atividade Recente</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>
<br>

<!--  PARÂMETROS  -->
<table width="690" class="Qx" cellspacing="0">


<!--  CADASTRAMENTO  -->
<tr bgcolor="#FFFFFF">
<td class="MT" align="left" nowrap><span class="PLTe">Vendedores</span>
	<br>
	<table cellspacing="3" cellpadding="0" style="margin-bottom:10px; width: 100%">
	<tr bgcolor="#FFFFFF">
		<td align="left" style="width:47%;"><span class="C" style="margin-left:0px;">Selecione o(s) vendedor(es)</span></td>
        <td style="width:6%;">&nbsp;</td>
        <td align="left" style="width:47%"><span class="C" style="margin-left:0px;">Vendedor(es) selecionado(s)</span></td>
    </tr>
    <tr>
		<td align="left">
			<select id="c_vendedor_escolher" name="c_vendedor_escolher" style="width:95%" size="10" multiple>
			<%=vendedores_monta_itens_select(c_vendedor) %>>
			</select>
            <br />
            <span class="C" id="spn_aviso" style="color:red;width:100%">Nenhum vendedor a ser processado.</span><span class="C">&nbsp;&nbsp;</span>
		</td>
        <td>
            <input type="button" id="btnAdiciona" value="&raquo;" />
            <br />
            <input type="button" id="btnRemove" value="&laquo;" />
        </td>
        <td align="left">
			<select id="c_vendedor" name="c_vendedor" style="width:95%" size="10" multiple>
			</select>
		</td>
	</tr>
	</table>
</td>
</tr>

</table>

<!-- ************   SEPARADOR   ************ -->
<table width="690" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc" align="left">&nbsp;</td></tr>
</table>
<br>


<table width="690" cellspacing="0">
<tr>
	<td align="left"><a name="bVOLTAR" id="bVOLTAR" href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="volta para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
	<td align="right"><div name="dCONFIRMA" id="dCONFIRMA"><a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fFILTROConfirma(fFILTRO)" title="executa a consulta">
		<img src="../botao/proximo.gif" width="176" height="55" border="0"></a></div>
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