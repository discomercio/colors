<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/Global.asp"    -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ========================================================
'	  R E L C O M I S S A O I N D I C A D O R E S . A S P
'     ========================================================
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
'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	if Not operacao_permitida(OP_LJA_REL_PEDIDOS_CANCELADOS , s_lista_operacoes_permitidas) then 
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if
	

'	FILTROS
	dim ckb_st_entrega_cancelado, c_dt_cancel_inicio, c_dt_cancel_termino
	dim c_vendedor
	dim c_loja

	c_dt_cancel_inicio = Trim(Request.Form("c_dt_cancel_inicio"))
	c_dt_cancel_termino = Trim(Request.Form("c_dt_cancel_termino"))
	c_vendedor = Trim(Request.Form("c_vendedor"))
	c_loja = Trim(Request.Form("c_loja"))

	
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
	<title>LOJA</title>
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

    function fFILTROConfirma(f) {
        var strDtRefYYYYMMDD, strDtRefDDMMYYYY;
        var b;

        if (f.ckb_st_entrega_cancelado.checked) {
            if (!consiste_periodo(f.c_dt_cancel_inicio, f.c_dt_cancel_termino)) return;
        }
        //  PERÍODO DE CADASTRO
        if ((f.c_dt_cancel_inicio.value != "") && (f.c_dt_cancel_termino.value != "")) {
            if (f.ckb_st_entrega_cancelado.checked) {
                if (trim(f.c_dt_cancel_inicio.value) != "") {
                    if (!isDate(f.c_dt_cancel_inicio)) {
                        alert("Data inválida!!");
                        f.c_dt_cadastro_inicio.focus();
                        return;
                    }
                }

                if (trim(f.c_dt_cancel_termino.value) != "") {
                    if (!isDate(f.c_dt_cancel_termino)) {
                        alert("Data inválida!!");
                        f.c_dt_cadastro_termino.focus();
                        return;
                    }
                }

                s_de = trim(f.c_dt_cancel_inicio.value);
                s_ate = trim(f.c_dt_cancel_termino.value);
                if ((s_de != "") && (s_ate != "")) {
                    s_de = retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(s_de));
                    s_ate = retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(s_ate));
                    if (s_de > s_ate) {
                        alert("Data de término é menor que a data de início!!");
                        f.c_dt_cadastro_termino.focus();
                        return;
                    }
                }
            }
        }
        else {
            alert("Selecione o período do cancelamento!");
            return;
        }
        //  Período de consulta está restrito por perfil de acesso?
        if (trim(f.c_MinDtInicialFiltroPeriodoYYYYMMDD.value) != "") {
            strDtRefDDMMYYYY = trim(f.c_dt_cancel_inicio.value);
            if (trim(strDtRefDDMMYYYY) != "") {
                strDtRefYYYYMMDD = retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(strDtRefDDMMYYYY));
                if (strDtRefYYYYMMDD < f.c_MinDtInicialFiltroPeriodoYYYYMMDD.value) {
                    alert("Data inválida para consulta: " + strDtRefDDMMYYYY + "\nO período de consulta não pode compreender datas anteriores a " + f.c_MinDtInicialFiltroPeriodoDDMMYYYY.value + "!!");
                    return;
                }
            }
            strDtRefDDMMYYYY = trim(f.c_dt_cancel_termino.value);
            if (trim(strDtRefDDMMYYYY) != "") {
                strDtRefYYYYMMDD = retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(strDtRefDDMMYYYY));
                if (strDtRefYYYYMMDD < f.c_MinDtInicialFiltroPeriodoYYYYMMDD.value) {
                    alert("Data inválida para consulta: " + strDtRefDDMMYYYY + "\nO período de consulta não pode compreender datas anteriores a " + f.c_MinDtInicialFiltroPeriodoDDMMYYYY.value + "!!");
                    return;
                }
            }
        }

        dCONFIRMA.style.visibility = "hidden";
        fFILTRO.c_hidden_reload.value = 1
        fFILTRO.c_hidden_sub_motivo.value = $("#c_cod_sub_motivo option:selected").index();
        fFILTRO.c_hidden_motivo.value = $("#c_motivo option:selected").index();            
        window.status = "Aguarde ...";
        f.submit();
    }

</script>

<script type="text/javascript">
    $(function () {
        var z, aux;
        z = 0;
        $("#divMsgAguardeObtendoDados").css('filter', 'alpha(opacity=50)');
        $("#c_dt_cancel_inicio").hUtilUI('datepicker_filtro_inicial');
        $("#c_dt_cancel_termino").hUtilUI('datepicker_filtro_final');
        $(".aviso").css('display', 'none');
        
        if (fFILTRO.c_hidden_motivo.value == "") {
            $("#motivo_alternativo").hide();
        }  
        else {
            CarregaSubMotivo(fFILTRO.c_hidden_motivo.value, '<%=GRUPO_T_CODIGO_DESCRICAO__CANCELAMENTOPEDIDO_MOTIVO%>');
        }
             
        //Every resize of window
        $(window).resize(function () {
            sizeDivAjaxRunning();
        });

        //Every scroll of window
        $(window).scroll(function () {
            sizeDivAjaxRunning();
        });

        //Dynamically assign height
        function sizeDivAjaxRunning() {
            var newTop = $(window).scrollTop() + "px";
            $("#divMsgAguardeObtendoDados").css("top", newTop);
        }      
       
    });

</script>
<script type="text/javascript">
    function CarregaSubMotivo(codigo_pai, grupo_pai) {
        if (codigo_pai != "") {
            var vazio = Option('', ' ');
            var strUrl, xmlhttp;
            xmlhttp = GetXmlHttpObject();
            if (xmlhttp == null) {
                alert("O browser NÃO possui suporte ao AJAX!!");
                return;
            }

            window.status = "Aguarde, pesquisando Sub-Motivo(s)";
            divMsgAguardeObtendoDados.style.visibility = "";

            strUrl = "../Global/AjaxSubMotivoCancelamento.asp";
            strUrl = strUrl + "?grupo_pai=" + grupo_pai;
            strUrl = strUrl + "&codigo_pai=" + codigo_pai;
            strUrl = strUrl + "&id=" + Math.random();
            xmlhttp.onreadystatechange = function () {
                var strResp;

                if (xmlhttp.readyState == 4) {
                    strResp = xmlhttp.responseText;
                    if (strResp == "")  {

                        $("#c_cod_sub_motivo").empty();                      
                        $("#motivo_alternativo").hide();
                        divMsgAguardeObtendoDados.style.visibility = "hidden";
                        window.status = "Concluído"
                    }
                    if (strResp != "") {
                        $("#c_cod_sub_motivo").empty();
                        try {
                            $('#c_cod_sub_motivo').append(vazio);
                            $('#c_cod_sub_motivo').append(xmlhttp.responseText);
                            $("#motivo_alternativo").show();                           
                            $("#c_cod_sub_motivo").prop('selectedIndex', fFILTRO.c_hidden_sub_motivo.value);
                            divMsgAguardeObtendoDados.style.visibility = "hidden";
                            window.status = "Concluído"

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
        else {
            $("#motivo_alternativo").hide();
            $("#c_cod_sub_motivo").empty();
        }
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


<body>
<center>
<!-- MENSAGEM: "Aguarde, obtendo dados" -->

	<div id="divMsgAguardeObtendoDados" name="divMsgAguardeObtendoDados" style="background-image: url('../Imagem/ajax_loader_gray_256.gif');background-repeat:no-repeat;background-position: center center;position:absolute;bottom:0px;left:0px;width:100%;height:100%;z-index:9;border: 1pt solid #C0C0C0;background-color: black;opacity:0.6;visibility:hidden;vertical-align: middle">

	</div>
<form id="fFILTRO" name="fFILTRO" method="post" action="RelPedidoCanceladoExec.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="c_MinDtInicialFiltroPeriodoYYYYMMDD" id="c_MinDtInicialFiltroPeriodoYYYYMMDD" value='<%=strMinDtInicialFiltroPeriodoYYYYMMDD%>'>
<input type="hidden" name="c_MinDtInicialFiltroPeriodoDDMMYYYY" id="c_MinDtInicialFiltroPeriodoDDMMYYYY" value='<%=strMinDtInicialFiltroPeriodoDDMMYYYY%>'>
<input type="hidden" name="c_hidden_sub_motivo" id="c_hidden_sub_motivo" value="" />
<input type="hidden" name="c_hidden_motivo" id="c_hidden_motivo" value="" />
<input type="hidden" name="c_hidden_reload" id="c_hidden_reload" value="0" />

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="690" cellpadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Relatório de Pedidos Cancelados</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>
<br>

<!--  PARÂMETROS  -->
<table width="690" class="Qx" cellspacing="0">
<!--  STATUS DE ENTREGA  -->
<tr bgcolor="#FFFFFF">
<td class="MT" align="left" nowrap><span class="PLTe">PERÍODO</span>
	<br>
	<table cellspacing="0" cellpadding="0" style="margin-bottom:10px;">
	<tr bgcolor="#FFFFFF">
		<td align="left">
		<input type="checkbox" tabindex="-1" id="ckb_st_entrega_cancelado" name="ckb_st_entrega_cancelado" onclick="if (fFILTRO.ckb_st_entrega_cancelado.checked) fFILTRO.c_dt_cancel_inicio.focus();"
			value="<%=ST_ENTREGA_CANCELADO%>"
			<% if (c_dt_cancel_inicio <> "") Or (c_dt_cancel_termino <> "") then Response.Write " checked"%>
			><span class="C" style="cursor:default" 
			onclick="fFILTRO.ckb_st_entrega_cancelado.click();">Pedidos cancelados entre</span
			><input class="Cc" maxlength="10" style="width:70px;" name="c_dt_cancel_inicio" id="c_dt_cancel_inicio" onblur="if (!isDate(this)) {alert('Data inválida!'); this.focus();}" onkeypress="if (digitou_enter(true)) fFILTRO.c_dt_cancel_termino.focus(); else fFILTRO.ckb_st_entrega_cancelado.checked=true; filtra_data();" onclick="fFILTRO.ckb_st_entrega_cancelado.checked = true;" onchange="if (trim(this.value)!='') fFILTRO.ckb_st_entrega_cancelado.checked=true;"
			<% if c_dt_cancel_inicio <> "" then Response.Write " value=" & chr(34) & c_dt_cancel_inicio & chr(34)%>
			/>&nbsp;<span class="C">e</span>&nbsp;<input class="Cc" maxlength="10" style="width:70px;" name="c_dt_cancel_termino" id="c_dt_cancel_termino" onblur="if (!isDate(this)) {alert('Data inválida!'); this.focus();}" onkeypress="if (digitou_enter(true)) bCONFIRMA.focus(); else fFILTRO.ckb_st_entrega_cancelado.checked=true; filtra_data();" onclick="fFILTRO.ckb_st_entrega_cancelado.checked = true;"  onchange="if (trim(this.value)!='') fFILTRO.ckb_st_entrega_cancelado.checked=true;"
			<% if c_dt_cancel_termino <> "" then Response.Write " value=" & chr(34) & c_dt_cancel_termino & chr(34)%>
			/>
		</td>
	</tr>
	</table>
</td></tr>



<!--  VENDEDOR  -->
<% if operacao_permitida(OP_LJA_CONSULTA_UNIVERSAL_PEDIDO_ORCAMENTO , s_lista_operacoes_permitidas) then  %>
<tr bgcolor="#FFFFFF">
<td class="MDBE" align="left" nowrap><span class="PLTe">VENDEDOR</span>
	<br>
	<table cellspacing="0" cellpadding="0" style="margin-bottom:10px; width: 100%">
	<tr bgcolor="#FFFFFF" style="width: 70px">
		<td align="right"><span class="C" style="margin-left:0px;">Vendedor</span></td>
		<td align="left">
			<select id="c_vendedor" name="c_vendedor" style="margin-right:100px;">
			    <option  value="">&nbsp;</option>
			    <% =vendedores_monta_itens_select(c_vendedor) %>
            
			</select>

		</td>
		</td>
	</tr>
	</table>
</td>
</tr>
<%else %>
<input type="hidden" id="c_vendedor" name="c_vendedor" value="<%=usuario%>">
<%end if %>
<!--  MOTIVO  -->
<tr bgcolor="#FFFFFF">
<td class="MDBE" align="left" nowrap><span class="PLTe">MOTIVO</span>
	<br>
	<table cellspacing="0" cellpadding="0" style="margin-bottom:10px; width: 100%">
	<tr bgcolor="#FFFFFF" style="width: 70px">
		<td align="right"><span class="C" style="margin-left:0px;">Descrição</span></td>
		<td align="left">
			<select id="c_motivo" name="c_motivo" style="margin-right:225px;"  onchange='CarregaSubMotivo(this.value,"<%=GRUPO_T_CODIGO_DESCRICAO__CANCELAMENTOPEDIDO_MOTIVO%>");'>			
			 <%=codigo_descricao_monta_itens_select_all(GRUPO_T_CODIGO_DESCRICAO__CANCELAMENTOPEDIDO_MOTIVO, "")%>
			</select>
		</td>
		</td>
	</tr>
    <tr id="motivo_alternativo">
       <td align="right"><span class="C" style="margin-left:0px;">Sub-Motivo</span></td>
		<td align="left">
			<select id="c_cod_sub_motivo" name="c_cod_sub_motivo" style="margin-right:225px;margin-top:4px;">			
			    
			</select>
		</td>
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