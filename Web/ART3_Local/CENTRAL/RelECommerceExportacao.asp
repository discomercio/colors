<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp"        -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ========================================================
'	  RelECommerceExportacao.asp
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

	dim usuario
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO)

'	CONECTA COM O BANCO DE DADOS
	dim cn
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

    dim default_rb_percentual_majoracao
    default_rb_percentual_majoracao = get_default_valor_texto_bd(usuario, "RelECommerceExportacao|rb_percentual_majoracao")

' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________

' ____________________________________________________________________________
' FABRICANTE MONTA ITENS SELECT
'
function fabricante_monta_itens_select(byval id_default)
dim x, r, strResp, ha_default, strSql, i
dim v
	id_default = Trim("" & id_default)
	v = split(id_default, ", ")
	ha_default=False
	strSql = "SELECT DISTINCT" & _
				" Coalesce(fabricante,'') AS fabricante" & _
			" FROM t_PRODUTO" & _
			" WHERE" & _
				" (Coalesce(fabricante,'') <> '')" & _
			" ORDER BY" & _
				" Coalesce(fabricante,'')"
	set r = cn.Execute(strSql)
	strResp = ""
  
	do while Not r.eof 
	    
		x = Trim("" & r("fabricante"))
		strResp = strResp & "<option "
            for i=LBound(v) to UBound(v) 
		        if (id_default<>"") And (v(i)=x) then
		            strResp = strResp & "selected"
		         end if
		   	 next
		strResp = strResp & " value='" & x & "'>"
		strResp = strResp & Trim("" & r("fabricante")) & "&nbsp;&nbsp;"
		strResp = strResp & "</option>" & chr(13)
		r.MoveNext	
 	loop

	fabricante_monta_itens_select = strResp
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
	<title>CENTRAL</title>
	</head>



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
function fFILTROConfirma( f ) {
    var i, blnPercMajChk

	if ((!f.ckb_normais.checked) && (!f.ckb_compostos.checked)) {
		alert("Selecionar pelo menos uma opção de exportação!");
		return;
	}
	if (f.c_loja.value == "") {
	    alert("Informe a LOJA!!");
	    f.c_loja.focus();
	    return;
	}
	if ((f.c_percentual_majoracao.value != "")&&(f.c_percentual_majoracao.value != "0")) {
	    blnPercMajChk = false
	    for (i = 0; i < f.rb_percentual_majoracao.length; i++) {
	        if (f.rb_percentual_majoracao[i].checked) {
	            blnPercMajChk = true;
	        }
	    }
	}
	if (blnPercMajChk == false) {
	    alert("Informe se o percentual de majoração será somado ou subtraído sobre o valor de venda!!");
	    return;
	}

	dCONFIRMA.style.visibility="hidden";
	window.status = "Aguarde ...";

	setTimeout('exibe_botao_confirmar()', 10000);

	f.submit();
}
function filtra_percentual_majoracao() {
    var letra;
    letra = String.fromCharCode(window.event.keyCode);
    if (((letra < "0") || (letra > "9")) && (letra != "-") && (letra != ".") && (letra != ",")) window.event.keyCode = 0;
}

function exibe_botao_confirmar() {
	dCONFIRMA.style.visibility = "";
	window.status = "";
}

function alternaSaida(f) {
    if (f.rb_saida[0].checked == true) {
        f.c_qtde_corte_estoque.disabled = true;
    }
    else {
        f.c_qtde_corte_estoque.disabled = false;
    }
}
function limpaCampoSelectFabricante() {
    $("#c_fabricante_ignorado").children().prop('selected', false);
}
</script>
<script language='JavaScript' type="text/javascript">
function SomenteNumero(e){
    var tecla=(window.event)?event.keyCode:e.which;   
    if((tecla>47 && tecla<58)) return true;
    else{
    	if (tecla==8 || tecla==0) return true;
	else  return false;
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


<body onload="fFILTRO.c_fabricante.focus();">
<center>

<form id="fFILTRO" name="fFILTRO" method="post" action="RelECommerceExportacaoExec.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">E-Commerce - Exportação</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>
<br>

<!--  FABRICANTE A SER IGNORADO -->
<table class="Qx" cellspacing="0">

    <tr bgcolor="#FFFFFF">
	    <td class="MT" align="left" nowrap>
		    <span class="PLTe">FABRICANTE(S) A SER(EM) IGNORADO(S)</span>
		    <br>
		    <table cellpadding="0" cellspacing="0">
		    <tr>
		    <td>
			    <select id="c_fabricante_ignorado" name="c_fabricante_ignorado" class="LST" size="5" style="width:100px; margin-left: 32px; margin-top:5px; margin-bottom: 5px; color:red;" multiple>
			    <% =fabricante_monta_itens_select(get_default_valor_texto_bd(usuario, "RelEcommerceExportacao|c_fabricante_ignorado")) %>
			    </select>
		    </td>
		    <td style="width:1px;"></td>
		    <td align="left" valign="top">
			    <a name="bLimparFabricante" id="bLimparFabricante" href="javascript:limpaCampoSelectFabricante()" title="limpa o filtro 'Fabricante a ser ignorado'">
						    <img src="../botao/botao_x_red.gif" style="vertical-align:bottom;margin-bottom:1px;margin-top:5px" width="20" height="20" border="0"></a>
		    </td>
		    </tr>
		    </table>
	    </td>
	</tr>

<!--  FABRICANTE -->

	<tr bgcolor="#FFFFFF">
	<td class="ME MD MB" align="left" nowrap><span class="PLTe">FABRICANTE</span>
	<br>
		<table cellspacing="0" cellpadding="0" style="margin:0px 20px 6px 30px;">
		<tr bgcolor="#FFFFFF">
			<td align="left">
				<input type="radio" tabindex="-1" id="rb_fabricante" name="rb_fabricante"
					value="UM"><span class="C" style="cursor:default" 
					onclick="fFILTRO.rb_fabricante[0].click();">Fabricante</span>
			</td>
			<td align="left">
				<input maxlength="4" class="Cc" style="width:50px;" name="c_fabricante" id="c_fabricante" onblur="this.value=normaliza_codigo(this.value,TAM_MIN_FABRICANTE);" onkeypress="if (digitou_enter(true)) fFILTRO.c_fabricante_de.focus(); else fFILTRO.rb_fabricante[0].click(); filtra_fabricante();">
			</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td align="left">
				<input type="radio" tabindex="-1" id="rb_fabricante" name="rb_fabricante"
					value="FAIXA"><span class="C" style="cursor:default" 
					onclick="fFILTRO.rb_fabricante[1].click();">Fabricantes</span>
			</td>
			<td align="left">
				<input maxlength="4" class="Cc" style="width:50px;" name="c_fabricante_de" id="c_fabricante_de" onblur="this.value=normaliza_codigo(this.value,TAM_MIN_FABRICANTE);" onkeypress="if (digitou_enter(true)) fFILTRO.c_fabricante_ate.focus(); else fFILTRO.rb_fabricante[1].click(); filtra_fabricante();">
				<span class="C">a</span>
				<input maxlength="4" class="Cc" style="width:50px;" name="c_fabricante_ate" id="c_fabricante_ate" onblur="this.value=normaliza_codigo(this.value,TAM_MIN_FABRICANTE);" onkeypress="if (digitou_enter(true)) fFILTRO.c_produto.focus(); else fFILTRO.rb_fabricante[1].click(); filtra_fabricante();">
			</td>
		</tr>
		</table>
	</td></tr>

<!--  PRODUTO  -->
	<tr bgcolor="#FFFFFF">
	<td class="MDBE" align="left" nowrap><span class="PLTe">PRODUTO</span>
	<br>
		<table cellspacing="0" cellpadding="0" style="margin:0px 20px 6px 30px;">
		<tr bgcolor="#FFFFFF">
			<td align="left">
				<input type="radio" tabindex="-1" id="rb_produto" name="rb_produto"
					value="UM"><span class="C" style="cursor:default" 
					onclick="fFILTRO.rb_produto[0].click();">Produto</span>
			</td>
			<td align="left">
				<input maxlength="13" class="Cc" style="width:100px;" name="c_produto" id="c_produto" onblur="this.value=ucase(normaliza_codigo(this.value,TAM_MIN_PRODUTO));" onkeypress="if (digitou_enter(true)) fFILTRO.c_produto_de.focus(); else fFILTRO.rb_produto[0].click(); filtra_produto();">
			</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td align="left">
				<input type="radio" tabindex="-1" id="rb_produto" name="rb_produto"
					value="FAIXA"><span class="C" style="cursor:default" 
					onclick="fFILTRO.rb_produto[1].click();">Produtos</span>
			</td>
			<td align="left">
				<input maxlength="13" class="Cc" style="width:100px;" name="c_produto_de" id="c_produto_de" onblur="this.value=ucase(normaliza_codigo(this.value,TAM_MIN_PRODUTO));" onkeypress="if (digitou_enter(true)) fFILTRO.c_produto_ate.focus(); else fFILTRO.rb_produto[1].click(); filtra_produto();">
				<span class="C">a</span>
				<input maxlength="13" class="Cc" style="width:100px;" name="c_produto_ate" id="c_produto_ate" onblur="this.value=ucase(normaliza_codigo(this.value,TAM_MIN_PRODUTO));" onkeypress="if (digitou_enter(true)) fFILTRO.c_grupo.focus(); else fFILTRO.rb_produto[1].click(); filtra_produto();">
			</td>
		</tr>
		</table>
	</td></tr>

<!--  GRUPO DE PRODUTOS  -->
	<tr bgcolor="#FFFFFF">
	<td class="MDBE" align="left" nowrap><span class="PLTe">GRUPO DE PRODUTOS</span>
	<br>
		<table cellspacing="0" cellpadding="0" style="margin:0px 20px 6px 30px;">
		<tr bgcolor="#FFFFFF">
			<td align="left">
				<input type="radio" tabindex="-1" id="rb_grupo" name="rb_grupo"
					value="UM"><span class="C" style="cursor:default" 
					onclick="fFILTRO.rb_grupo[0].click();">Grupo</span>
			</td>
			<td align="left">
				<input maxlength="2" class="Cc" style="width:60px;" name="c_grupo" id="c_grupo" onkeypress="if (digitou_enter(true)) fFILTRO.c_grupo_de.focus(); else fFILTRO.rb_grupo[0].click();" onblur="this.value=ucase(this.value);">
			</td>
		</tr>
		<tr bgcolor="#FFFFFF">
			<td align="left">
				<input type="radio" tabindex="-1" id="rb_grupo" name="rb_grupo"
					value="FAIXA"><span class="C" style="cursor:default" 
					onclick="fFILTRO.rb_grupo[1].click();">Grupos</span>
			</td>
			<td align="left">
				<input maxlength="2" class="Cc" style="width:60px;" name="c_grupo_de" id="c_grupo_de" onkeypress="if (digitou_enter(true)) fFILTRO.c_grupo_ate.focus(); else fFILTRO.rb_grupo[1].click();" onblur="this.value=ucase(this.value);">
				<span class="C">a</span>
				<input maxlength="2" class="Cc" style="width:60px;" name="c_grupo_ate" id="c_grupo_ate" onkeypress="if (digitou_enter(true)) fFILTRO.c_loja.focus(); else fFILTRO.rb_grupo[1].click();" onblur="this.value=ucase(this.value);">
			</td>
		</tr>
		</table>
	</td></tr>

<!--  OPÇÕES DE EXPORTAÇÃO -->
	<tr bgcolor="#FFFFFF">
	<td class="MDBE" align="left" nowrap><span class="PLTe">OPÇÕES DE EXPORTAÇÃO</span>
	<br>
		<table cellspacing="0" cellpadding="0" style="margin:0px 20px 6px 30px;">
			<tr bgcolor="#FFFFFF"><td align="left">
				<input type="checkbox" tabindex="-1" id="ckb_normais" name="ckb_normais" value="1" checked>
				<span class="C" style="cursor:default" >Produtos Normais</span>
				</td></tr>
			<tr bgcolor="#FFFFFF"><td align="left">
				<input type="checkbox" tabindex="-1" id="ckb_compostos" name="ckb_compostos" value="2" checked>
				<span class="C" style="cursor:default" >Produtos Compostos</span>
				</td></tr>
		</table>
	</td></tr>

<!-- LOJA -->
    <tr bgcolor="#FFFFFF">
	<td class="MDBE" align="left" nowrap><span class="PLTe">LOJA</span>
	<br />
        <table cellspacing="0" cellpadding="0" style="margin:0px 20px 6px 30px;">
			<tr bgcolor="#FFFFFF"><td align="left" dir="auto">
		<span class="C" style="cursor:default" >Loja</span>
        <input type="text" maxlength="3" id="c_loja" name="c_loja" class="Cc" size="3" value="<%=get_default_valor_texto_bd(usuario, "RelECommerceExportacao|c_loja")%>">
                </td></tr></table>
	</td></tr>

<!-- PERCENTUAL DE MAJORAÇÃO -->
    <tr>
        <td class="MDBE" align="left" nowrap><span class="PLTe">PERCENTUAL DE MAJORAÇÃO</span>
            <br />
            <table cellspacing="0" cellpadding="0" style="margin:5px 20px 6px 30px">
                <tr>
                    <td align="left" colspan="2">
                        <span class="C" style="cursor:default">Percentual</span>
                        <input type="text" maxlength="3" id="c_percentual_majoracao" name="c_percentual_majoracao" class="Cc" size="3" value="<%=get_default_valor_texto_bd(usuario, "RelECommerceExportacao|c_percentual_majoracao")%>" 
                          onblur="this.value=formata_numero(this.value,1);" onkeypress="filtra_percentual_majoracao();" /><span class="C" style="cursor:default">%</span>
                    </td>
                </tr>
                <tr>
                    <td align="left">
                         <input type="radio" name="rb_percentual_majoracao" id="rb_percentual_majoracao" value="1" <%if default_rb_percentual_majoracao = "1" then Response.Write "checked" %> /><span class="C" style="cursor:default;color:green;" 
					onclick="fFILTRO.rb_percentual_majoracao[0].click();">Somar</span>
                    </td>
                </tr>
                <tr>
                    <td align="left">
                        <input type="radio" name="rb_percentual_majoracao" id="rb_percentual_majoracao" value="2" <%if default_rb_percentual_majoracao = "2" then Response.Write "checked" %> /><span class="C" style="cursor:default;color:red;" 
					onclick="fFILTRO.rb_percentual_majoracao[1].click();">Subtrair</span>
                    </td>
                </tr>
            </table>
        </td>
    </tr>

<!--  OPÇÕES DE SAÍDA -->
	<tr bgcolor="#FFFFFF">
	<td class="MDBE" align="left" nowrap><span class="PLTe">SAÍDA</span>
	<br>
		<table cellspacing="0" cellpadding="0" style="margin:0px 20px 6px 30px;">
			<tr bgcolor="#FFFFFF"><td align="left">
				<input type="radio" tabindex="-1" id="rb_saida" name="rb_saida" value="1">
				<span class="C" style="cursor:default" onclick="fFILTRO.rb_saida[0].click();alternaSaida(fFILTRO)">Normal</span>
				</td></tr>
			<tr bgcolor="#FFFFFF"><td align="left">
				<input type="radio" tabindex="-1" id="rb_saida" name="rb_saida" value="2" checked>
				<span class="C" style="cursor:default" onclick="fFILTRO.rb_saida[1].click();alternaSaida(fFILTRO)">Padrão Magento</span>
				</td></tr>
            	
            <tr bgcolor="#FFFFFF">

                <td align="left" colspan="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                    <span class="C">Qtde corte estoque</span>
				    <input maxlength="3" class="Cc" style="width:30px;" name="c_qtde_corte_estoque" id="c_qtde_corte_estoque" value="<%=get_default_valor_texto_bd(usuario, "RelECommerceExportacao|c_qtde_corte_estoque")%>"  onkeypress='return SomenteNumero(event)'>
			    </td></tr>
		</table>
	</td></tr>
    </table>

<!-- ************   SEPARADOR   ************ -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc" align="left">&nbsp;</td></tr>
</table>
<br>


<table width="649" cellspacing="0">
<tr>
	<td align="left"><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()" title="volta para a página anterior">
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
