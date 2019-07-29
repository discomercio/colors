<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/global.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ===============================================
'	  RelEstoqueEcommerceFiltro.asp
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

	dim usuario, s, intIdx
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	if (Not operacao_permitida(OP_CEN_REL_ESTOQUE2, s_lista_operacoes_permitidas)) then 
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if
	
'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	
	dim strDtHojeYYYYMMDD, strDtHojeDDMMYYYY
	strDtHojeYYYYMMDD = formata_data_yyyymmdd(Date)
	strDtHojeDDMMYYYY = formata_data(Date)

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
				" Coalesce(nome,'') AS nome" & _
			" FROM t_FABRICANTE" & _
			" WHERE" & _
				" (Coalesce(nome,'') <> '')" & _
			" ORDER BY" & _
				" Coalesce(nome,'')"
	set r = cn.Execute(strSql)
	strResp = ""
  
	do while Not r.eof 
	    
		x = Trim("" & r("nome"))
		strResp = strResp & "<option "
            for i=LBound(v) to UBound(v) 
		        if (id_default<>"") And (v(i)=x) then
		            strResp = strResp & "selected"
		         end if
		   	 next
		strResp = strResp & " value='" & x & "'>"
		strResp = strResp & Trim("" & r("nome")) & "&nbsp;&nbsp;"
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



<html>


<head>
	<title>CENTRAL</title>
	</head>


<script src="<%=URL_FILE__JQUERY%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
function fESTOQConsulta( f ) {
var i, b;
	b=false;
	for (i = 0; i < f.rb_estoque.length; i++) {
	    if (f.rb_estoque[i].checked) {
	        b = true;
	        break;
	    }
	}
	if (!b) {
		alert("Selecione o estoque a ser consultado!!");
		return;
		}

	b = false;
	if (!f.rb_detalhe.checked) {
		alert("Selecione o tipo de detalhamento da consulta!!");
		return;
		}

	b = false;
	for (i = 0; i < f.rb_saida.length; i++) {
		if (f.rb_saida[i].checked) {
			b = true;
			break;
		}
	}
	if (!b) {
		alert("Selecione o tipo de saída do relatório!!");
		return;
	}

	if (trim(f.c_produto.value)!="") {
		if (!isEAN(trim(f.c_produto.value))) {
			if (trim(f.c_fabricante.value)=="") {
				alert("Informe o fabricante do produto!!");
				f.c_fabricante.focus();
				return;
				}
			}
		}
	
	if ((!f.ckb_normais.checked) && (!f.ckb_compostos.checked)) {
	    alert("Selecionar pelo menos uma opção de exportação!");
	    return;
	}
	dCONFIRMA.style.visibility="hidden";
	window.status = "Aguarde ...";

	if (f.rb_saida[1].checked) setTimeout('exibe_botao_confirmar()', 10000);

	f.submit();
}

function exibe_botao_confirmar() {
	dCONFIRMA.style.visibility = "";
	window.status = "";
}
</script>
<script language="JavaScript" type="text/javascript">
    function limpaCampoSelect(c) {
        c.options[0].selected = true;
    }
    function limpaCampoSelectFabricante() {
        $("#c_fabricante").children().prop('selected', false);
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

<style type="text/css">
#rb_estoque {
	margin: 0pt 2pt 0pt 15pt;
	vertical-align: top;
	}
#rb_detalhe {
	margin: 0pt 2pt 0pt 15pt;
	vertical-align: top;
	}
#rb_saida {
	margin: 0pt 2pt 0pt 15pt;
	vertical-align: top;
	}
.rbOpt
{
	vertical-align:bottom;
}
.lblOpt
{
	vertical-align:bottom;
}
</style>


<body onload="if (trim(fESTOQ.c_fabricante.value)=='') fESTOQ.c_fabricante.focus();">
<center>

<form id="fESTOQ" name="fESTOQ" method="post" action="RelEstoqueEcommerceExec.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Estoque (E-Commerce)</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>
<br>

<!--  PARÂMETROS DA CONSULTA  -->
<table class="Qx" cellspacing="0">
<!--  ESTOQUE  -->
	<tr bgcolor="#FFFFFF">
	<td colspan="2" class="MT" nowrap><span class="PLTe">Estoque de Interesse</span>

		<br><input type="radio" class="rbOpt" tabindex="-1" id="rb_estoque" name="rb_estoque" value="<%=ID_ESTOQUE_VENDA%>" <% if get_default_valor_texto_bd(usuario, "RelEstoqueEcommerceFiltro|rb_estoque") <> "VENDA_SHOW_ROOM" then Response.Write " checked" %>><span class="C lblOpt" style="cursor:default" onclick="fESTOQ.rb_estoque[0].click();"
			>Venda</span>

        <br><input type="radio" class="rbOpt venda_show_room" tabindex="-1" id="rb_estoque" name="rb_estoque" value="VENDA_SHOW_ROOM" <% if get_default_valor_texto_bd(usuario, "RelEstoqueEcommerceFiltro|rb_estoque") = "VENDA_SHOW_ROOM" then Response.Write " checked" %>><span class="C lblOpt" style="cursor:default" onclick="fESTOQ.rb_estoque[1].click();"
			>Venda + Show-Room</span> 
        	
	</td>
	</tr>

<!--  TIPO DE DETALHAMENTO  -->
	<tr bgcolor="#FFFFFF">
	<td colspan="2" class="MDBE" nowrap><span class="PLTe">Tipo de Detalhamento</span>

		<br><input type="radio" class="rbOpt" tabindex="-1" id="rb_detalhe" name="rb_detalhe" value="INTERMEDIARIO" checked><label for="rb_detalhe" ><span class="C lblOpt" style="cursor:default" onclick="fESTOQ.rb_detalhe[0].click();"
			>Intermediário (custos médios)</span></label>
	</td>
	</tr>

<!--  SAÍDA DO RELATÓRIO  -->
	<tr bgcolor="#FFFFFF">
	<td colspan="2" class="MDBE" nowrap><span class="PLTe">Saída do Relatório</span>
		<br><input type="radio" class="rbOpt" tabindex="-1" id="rb_saida" name="rb_saida" value="Html" onclick="dCONFIRMA.style.visibility='';" <% if get_default_valor_texto_bd(usuario, "RelEstoqueEcommerceFiltro|rb_saida") = "Html" then Response.Write " checked" %>><span class="C lblOpt" style="cursor:default" onclick="fESTOQ.rb_saida[0].click();dCONFIRMA.style.visibility='';"
			>Html</span>

		<br><input type="radio" class="rbOpt" tabindex="-1" id="rb_saida" name="rb_saida" value="XLS" onclick="dCONFIRMA.style.visibility='';"<% if get_default_valor_texto_bd(usuario, "RelEstoqueEcommerceFiltro|rb_saida") = "XLS" then Response.Write " checked" %>><span class="C lblOpt" style="cursor:default" onclick="fESTOQ.rb_saida[1].click();dCONFIRMA.style.visibility='';"
			>Excel</span>
	</td>
	</tr>

<!--  FABRICANTE  -->
	<tr bgcolor="#FFFFFF">
	<td class="ME MD MB" align="left" nowrap>
		<span class="PLTe">FABRICANTE </span>
		<br>
		<table cellpadding="0" cellspacing="0">
		<tr>
		<td>
			<select id="c_fabricante" name="c_fabricante" class="LST" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;" size="10"style="margin:4px 4px 4px 20px ;width:170px" multiple>
			<% =fabricante_monta_itens_select(get_default_valor_texto_bd(usuario, "RelEstoqueEcommerceFiltro|c_fabricante")) %>
			</select>
		</td>
		<td style="width:1px;"></td>
		<td align="left" valign="top">
			<a name="bLimparFabricante" id="bLimparFabricante" href="javascript:limpaCampoSelectFabricante()" title="limpa o filtro 'Fabricante'">
						<img src="../botao/botao_x_red.gif" style="vertical-align:bottom;margin-bottom:1px;" width="20" height="20" border="0"></a>
		</td>
		</tr>
		</table>
	</td>
	</tr>
<!--  PRODUTO  -->
    <tr bgcolor="#FFFFFF">
    <td class="MDBE" NOWRAP><span class="PLTe">Produto</span>
		<br><input name="c_produto" id="c_produto" class="PLLe" maxlength="13" style="margin-left:2pt;width:100px;" onkeypress="if (digitou_enter(true)) bCONFIRMA.focus(); filtra_produto();" onblur="this.value=normaliza_codigo(this.value,TAM_MIN_PRODUTO); this.value=ucase(trim(this.value));"></td>
    </tr>
<!--  LOJA  -->
    <tr bgcolor="#FFFFFF">
    <td class="MDBE" NOWRAP><span class="PLTe">Loja</span>
		<br><input name="c_loja" id="c_loja" class="PLLe" maxlength="3" style="margin-left:2pt;width:100px;" value="<%=get_default_valor_texto_bd(usuario, "RelEstoqueEcommerceFiltro|c_loja")%>"  onkeypress="if (!digitou_enter(false) && !digitou_char('-')) filtra_numerico();" onblur="this.value=normaliza_lista_lojas(this.value);"></td>
    </tr>
<!-- EMPRESA -->
    <tr bgcolor="#FFFFFF">
		<td colspan="2" class="MDBE" align="left" nowrap><span class="PLTe">Empresa</span>
		<br>
			<select id="c_empresa" name="c_empresa" style="margin:1px 3px 6px 10px;min-width:100px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
			<% =apelido_empresa_nfe_emitente_monta_itens_select(Null) %>
			</select>			
        </td>
    </tr>
<!--  OPÇÕES DE EXPORTAÇÃO  -->
    <tr bgcolor="#FFFFFF">
	<td colspan="2" class="MDBE" nowrap><span class="PLTe">Opções de Exportação </span>
		<br><input type="checkbox" class="rbOpt" tabindex="-1" id="ckb_normais" name="ckb_normais" value="Produtos Normais" style="margin-left:20px"<% if get_default_valor_texto_bd(usuario, "RelEstoqueEcommerceFiltro|ckb_normais") <> "" then Response.Write " checked" %> ><label for="ckb_normais"><span class="C lblOpt" style="cursor:default" onclick="fESTOQ.ckb_normais[0].click();"
			>Produtos Normais</span></label>

		<br><input type="checkbox" class="rbOpt" tabindex="-1" id="ckb_compostos" name="ckb_compostos" value="Produtos Compostos"style="margin-left:20px" <% if get_default_valor_texto_bd(usuario, "RelEstoqueEcommerceFiltro|ckb_compostos") <> "" then Response.Write " checked" %>><label for="ckb_compostos"><span class="C lblOpt" style="cursor:default" onclick="fESTOQ.ckb_compostos[1].click();"
			>Produtos Compostos</span></label>
	</td>
	</tr>

</table>


<!-- ************   SEPARADOR   ************ -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>
<br>


<table width="649" cellspacing="0">
<tr>
	<td><a name="bCANCELA" id="bCANCELA" href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="cancela a operação">
		<img src="../botao/cancelar.gif" width="176" height="55" border="0"></a></td>
	<td align="right"><div name="dCONFIRMA" id="dCONFIRMA"><a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fESTOQConsulta(fESTOQ)" title="executa a consulta">
		<img src="../botao/confirmar.gif" width="176" height="55" border="0"></a></div>
	</td>
</tr>
</table>
</form>

</center>
</body>
</html>
