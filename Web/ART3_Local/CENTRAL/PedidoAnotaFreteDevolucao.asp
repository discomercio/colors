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
'	  PedidoAnotaFreteDevolucao.asp
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

	dim s
	dim idx, intCounter

'	CONECTA COM O BANCO DE DADOS
	dim cn
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	if Not operacao_permitida(OP_CEN_ANOTA_VALOR_FRETE_NO_PEDIDO, s_lista_operacoes_permitidas) then 
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

' Constantes locais
    Const COD_TIPO_FRETE__ENTREGA_NORMAL = "001"
    Const COD_TIPO_FRETE__DEVOLUCAO = "002"
    Const COD_TIPO_FRETE__REENTREGA = "003"
    Const COD_TIPO_FRETE__AGENDAMENTO = "004"
    Const COD_TIPO_FRETE__TAXA_DESCARREGAMENTO = "005"

' =============================
' F U N Ç Õ E S  
' =============================

' _____________________________________________
' TIPO_FRETE_MONTA_ITENS_SELECT
'
function tipo_frete_monta_itens_select(byval id_default)
dim x, r, strResp
	id_default = Trim("" & id_default)

	set r = cn.Execute("SELECT * FROM t_CODIGO_DESCRICAO WHERE grupo='Pedido_TipoFrete' AND st_inativo=0")
	strResp = ""
	do while Not r.eof 
        if Trim(r("codigo")) = COD_TIPO_FRETE__DEVOLUCAO Or Trim(r("codigo")) = COD_TIPO_FRETE__REENTREGA then
		    x = Trim("" & r("codigo"))
		    if (id_default=x) then
			    strResp = strResp & "<option selected"
		    else
			    strResp = strResp & "<option"
			    end if
		    strResp = strResp & " value='" & x & "'>"
		    strResp = strResp & Trim("" & r("descricao"))
		    strResp = strResp & "</option>" & chr(13)
        end if
		r.MoveNext
		loop
	strResp = "<option value=''>&nbsp;</option>" & strResp
	tipo_frete_monta_itens_select = strResp
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
<script src="<%=URL_FILE__JQUERY_MY_GLOBAL%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
	$(document).ready(function() {
		$('#c_transportadora').keypress(trataKeypressTransportadora);
	});
</script>

<script language="JavaScript" type="text/javascript">
var MAX_ITENS_ANOTA_FRETE_PEDIDO=<%=MAX_ITENS_ANOTA_FRETE_PEDIDO%>;

function trataKeypressTransportadora(evento) {
var blnCancelarKeystroke = false;
	try {
		if (HHO.digitouEnter(evento)) {
			blnCancelarKeystroke = true;
			fFILTRO.c_NF[0].focus();
		}
	}
	catch (e) {
		alert(e.message);
	}
	finally {
		if (blnCancelarKeystroke) {
			evento.preventDefault();
			return false;
		}
		else {
			return true;
		}
	}
}

function trataKeyPressCampoNF(idx){
var f;
	f = fFILTRO;
	if (digitou_enter(true)){
		if (idx==0){
			f.c_serie_NF[idx].focus();
			return;
		}
		if (idx>0){
			if (tem_info(f.c_NF[idx].value)||tem_info(f.c_NF[idx-1].value)||tem_info(f.c_pedido[idx-1].value)){
			    f.c_serie_NF[idx].focus();
				return;
			}
		}
		bCONFIRMA.focus();
	}
}
function trataKeyPressCampoSerieNF(idx){
    var f;
    f = fFILTRO;
    if (digitou_enter(true)){
        if (idx==0){
            f.c_valor_nf[idx].focus();
            return;
        }
        if (idx>0){
            if (tem_info(f.c_serie_NF[idx].value)||tem_info(f.c_serie_NF[idx-1].value)||tem_info(f.c_pedido[idx-1].value)){
                f.c_valor_nf[idx].focus();
                return;
            }
        }
        bCONFIRMA.focus();
    }
}

function trataKeyPressCampoValorFrete(idx) {
var f;
	f = fFILTRO;
	if (digitou_enter(true)){
		if ((!tem_info(f.c_valor_frete[idx].value)) && (!tem_info(f.c_pedido[idx].value))) {
			if (idx>0) bCONFIRMA.focus();
			return;
		}
		if (tem_info(f.c_NF[idx].value)&&tem_info(f.c_valor_frete[idx].value)) {
			if (idx==(MAX_ITENS_ANOTA_FRETE_PEDIDO-1)){
				bCONFIRMA.focus();
				return;
			}
			else {
				f.c_pedido[idx].focus();
				return;
			}
		}
		if (tem_info(f.c_valor_frete[idx].value)) f.c_pedido[idx].focus();
	}
}

function trataKeyPressCampoPedido(idx){
var f;
	f = fFILTRO;
	if (digitou_enter(true)){
		if (idx==(MAX_ITENS_ANOTA_FRETE_PEDIDO-1)){
			bCONFIRMA.focus();
			return;
		}
		if (tem_info(f.c_NF[idx].value)||tem_info(f.c_valor_frete[idx].value)||tem_info(f.c_pedido[idx].value)){
		    f.c_tipo_frete[idx].focus();
			return;
		}
		bCONFIRMA.focus();
	}
}

function realca_cor_linha(indice_row) {
var s_row;
	s_row = "#TR_" + indice_row;
	$(s_row).addClass("rowRealcado");
	$(s_row + " input").addClass("rowRealcado");
	$(s_row + " td:first-child").css("background-color","#FFFFFF");
}

function normaliza_cor_linha(indice_row) {
var sow;
	s_row = "#TR_" + indice_row;
	$(s_row).removeClass("rowRealcado");
	$(s_row + " input").removeClass("rowRealcado");
}

function fFILTROConfirma( f ) {
var vl_frete, i, b, ha_item;

	ha_item=false;
	for (i=0; i < f.c_pedido.length; i++) {
	    b = false;
	    if (trim(f.c_NF[i].value) != "") b = true;
	    if (trim(f.c_pedido[i].value)!="") b=true;
	    if (trim(f.c_valor_frete[i].value)!="") b=true;
	    if (trim(f.c_emitente[i].value)!="") b=true;
	    if (trim(f.c_tipo_frete[i].value)!="") b=true;
		
	    if (b) {
	        ha_item=true;
	        if (trim(f.c_emitente[i].value)=="") {
	            alert("Informe o emitente da NF!!");
	            f.c_emitente[i].focus();
	            return;
	        }
	        if (trim(f.c_tipo_frete[i].value)=="") {
	            alert("Informe o tipo de frete!!");
	            f.c_tipo_frete[i].focus();
	            return;
	        }
	        if (trim(f.c_tipo_frete[i].value)=="002") {
	            if (trim(f.c_NF[i].value) == "") {
	            alert("É necessário informar o número da NF!!");
	            f.c_NF[i].focus();
	            return;
	            }
	            if (trim(f.c_serie_NF[i].value) == "") {
	                alert("É necessário informar a série da NF!!");
	                f.c_serie_NF[i].focus();
	                return;
	            }
	            if (trim(f.c_valor_frete[i].value)=="") {
	                alert("Informe o valor do frete!!");
	                f.c_valor_frete[i].focus();
	                return;
	            }
	            vl_frete=converte_numero(f.c_valor_frete[i].value);
	            if (vl_frete<0) {
	                alert("Valor do frete é inválido!!");
	                f.c_valor_frete[i].focus();
	                return;
	            }
	            if(trim(f.c_pedido[i].value) == "") {
	                alert("É necessário informar o número do pedido!!");
	                f.c_pedido[i].focus();
	                return;
	            }
	            if (trim(f.c_valor_nf[i].value)=="") {
	                alert("Informe o valor da NF de devolução!!");
	                f.c_valor_nf.focus();
	                return;
	            }
	        }
	        if (trim(f.c_tipo_frete[i].value)=="003") {
                if ((trim(f.c_NF[i].value) == "") && (trim(f.c_pedido[i].value) == "")) {
	                alert("Informe o número da NF ou número do pedido!!");
	                f.c_NF[i].focus();
	                return;
	            }
	            if (trim(f.c_NF[i].value) != "") {
	                if (trim(f.c_emitente[i].value)=="") {
	                    alert("Informe o emitente da NF!!");
	                    f.c_emitente[i].focus();
	                    return;
	                }
	                if (trim(f.c_serie_NF[i].value) == "") {
	                    alert("Informe a série da NF!!");
	                    f.c_serie_NF[i].focus();
	                    return;
	                }
	            }
	            if (trim(f.c_valor_frete[i].value)=="") {
	                alert("Informe o valor do frete!!");
	                f.c_valor_frete[i].focus();
	                return;
	            }
	            if (trim(f.c_emitente[i].value)=="-1") {
	                alert("Não é permitido anotar frete de Reentrega em uma NF cujo emitente é Cliente!!");
	                f.c_emitente[i].focus();
	                return;
	            }
	            vl_frete=converte_numero(f.c_valor_frete[i].value);
	            if (vl_frete<0) {
	                alert("Valor do frete é inválido!!");
	                f.c_valor_frete[i].focus();
	                return;
	            }
	        }
	    }
	}
	
	if (!ha_item) {
		alert("Nenhuma informação foi preenchida!!");
		f.c_pedido[0].focus();
		return;
		}

	if (trim(f.c_transportadora.value) == "") {
		alert("Informe a transportadora!!");
		f.c_transportadora.focus();
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
<link href="<%=URL_FILE__EPRINTER_CSS%>" Rel="stylesheet" Type="text/css" media="print">

<style type="text/css">
.rowRealcado
{
	background-color:#98FB98;
}
</style>


<body onload="fFILTRO.c_transportadora.focus();">
<center>

<form id="fFILTRO" name="fFILTRO" method="post" action="PedidoAnotaFreteDevolucaoConsiste.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Anotar Frete no Pedido</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>
<br />

<table>
	<tr>
		<td align="left"><span class="PLTe">Transportadora</span></td>
	</tr>
	<tr>
		<td align="left">
		<select id="c_transportadora" name="c_transportadora" style="margin:1px 2px 6px 2px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) {this.options[0].selected=true;}">
		<% =transportadora_monta_itens_select(Null) %>
		</select>
		</td>
	</tr>
</table>

<br />
<br />

<table class="Qx" cellSpacing="0">
	<tr bgColor="#FFFFFF">
	<td align="left">&nbsp;</td>
	<td class="MB" align="left" valign="bottom"><span class="PLTe">Emitente da NF</span></td>    
	<td class="MB" align="left" valign="bottom"><span class="PLTe">Número NF</span></td>
	<td class="MB" align="left" valign="bottom"><span class="PLTe">Série NF</span></td>
	<td class="MB" align="right" valign="bottom"><span class="PLTd">Valor NF (<%=SIMBOLO_MONETARIO%>)</span></td>
	<td class="MB" align="right" valign="bottom"><span class="PLTd">Frete (<%=SIMBOLO_MONETARIO%>)</span></td>
	<td class="MB" align="left" valign="bottom"><span class="PLTe">Pedido</span></td>
	<td class="MB" align="left" valign="bottom"><span class="PLTe">Tipo de frete</span></td>
	</tr>
<%
	for intCounter=1 to MAX_ITENS_ANOTA_FRETE_PEDIDO
		idx = intCounter - 1
%>

	<tr id="TR_<%=intCounter%>">
	<!--  ORDEM  -->
	<td align="right" valign="bottom">
		<span class="PLLd" style="margin-bottom:3px;"><%=Cstr(intCounter)%>.</span>
	</td>
        <!--  EMITENTE  -->
	<td class="MDBE" align="left" nowrap>
		<select name='c_emitente' id='c_emitente' style='width: 130px' onfocus="realca_cor_linha(<%=intCounter%>);" onblur="normaliza_cor_linha(<%=intCounter%>);">
                        <%=nfe_emitente_monta_itens_select(Null)%>
        </select>
		</td>
	    <!--  NÚMERO NF  -->
	<td class="MDB" align="left" nowrap>
		<input maxlength="9" class="PLLe" style="width:80px;" name="c_NF" id="c_NF" onfocus="this.select(); realca_cor_linha(<%=intCounter%>);" onblur="normaliza_cor_linha(<%=intCounter%>); this.value=retorna_so_digitos(trim(this.value)); if (tem_info(this.value)) fFILTRO.c_pedido[<%=idx%>].tabIndex=-1; else fFILTRO.c_pedido[<%=idx%>].tabIndex=0;" onkeypress="trataKeyPressCampoNF(<%=idx%>); filtra_numerico();">
		</td>
        <!--  SÉRIE NF  -->
	<td class="MDB" align="left" nowrap>
		<input maxlength="3" class="PLLe" style="width:50px;" name="c_serie_NF" id="c_serie_NF" onfocus="this.select(); realca_cor_linha(<%=intCounter%>);" onblur="normaliza_cor_linha(<%=intCounter%>); this.value=retorna_so_digitos(trim(this.value)); if (tem_info(this.value)) fFILTRO.c_pedido[<%=idx%>].tabIndex=-1; else fFILTRO.c_pedido[<%=idx%>].tabIndex=0;" onkeypress="trataKeyPressCampoSerieNF(<%=idx%>); filtra_numerico();">
		</td>
    <!--  VALOR NF  -->
	<td class="MDB" align="right" nowrap>
		<input maxlength="12" class="PLLd" style="width:100px;text-align:right;" name="c_valor_nf" id="c_valor_nf" onfocus="this.select(); realca_cor_linha(<%=intCounter%>);" onblur="normaliza_cor_linha(<%=intCounter%>); this.value=formata_moeda(this.value);" onkeypress="if(digitou_enter(true)) fFILTRO.c_valor_frete[<%=idx%>].focus(); filtra_moeda();">
		</td>
	<!--  VALOR FRETE  -->
	<td class="MDB" align="right" nowrap>
		<input maxlength="12" class="PLLd" style="width:100px;text-align:right;" name="c_valor_frete" id="c_valor_frete" onfocus="this.select(); realca_cor_linha(<%=intCounter%>);" onblur="normaliza_cor_linha(<%=intCounter%>); this.value=formata_moeda(this.value);" onkeypress="trataKeyPressCampoValorFrete(<%=idx%>); filtra_moeda();">
		</td>
	<!--  PEDIDO  -->
	<td class="MDB" align="left" nowrap>
		<input maxlength="10" class="PLLe" style="width:80px;" name="c_pedido" id="c_pedido" onfocus="this.select(); realca_cor_linha(<%=intCounter%>);" onblur="normaliza_cor_linha(<%=intCounter%>); if (normaliza_num_pedido(this.value)!='') this.value=normaliza_num_pedido(this.value);" onkeypress="trataKeyPressCampoPedido(<%=idx%>); filtra_pedido();">
		</td>
    <!--  TIPO DE FRETE  -->
	<td class="MDB" align="left" nowrap>
        <select name="c_tipo_frete" id="c_tipo_frete" style="width:150px" onfocus="realca_cor_linha(<%=intCounter%>);" onblur="normaliza_cor_linha(<%=intCounter%>);">
        <%=tipo_frete_monta_itens_select(Null)%>
        </select>
		</td>
	</tr>
<% next %>
</table>


<!-- ************   SEPARADOR   ************ -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc" align="left">&nbsp;</td></tr>
</table>
<br>


<table class="notPrint" width="649" cellSpacing="0">
<tr>
	<td align="left"><a name="bVOLTA" id="bVOLTA" href="resumo.asp?<%=MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="volta para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
	<td align="right"><div name="dCONFIRMA" id="dCONFIRMA"><a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fFILTROConfirma(fFILTRO)" title="vai para a página de verificação dos dados">
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
