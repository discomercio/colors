<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/Global.asp"    -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ===============================================
'	  R E L C O N T A G E M E S T O Q U E . A S P
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

	dim usuario, s
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

    '	CONECTA COM O BANCO DE DADOS
	dim cn
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	if Not operacao_permitida(OP_CEN_REL_CONTAGEM_ESTOQUE, s_lista_operacoes_permitidas) then 
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if






' ____________________________________________________________________________
' GRUPO MONTA ITENS SELECT
'
function grupo_monta_itens_select(byval id_default)
dim x, r, strResp, ha_default, strSql, v, i, sDescricao
	id_default = Trim("" & id_default)
	v = split(id_default, ", ")
	ha_default=False
	strSql = "SELECT DISTINCT" & _
				" tP.grupo," & _
				" tPG.descricao" & _
			" FROM t_PRODUTO tP" & _
				" LEFT JOIN t_PRODUTO_GRUPO tPG ON (tP.grupo = tPG.codigo)" & _
			" WHERE" & _
				" (LEN(Coalesce(tP.grupo,'')) > 0)" & _
			" ORDER BY" & _
				" tP.grupo"
	set r = cn.Execute(strSql)
	strResp = ""
	do while Not r.eof 
		x = Trim("" & r("grupo"))
		sDescricao = Trim("" & r("descricao"))
		strResp = strResp & "<option "
		for i=LBound(v) to UBound(v) 
			if (id_default<>"") And (v(i)=x) then
				strResp = strResp & "selected"
				end if
			next
		strResp = strResp & " value='" & x & "'>"
		strResp = strResp & Trim("" & r("grupo"))
		if sDescricao <> "" then strResp = strResp & " &nbsp;(" & sDescricao & ")"
		strResp = strResp & "</option>" & chr(13)
		r.MoveNext	
 	loop
		
	grupo_monta_itens_select = strResp
	r.close
	set r=nothing
end function

'----------------------------------------------------------------------------------------------
' SUBGRUPO MONTA ITENS SELECT
function subgrupo_monta_itens_select(byval id_default)
dim x, r, strSql, strResp, ha_default, v, i, sDescricao
	id_default = Trim("" & id_default)
	v = split(id_default, ", ")
	ha_default=False
	strSql = "SELECT DISTINCT tP.subgrupo, tPS.descricao FROM t_PRODUTO tP LEFT JOIN t_PRODUTO_SUBGRUPO tPS ON (tP.subgrupo = tPS.codigo) WHERE LEN(Coalesce(tP.subgrupo,'')) > 0 ORDER by tP.subgrupo"
	set r = cn.Execute(strSql)
	strResp = ""
	do while Not r.eof 
		x = UCase(Trim("" & r("subgrupo")))
		sDescricao = Trim("" & r("descricao"))
		strResp = strResp & "<option "
		for i=LBound(v) to UBound(v) 
			if (id_default<>"") And (v(i)=x) then
				strResp = strResp & "selected"
				end if
			next
		strResp = strResp & " VALUE='" & x & "'>"
		strResp = strResp & x
		if sDescricao <> "" then strResp = strResp & " &nbsp;(" & sDescricao & ")"
		strResp = strResp & "</OPTION>" & chr(13)
		r.MoveNext
		loop
	
	subgrupo_monta_itens_select = strResp
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
<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>

<script type="text/javascript">
    $(function () {
        $("#c_grupo").change(function () {
            $("#spnCounterGrupo").text($("#c_grupo :selected").length);
        });

        $("#c_subgrupo").change(function () {
            $("#spnCounterSubgrupo").text($("#c_subgrupo :selected").length);
        });

        $("#spnCounterGrupo").text($("#c_grupo :selected").length);
        $("#spnCounterSubgrupo").text($("#c_subgrupo :selected").length);
    });

    function limpaCampoSelectGrupo() {
        $("#c_grupo").children().prop('selected', false);
        $("#spnCounterGrupo").text($("#c_grupo :selected").length);
    }
    function limpaCampoSelectSubgrupo() {
        $("#c_subgrupo").children().prop('selected', false);
        $("#spnCounterSubgrupo").text($("#c_subgrupo :selected").length);
    }
</script>

<script language="JavaScript" type="text/javascript">
function fESTOQConsulta(f) {
var i, b;

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

<form id="fESTOQ" name="fESTOQ" method="post" action="RelContagemEstoqueExec.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Contagem de Estoque</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>
<br>

<!--  PARÂMETROS DA CONSULTA  -->
<table class="Qx" cellSpacing="0">
<!--  FABRICANTE  -->
	<tr bgColor="#FFFFFF">
	<td class="MT" NOWRAP><span class="PLTe">FABRICANTE</span>
		<br><input name="c_fabricante" id="c_fabricante" class="PLLe" maxlength="4" style="margin-left:2pt;width:100px;" onkeypress="if (digitou_enter(true)) bCONFIRMA.focus(); filtra_fabricante();" onblur="this.value=normaliza_codigo(this.value,TAM_MIN_FABRICANTE);"
			value="<%=get_default_valor_texto_bd(usuario, "RelContagemEstoque|c_fabricante")%>"
			/>
	</td>
	</tr>

    <!--  EMPRESA  -->
	<tr bgColor="#FFFFFF">
		<td class="MDBE" NOWRAP><span class="PLTe">EMPRESA</span>
		<br>
			<select id="c_empresa" name="c_empresa" style="margin:1px 10px 6px 2px;min-width:100px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
			    <%=apelido_empresa_nfe_emitente_monta_itens_select(get_default_valor_texto_bd(usuario, "RelContagemEstoque|c_empresa")) %>
			</select>
		</td>
	</tr>

	<!-- GRUPO DE PRODUTOS -->
	<tr bgcolor="#FFFFFF">
	<td class="ME MD MB" align="left" nowrap>
		<span class="PLTe">GRUPO DE PRODUTOS</span>
		<br>
		<table cellpadding="0" cellspacing="0">
		<tr>
		<td>
			<select id="c_grupo" name="c_grupo" class="LST" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;" size="10" style="min-width:250px" multiple>
			<% =grupo_monta_itens_select(get_default_valor_texto_bd(usuario, "RelContagemEstoque|c_grupo")) %>
			</select>
		</td>
		<td style="width:1px;"></td>
		<td align="left" valign="top">
			<a name="bLimparGrupo" id="bLimparGrupo" href="javascript:limpaCampoSelectGrupo()" title="limpa o filtro 'Grupo de Produtos'">
						<img src="../botao/botao_x_red.gif" style="vertical-align:bottom;margin-bottom:1px;" width="20" height="20" border="0"></a>
                        <br />
                        (<span class="Lbl" id="spnCounterGrupo"></span>)
		</td>
		</tr>
		</table>
	</td>
	</tr>

	<!-- SUBGRUPO DE PRODUTOS -->
	<tr bgcolor="#FFFFFF">
	<td class="ME MD MB" align="left" nowrap>
		<span class="PLTe">SUBGRUPO DE PRODUTOS</span>
		<br>
		<table cellpadding="0" cellspacing="0">
		<tr>
		<td>
			<select id="c_subgrupo" name="c_subgrupo" class="LST" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;" size="10" style="min-width:250px" multiple>
			<% =subgrupo_monta_itens_select(get_default_valor_texto_bd(usuario, "RelContagemEstoque|c_subgrupo")) %>
			</select>
		</td>
		<td style="width:1px;"></td>
		<td align="left" valign="top">
			<a name="bLimparSubgrupo" id="bLimparSubgrupo" href="javascript:limpaCampoSelectSubgrupo()" title="limpa o filtro 'Subgrupo de Produtos'">
						<img src="../botao/botao_x_red.gif" style="vertical-align:bottom;margin-bottom:1px;" width="20" height="20" border="0"></a>
                        <br />
                        (<span class="Lbl" id="spnCounterSubgrupo"></span>)
		</td>
		</tr>
		</table>
	</td>
	</tr>

<!--  SAÍDA DO RELATÓRIO  -->
	<tr bgColor="#FFFFFF">
	<td class="MDBE" NOWRAP><span class="PLTe">SAÍDA DO RELATÓRIO</span>
		<br><input type="radio" class="rbOpt" tabindex="-1" id="rb_saida" name="rb_saida" value="Html" onclick="dCONFIRMA.style.visibility='';" checked><span class="C lblOpt" style="cursor:default" onclick="fESTOQ.rb_saida[0].click();dCONFIRMA.style.visibility='';"
			>Html</span>

		<br><input type="radio" class="rbOpt" tabindex="-1" id="rb_saida" name="rb_saida" value="XLS" onclick="dCONFIRMA.style.visibility='';"><span class="C lblOpt" style="cursor:default" onclick="fESTOQ.rb_saida[1].click();dCONFIRMA.style.visibility='';"
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
