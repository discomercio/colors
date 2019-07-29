<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/Global.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ========================================================
'	  MultiCDAssocProdRegraExec.asp
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
	dim cn, rs, msg_erro
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	if Not operacao_permitida(OP_CEN_MULTI_CD_ASSOCIACAO_PRODUTO_REGRA, s_lista_operacoes_permitidas) then 
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if
	
	dim s, s_filtro, intQtdeProdutos
    dim rb_regra
    dim c_fabricante, c_produto, c_grupo

	rb_regra = Trim(Request.Form("rb_regra"))
    c_fabricante = Trim(Request.Form("c_fabricante"))
    c_produto = Trim(Request.Form("c_produto"))
    c_grupo = Trim(Request.Form("c_grupo"))

    dim alerta
    alerta = ""



' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________

' _____________________________________________
' T_WMS_REGRA_CD_MONTA_ITENS_SELECT
'
function t_wms_regra_cd_monta_itens_select(byval id_default)
dim x, r, s, strResp, ha_default
	id_default = Trim("" & id_default)
	ha_default=False
	s = "SELECT " & _
			"*" & _
		" FROM t_WMS_REGRA_CD" & _
		" WHERE" & _
			" (st_inativo = 0)"
			
	if Trim("" & id_default) <> "" then
		s = s & " OR (id = " & id_default & ")"
		end if

	s = s & " ORDER BY apelido"

	set r = cn.Execute(s)
	strResp = "<option selected value=''>&nbsp;</option>"
	do while Not r.eof 
		x = Trim("" & r("id"))
		if (id_default<>"") And (id_default=x) then
			strResp = strResp & "<option selected"
			ha_default=True
		else
			strResp = strResp & "<option"
			end if
		
		if CLng(r("st_inativo")) = CLng(1) then
			strResp = strResp & " class='ColorRed'"
			end if

		strResp = strResp & " value='" & x & "'>"
		strResp = strResp & Trim("" & r("apelido"))
		strResp = strResp & "</option>" & chr(13)
		r.MoveNext
		loop
		
	t_wms_regra_cd_monta_itens_select = strResp
	r.close
	set r=nothing
end function

' _____________________________________________
' T_WMS_REGRA_CD_MONTA_ITENS_SELECT
'
function t_wms_regra_cd_monta_itens_select_todos(byval id_default)
dim x, r, s, strResp, ha_default
	id_default = Trim("" & id_default)
	ha_default=False
	s = "SELECT * FROM t_WMS_REGRA_CD ORDER BY apelido"

	set r = cn.Execute(s)
	strResp = "<option selected value=''>&nbsp;</option>"
	do while Not r.eof 
		x = Trim("" & r("id"))
		if (id_default<>"") And (id_default=x) then
			strResp = strResp & "<option selected"
			ha_default=True
		else
			strResp = strResp & "<option"
			end if

    	if CLng(r("st_inativo")) = CLng(1) then
		    strResp = strResp & " class='ColorRed'"
		end if

		strResp = strResp & " value='" & x & "'>"
		strResp = strResp & Trim("" & r("apelido"))
		strResp = strResp & "</option>" & chr(13)
		r.MoveNext
		loop
		
	t_wms_regra_cd_monta_itens_select_todos = strResp
	r.close
	set r=nothing
end function

' _____________________________________
' CONSULTA EXECUTA
'
sub consulta_executa
dim s, s_sql, x
dim r
dim cab_table, cab

	s_sql = _
		"SELECT " & _
            "tP.fabricante " & _
	        ",tP.produto " & _
	        ",tP.descricao " & _
	        ",tP.descricao_html " & _
	        ",tP.grupo " & _
	        ",tPWRC.id_wms_regra_cd " & _
	        ",Coalesce(( " & _
			        "SELECT COUNT(*) " & _
			        "FROM t_PRODUTO_X_WMS_REGRA_CD " & _
			        "WHERE (fabricante = tP.fabricante) AND (produto = tP.produto) " & _
			        "), 0) AS qtde_regra " & _
        "FROM t_PRODUTO tP " & _
        "LEFT JOIN t_PRODUTO_X_WMS_REGRA_CD tPWRC ON ( " & _
		        "(tP.fabricante = tPWRC.fabricante) " & _
		        "AND (tP.produto = tPWRC.produto) " & _
		        ") " & _
        "WHERE (excluido_status = 0) " & _
			"AND (tP.fabricante + '|' + tP.produto NOT IN (SELECT fabricante_composto + '|' + produto_composto FROM t_EC_PRODUTO_COMPOSTO))" & _
	        "AND (tP.descricao <> '.')"

    if c_fabricante <> "" then
        s_sql = s_sql & " AND (tP.fabricante = '" & c_fabricante & "')"
    end if
    if c_produto <> "" then
        s_sql = s_sql & " AND (tP.produto = '" & c_produto & "')"
    end if
    if c_grupo <> "" then
        s_sql = s_sql & " AND (tP.grupo = '" & c_grupo & "')"
    end if
	
	if rb_regra = "COM_REGRA" then
		s_sql = "SELECT * FROM (" & s_sql & ") t WHERE (qtde_regra > 0) ORDER BY fabricante, produto"
	elseif rb_regra = "SEM_REGRA" then
		s_sql = "SELECT * FROM (" & s_sql & ") t WHERE (qtde_regra = 0) ORDER BY fabricante, produto"
	else
		s_sql = "SELECT * FROM (" & s_sql & ") t ORDER BY fabricante, produto"
	end if

	cab_table = "<TABLE cellSpacing=0 cellPadding=0>" & chr(13)
	cab = "	<TR style='background:azure' NOWRAP>" & chr(13) & _
		  "		<TD class='MDTE tdFabricante' style='vertical-align:bottom'><P class='Rc'>Fabr</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdProduto' style='vertical-align:bottom'><P class='Rc'>Produto</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdDescricao' style='vertical-align:bottom'><P class='R'>Descrição</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdGrupo' style='align:center;vertical-align:bottom'><P class='Rc'>Grupo</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdRegra' style='vertical-align:bottom'><P class='Rc'>Regra</P></TD>" & chr(13) & _
		  "	</TR>" & chr(13)
	
	x = cab_table & cab
	intQtdeProdutos = 0
	
	set r = cn.execute(s_sql)
	do while Not r.Eof
	
	 ' CONTAGEM
		intQtdeProdutos = intQtdeProdutos + 1

		x = x & "	<TR NOWRAP>" & chr(13)
		
	'> FABRICANTE
		s = Trim("" & r("fabricante"))
		x = x & "		<TD class='MDTE tdFabricante'><P class='Cc'>" & s & "</P>" & chr(13) & _
                "           <input type='hidden' name='c_fabricante_" & Cstr(intQtdeProdutos) & "' id='c_fabricante_" & Cstr(intQtdeProdutos) & "' value='" & s & "'>" & chr(13) & _
                "       </TD>" & chr(13)

    '> PRODUTO
		s = Trim("" & r("produto"))
		x = x & "		<TD class='MTD tdProduto'><P class='Cc'>" & s & "</P></TD>" & chr(13) & _
                "           <input type='hidden' name='c_produto_" & Cstr(intQtdeProdutos) & "' id='c_produto_" & Cstr(intQtdeProdutos) & "' value='" & s & "'>" & chr(13) & _
                "       </TD>" & chr(13)

    '> DESCRIÇÃO
		s = Trim("" & r("descricao_html"))
		x = x & "		<TD class='MTD tdDescricao'><P class='C'>" & s & "</P></TD>" & chr(13)

	'> GRUPO
		s = Trim("" & r("grupo"))
		x = x & "		<TD class='MTD tdGrupo' align='center'><P class='Cc'>" & s & "</P></TD>" & chr(13)

	'> REGRA
        s = Trim("" & r("id_wms_regra_cd"))
		x = x & "		<TD class='MTD tdRegra'>" & chr(13) & _
				"           <select id='c_regra_" & Cstr(intQtdeProdutos) & "' name='c_regra_" & Cstr(intQtdeProdutos) & "' style='min-width:150px;margin:4pt 4pt 4pt 4pt;' onkeyup='if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;'>" & chr(13) & _
                                t_wms_regra_cd_monta_itens_select(s) & _
				"           </select>" & chr(13) & _
				"		</TD>" & chr(13)

		if (intQtdeProdutos mod 100) = 0 then
			Response.Write x
			x = ""
			end if
			
		r.MoveNext
		loop

' ATRIBUIÇÃO MÚLTIPLA
	x = x & "	<TR>" & chr(13) & _
				"		<TD COLSPAN='5' class='MC' style='border-left:0px;border-right:0px;'>&nbsp;</TD>" & chr(13) & _
				"	</TR>" & chr(13) & _
				"	<TR NOWRAP style='background:#E2E27F'>" & chr(13) & _
				"		<TD COLSPAN='5' class='MT'><p class='C'>ATRIBUIR/REMOVER REGRA PARA TODOS OS PRODUTOS</p></TD>" & chr(13) & _
				"	</TR>" & chr(13) & _
                "	<TR NOWRAP>" & chr(13) & _
				"		<TD COLSPAN='5' class='MD ME'>" & chr(13) & _
                "           <select id='c_regra_todos' name='c_regra_todos' style='min-width:150px;margin:4pt 4pt 4pt 4pt;' onkeyup='if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;'>" & chr(13) & _
                                t_wms_regra_cd_monta_itens_select_todos(Null) & _
				"           </select>" & chr(13) & _
                "       </TD>" & chr(13) & _
                "   </TR>" & chr(13) & _
                "	<TR NOWRAP>" & chr(13) & _
				"		<TD COLSPAN='2' width='right' class='ME MB'>" & chr(13) & _
                "           <button type='button' name='bAtribuirTodos' id='bAtribuirTodos' class='Button' onclick='atribuir_todos();' title='atribuir regra selecionada para todos os produtos' style='margin: 6px'>Atribuir a todos</button>" & chr(13) & _
                "       </TD>" & chr(13) & _
				"		<TD COLSPAN='3' width='right' class='MD MB'>" & chr(13) & _
                "           <button type='button' name='bRemoverTodos' id='bRemoverTodos' class='Button' onclick='remover_todos();' title='remover regra selecionada de todos os produtos' style='margin: 6px'>Remover todos</button>" & chr(13) & _
                "       </TD>" & chr(13) & _
				"	</TR>" & chr(13)
	
'	TOTAL GERAL
	if intQtdeProdutos > 0 then
		x = x & "	<TR>" & chr(13) & _
				"		<TD COLSPAN='5' class='' style='border-left:0px;border-right:0px;'>&nbsp;</TD>" & chr(13) & _
				"	</TR>" & chr(13) & _
				"	<TR NOWRAP style='background:honeydew'>" & chr(13) & _
				"		<TD COLSPAN='5' class='MT'><p class='C'>TOTAL: &nbsp;" & intQtdeProdutos & " produto(s)</p></TD>" & chr(13) & _
				"	</TR>" & chr(13)
		end if
	
  ' MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	if intQtdeProdutos = 0 then
		x = cab_table & cab
		x = x & "	<TR NOWRAP>" & chr(13) & _
				"		<TD class='MT' colspan='5'><P class='ALERTA'>&nbsp;NENHUM REGISTRO ENCONTRADO&nbsp;</P></TD>" & chr(13) & _
				"	</TR>" & chr(13)
		end if

  ' FECHA TABELA
	x = x & "</TABLE>" & chr(13)
	
	x = x & "<input type='hidden' name='c_qtde_produtos' id='c_qtde_produtos' value='" & Cstr(intQtdeProdutos) & "'>" & chr(13)

	Response.write x

	if r.State <> 0 then r.Close
	set r=nothing
	
end sub

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



<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY%>" language="JavaScript" type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">

function fRELConfirma(f) {

	dCONFIRMA.style.visibility = "hidden";
	window.status = "Aguarde ...";
	f.submit();
}
function atribuir_todos() {
    var regraSelecionada;
    regraSelecionada = $('select[name=c_regra_todos]').val();

    if ($('select[name=c_regra_todos] option:selected').hasClass('ColorRed')) {
        alert("Esta regra está INATIVA!!\n\nNão é possível atribuí-la ao produto!!");
        $('#c_regra_todos').focus();
        return;
    }
    $('.tdRegra select').val(regraSelecionada);

}
function remover_todos() {
    var regraSelecionada, c, i;
    regraSelecionada = $('select[name=c_regra_todos]').val();
    c = $('.tdRegra select').length;

    if (regraSelecionada == "") {
        alert("Selecione uma regra para removê-la de todos os produtos!!");
        $('#c_regra_todos').focus();
        return;
    }
    for (i = 1; i <= c; i++) {
        if ($('.tdRegra:eq(' + i + ') select').val() == regraSelecionada) {
            $('.tdRegra:eq(' + i + ') select').val('');
        }
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

<link href="<%=URL_FILE__E_CSS%>" Rel="stylesheet" Type="text/css">
<link href="<%=URL_FILE__EPRINTER_CSS%>" Rel="stylesheet" Type="text/css" media="print">
<link href="<%=URL_FILE__ESCREEN_CSS%>" Rel="stylesheet" Type="text/css" media="screen">

<style type="text/css">
html 
{
	overflow-y: scroll;
}
.tdFabricante {
    vertical-align:middle;
    width: 35px;
}
.tdProduto {
    vertical-align: middle;
    width: 60px;
}
.tdDescricao {
    vertical-align: middle;
    width: 280px;
}
.tdGrupo {
    vertical-align: middle;
    width: 50px;
}
.tdRegra {
    vertical-align: middle;
    min-width: 160px;
}
.ColorRed {
	color: red !important;
}
</style>


<body>
<center>

<form id="fREL" name="fREL" method="post" action="MultiCDAssocProdRegraGravaDados.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="680" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Cadastro de Produtos: Atribuição de Regras</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>

<!-- FILTROS -->
<% 
	s_filtro = "<table width='680' cellPadding='0' CellSpacing='0' style='border-bottom:1px solid black' border='0'>" & chr(13)

    if c_fabricante <> "" then s = c_fabricante else s = "Todos"
	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' NOWRAP>" & _
				"<p class='N'>Fabricante:&nbsp;</p></td><td valign='top' width='99%'>" & _
				"<p class='N'>" & s & "</p></td>" & chr(13) & _
				"	</tr>" & chr(13)

    if c_produto <> "" then s = c_produto else s = "Todos"
    s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' NOWRAP>" & _
				"<p class='N'>Produto:&nbsp;</p></td><td valign='top'>" & _
				"<p class='N'>" & s & "</p></td>" & chr(13) & _
				"	</tr>" & chr(13)

    if c_grupo <> "" then s = c_grupo else s = "Todos"
    s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' NOWRAP>" & _
				"<p class='N'>Grupo:&nbsp;</p></td><td valign='top'>" & _
				"<p class='N'>" & s & "</p></td>" & chr(13) & _
				"	</tr>" & chr(13)

    if rb_regra = "COM_REGRA" then
        s = "Somente produtos com regra"
    elseif rb_regra = "SEM_REGRA" then
        s = "Somente produtos sem regra" 
    else
        s = "Todos"
    end if
    s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' NOWRAP>" & _
				"<p class='N'>Regra:&nbsp;</p></td><td valign='top'>" & _
				"<p class='N'>" & s & "</p></td>" & chr(13) & _
				"	</tr>" & chr(13)
	
	s_filtro = s_filtro & "</table>" & chr(13)
	Response.Write s_filtro
%>

<!--  RELATÓRIO  -->
<br>

<% if alerta = "" then
     consulta_executa
   else %>
    <div class="MtAlerta" style="width:400px;FONT-WEIGHT:bold;" align="CENTER">
        <P style='margin:5px 2px 5px 2px;'><%=alerta%></P>
    </div>
<% end if %>

<!-- ************   SEPARADOR   ************ -->
<table width="680" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>
<br>


<table class="notPrint" width="680" cellSpacing="0">
<tr>
	<td><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()"	title="volta para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
	<td align="right"><div name="dCONFIRMA" id="dCONFIRMA"><a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fRELConfirma(fREL)" title="executa a consulta">
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
