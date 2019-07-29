<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ===============================================
'	  TabelaCustoFinanceiroFornecedorFiltro.asp
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

'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	if Not operacao_permitida(OP_CEN_CAD_TABELA_CUSTO_FINANCEIRO_FORNECEDOR, s_lista_operacoes_permitidas) then 
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if






' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________

' ____________________________________________________________________________
' FORNECEDOR COM TABELA CUSTO FINANCEIRO MONTA ITENS SELECT
'
function fornecedor_com_tabela_custo_financeiro_monta_itens_select(byval id_default)
dim x, r, strResp, strItem, ha_default, strSql
	id_default = Trim("" & id_default)
	ha_default=False
	strSql = "SELECT " & _
				"*" & _
			" FROM " & _
				"(" & _ 
					"SELECT" & _
						" fabricante," & _
						" nome," & _
						" (" & _
							"SELECT Coalesce(Count(*),0) AS QtdeReg FROM t_PERCENTUAL_CUSTO_FINANCEIRO_FORNECEDOR WHERE fabricante=tf.fabricante" & _
							") AS QtdeReg" & _
					" FROM t_FABRICANTE tf" & _
				") t" & _
			" WHERE" & _
				" QtdeReg > 0" & _
			" ORDER BY" & _
				" fabricante"
	set r = cn.Execute(strSql)
	strResp = ""
	do while Not r.Eof 
		x = Trim("" & r("fabricante"))
		if (id_default<>"") And (id_default=x) then
			strResp = strResp & "<OPTION SELECTED"
			ha_default=True
		else
			strResp = strResp & "<OPTION"
			end if
		strResp = strResp & " VALUE='" & x & "'>"
		strItem = Trim("" & r("fabricante")) & " - " & iniciais_em_maiusculas(Trim("" & r("nome")))
		strResp = strResp & strItem
		strResp = strResp & "</OPTION>" & chr(13)
		r.MoveNext
		loop

	if Not ha_default then
		strResp = "<OPTION SELECTED VALUE=''>&nbsp;</OPTION>" & chr(13) & strResp
		end if
		
	fornecedor_com_tabela_custo_financeiro_monta_itens_select = strResp
	r.close
	set r=nothing
end function


' ____________________________________________________________________________
' FORNECEDOR MONTA ITENS SELECT
'
function fornecedor_monta_itens_select(byval id_default)
dim x, r, strResp, strItem, ha_default, strSql
	id_default = Trim("" & id_default)
	ha_default=False
	strSql = "SELECT" & _
				" fabricante," & _
				" nome," & _
				" (" & _
					"SELECT Coalesce(Count(*),0) AS QtdeReg FROM t_PERCENTUAL_CUSTO_FINANCEIRO_FORNECEDOR WHERE fabricante=tf.fabricante" & _
					") AS QtdeReg" & _
			" FROM t_FABRICANTE tf" & _
			" ORDER BY" & _
				" fabricante"
	set r = cn.Execute(strSql)
	strResp = ""
	do while Not r.Eof 
		x = Trim("" & r("fabricante"))
		if (id_default<>"") And (id_default=x) then
			strResp = strResp & "<OPTION SELECTED"
			ha_default=True
		else
			strResp = strResp & "<OPTION"
			end if
		strResp = strResp & " VALUE='" & x & "'>"
		strItem = Trim("" & r("fabricante")) & " - " & iniciais_em_maiusculas(Trim("" & r("nome")))
		if CLng(r("QtdeReg")) > 0 then strItem = "* " & strItem
		strResp = strResp & strItem
		strResp = strResp & "</OPTION>" & chr(13)
		r.MoveNext
		loop

	if Not ha_default then
		strResp = "<OPTION SELECTED VALUE=''>&nbsp;</OPTION>" & chr(13) & strResp
		end if
		
	fornecedor_monta_itens_select = strResp
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



<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
function fFILTROConfirma( f ) {

	if (trim(f.c_fabricante.value)=="") {
		alert("Selecione um fornecedor!");
		return;
		}

	if (f.ckb_clonar_tabela.checked) {
		if (trim(f.c_fabricante_a_clonar.value)=="") {
			alert("Informe o fornecedor que deve ser usado como base para a clonagem da tabela de custo financeiro!!");
			return;
			}
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


<body onload="fFILTRO.c_fabricante.focus();">
<center>

<form id="fFILTRO" name="fFILTRO" method="post" action="TabelaCustoFinanceiroFornecedorEdita.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Tabela de Custo Financeiro por Fornecedor</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>
<br>

<!--  TEXTO EXPLICATIVO -->
<table width="350" cellPadding="0" CellSpacing="0">
<tr><td><p class="Expl">OBS:</p></td></tr>
<tr><td>
	<p class="Expl">O fornecedor destacado com asterisco (*) possui uma tabela de custo financeiro já cadastrada.</p>
	</td>
</tr>
</table>
<br>

<table class="Qx" cellSpacing="0">
<!--  FORNECEDOR  -->
	<tr bgColor="#FFFFFF">
	<td class="MT" nowrap><span class="PLTe">FORNECEDOR</span>
	<br>
		<select id="c_fabricante" name="c_fabricante" style="margin-left:10px;margin-right:10px;margin-bottom:8px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
		<% =fornecedor_monta_itens_select(Null) %>
		</select>
	</td>
	</tr>

	<tr>
	<td>&nbsp;</td>
	</tr>
	
	<tr>
	<td class="MT">
		<input type="checkbox" tabindex="-1" id="ckb_clonar_tabela" name="ckb_clonar_tabela"
				style="margin-left:6px;"
				value="S"><span class="C" style="cursor:default;" 
				onclick="fFILTRO.ckb_clonar_tabela.click();">Clonar tabela de outro fornecedor</span>
	<br>
		<select id="c_fabricante_a_clonar" name="c_fabricante_a_clonar" 
				style="margin-left:10px;margin-right:10px;margin-bottom:8px;" 
				onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;"
				onchange="if (this.selectedIndex > 0){fFILTRO.ckb_clonar_tabela.checked=true;}">
		<% =fornecedor_com_tabela_custo_financeiro_monta_itens_select(Null) %>
		</select>
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
	<td><a href="MenuCadastro.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="volta para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
	<td align="right"><div name="dCONFIRMA" id="dCONFIRMA"><a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fFILTROConfirma(fFILTRO)" title="próxima página">
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
