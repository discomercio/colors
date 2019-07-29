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
'	  MultiCDAssocProdRegraMenu.asp
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

'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	if Not operacao_permitida(OP_CEN_MULTI_CD_ASSOCIACAO_PRODUTO_REGRA, s_lista_operacoes_permitidas) then 
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

' ____________________________________________________________________________
' GRUPO MONTA ITENS SELECT
'
function t_produto_grupo_monta_itens_select(byval id_default)
dim x, r, strResp, ha_default, strSql, v, i
	id_default = Trim("" & id_default)
	v = split(id_default, ", ")
	ha_default=False
	strSql = "SELECT" & _
				" codigo," & _
                " descricao" & _
			" FROM t_PRODUTO_GRUPO" & _
			" WHERE" & _
				" (Coalesce(codigo,'') <> '')" & _
                " AND (inativo = 0)" & _
			" ORDER BY" & _
				" Coalesce(codigo,'')"
	set r = cn.Execute(strSql)
	strResp = "<option value=''>&nbsp;</option>"
	do while Not r.eof 
	    
		x = Trim("" & r("codigo"))
		strResp = strResp & "<option "
            for i=LBound(v) to UBound(v) 
		        if (id_default<>"") And (v(i)=x) then
		            strResp = strResp & "selected"
		         end if
		   	 next
		strResp = strResp & " value='" & x & "'>"
		strResp = strResp & Trim("" & r("codigo")) & "&nbsp;-&nbsp;" & Trim("" & r("descricao"))
		strResp = strResp & "</option>" & chr(13)
		r.MoveNext	
 	loop
		
	t_produto_grupo_monta_itens_select = strResp
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
<script src="<%=URL_FILE__JQUERY%>" language="JavaScript" type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
function fFILTROConfirma( f ) {
	
    if (!$('input[name=rb_regra]:checked').val()) {
        alert("Escolha uma opção de filtro para regra!!");
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


<body>
<center>

<form id="fFILTRO" name="fFILTRO" method="post" action="MultiCDAssocProdRegraExec.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Multi CD: Associação do Produto com a Regra</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>
<br>

<!--  PARÂMETROS  -->
<table class="Qx" cellSpacing="0">

<!--  PRODUTOS  -->
<tr bgColor="#FFFFFF">
<td class="MDBE MT" NOWRAP><span class="PLTe">PRODUTOS</span>
	<br>
	<table cellSpacing="0" cellPadding="0" style="margin-bottom:10px;">
	
	<tr bgColor="#FFFFFF"><td colspan="2" valign="bottom">
		<span class="C" style="cursor:default">Fabricante</span>
        </td><td>
			<input class="Cc" maxlength="3" style="width:40px; margin-left: 5px;" name="c_fabricante" id="c_fabricante" onblur="this.value=normaliza_codigo(this.value, TAM_MIN_FABRICANTE);" onkeypress="if (digitou_enter(true) && tem_info(this.value)) bCONFIRMA.focus(); else this.click(); filtra_numerico();">
		</td></tr>
	<tr bgColor="#FFFFFF"><td colspan="2" valign="bottom">
		<span class="C" style="cursor:default">Produto</span>
        </td><td>
			<input class="Cc" maxlength="8" style="width:80px; margin-left: 5px;" name="c_produto" id="c_produto" onblur="this.value=normaliza_codigo(this.value, TAM_MIN_PRODUTO);" onkeypress="if (digitou_enter(true) && tem_info(this.value)) bCONFIRMA.focus(); else this.click(); filtra_numerico();">
		</td></tr>
    <tr bgColor="#FFFFFF"><td colspan="2" valign="bottom">
		<span class="C" style="cursor:default">Grupo</span>
        </td><td>
			<select name="c_grupo" id="c_grupo" style="margin-left: 5px; margin-right: 10px">
                <%=t_produto_grupo_monta_itens_select(Null) %>
			</select>
		</td></tr>
	</table>
</td></tr>

<!--  REGRA  -->
<tr bgColor="#FFFFFF">
<td class="MDBE" NOWRAP><span class="PLTe">REGRA</span>
	<br>
	<table cellSpacing="0" cellPadding="0" style="margin-bottom:10px;">
	<tr bgColor="#FFFFFF"><td>
		<input type="radio" tabindex="-1" id="rb_regra" name="rb_regra"
			value="COM_REGRA" checked><span class="C" style="cursor:default" 
			onclick="fFILTRO.rb_regra[0].click();">Somente produtos com regra</span>
		</td></tr>
    <tr bgColor="#FFFFFF"><td>
		<input type="radio" tabindex="-1" id="rb_regra" name="rb_regra"
			value="SEM_REGRA" checked><span class="C" style="cursor:default" 
			onclick="fFILTRO.rb_regra[1].click();">Somente produtos sem regra</span>
		</td></tr>
    <tr bgColor="#FFFFFF"><td>
		<input type="radio" tabindex="-1" id="rb_regra" name="rb_regra"
			value="TODOS" checked><span class="C" style="cursor:default" 
			onclick="fFILTRO.rb_regra[2].click();">Todos</span>
		</td></tr>
	</table>
</td></tr>

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