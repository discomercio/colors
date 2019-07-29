<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/Global.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =======================================================
'	  RelClientesNegativadosFiltro.asp
'     =======================================================
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

	if Not operacao_permitida(OP_CEN_REL_CLIENTE_SPC, s_lista_operacoes_permitidas) then
		Response.Redirect("aviso.asp?id=" & ERR_NIVEL_ACESSO_INSUFICIENTE)
		end if

	dim intIdx
	dim url_back, strUrlBotaoVoltar
	url_back = Trim(Request("url_back"))
	if url_back <> "" then
		strUrlBotaoVoltar = "Resumo.asp?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))
	else
		strUrlBotaoVoltar = "javascript:history.back()"
		end if


' _____________________________________________________________________________________________
'   CONSULTA TOTAL DE CLIENTES NEGATIVADOS POR MÊS
'
    dim r,mes1,mes2,mes3,data,strSql
    dim qtde1,qtde2,qtde3

    qtde1 = 0
    qtde2 = 0
    qtde3 = 0
    data = Date()
    mes3 = year(data) & "-" & RIGHT("0" & MONTH(data),2)

    mes3 = Year(data) & "-" & Right("0" & Month(data), 2)
    mes2 = Year(dateAdd("m", -1, data)) & "-" & Right("0" & Month(dateAdd("m", -1, data)), 2)
    mes1 = Year(dateAdd("m", -2, data)) & "-" & Right("0" & Month(dateAdd("m", -2, data)), 2)

    strSql = "SELECT convert(VARCHAR(7), data, 121) as mes," & _
                       " count(*) as qtde" & _
                       " FROM t_CLIENTE_SPC_HISTORICO" & _
                       " WHERE (spc_negativado_status = '1')" & _
                       " AND convert(VARCHAR(7), data, 121) IN (" & _
                       " '" & mes1 & "'," & _
                       " '" & mes2 & "'," & _
                       " '" & mes3 & "'" & _
                       " )" & _
                       " AND data >= '" & mes1 & "-01'" & _
                       " GROUP BY convert(VARCHAR(7), data, 121)"

    set r = cn.Execute(strSql)
    	
	do while Not r.eof 
        
        if mes1 = Trim("" & r("mes")) then qtde1 = r("qtde")
        if mes2 = Trim("" & r("mes")) then qtde2 = r("qtde")
        if mes3 = Trim("" & r("mes")) then qtde3 = r("qtde")
        
        r.MoveNext
    loop

    r.close
	set r=nothing
   
' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S
' _____________________________________________________________________________________________


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
<script src="<%=URL_FILE__JQUERY_UI_I18N%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>

<script type="text/javascript">
	$(function() {
		$("#c_dt_inicio").hUtilUI('datepicker_filtro_inicial');
		$("#c_dt_termino").hUtilUI('datepicker_filtro_final');
	});
</script>

<script language="JavaScript" type="text/javascript">
function fFILTROConfirma(f) {
	
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
<link href="<%=URL_FILE__JQUERY_UI_CSS%>" rel="stylesheet" type="text/css">


<body onload="fFILTRO.c_dt_inicio.focus();">
<center>

<form id="fFILTRO" name="fFILTRO" method="post" action="RelClientesNegativadosExec.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black;">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Clientes Negativados (SPC)</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>
<br>

<table class="Qx" cellspacing="0" style="width:320px;">
<!--  CLIENTE  -->
	<tr>
		<td class="ME MD MC" align="left" nowrap><span class="PLTe">CNPJ/CPF</span></td>
	</tr>
	<tr>
		<td class="ME MD MB" align="left">
			<table cellspacing="0" cellpadding="0" style="margin:0px 20px 6px 10px;">
			<tr bgcolor="#FFFFFF">
				<td align="left">
					<textarea rows="6" name="c_cliente_cnpj_cpf" id="c_cliente_cnpj_cpf" style="width:220px;" onkeypress="if (!digitou_enter(false)) filtra_cnpj_cpf();" onblur="this.value=normaliza_lista_cnpj_cpf(this.value);"></textarea>
				</td>
			</tr>
			</table>
		</td>
	</tr>

<!--  NEGATIVAÇÃO  -->
	<tr>
		<td class="ME MD" align="left" nowrap><span class="PLTe">OPÇÕES DE NEGATIVAÇÃO</span></td>
	</tr>
	<tr>
		<td class="ME MD MB" align="left">
			<% intIdx=-1 %>
			<input type="radio" id="rb_negativado" name="rb_negativado" value="1" class="CBOX" style="margin-left:20px;">
			<% intIdx=intIdx+1 %>
			<span style="cursor:default" class="rbLink" onclick="fFILTRO.rb_negativado[<%=Cstr(intIdx)%>].click();">Apenas clientes negativados</span>
			<br />
			<input type="radio" id="rb_negativado" name="rb_negativado" value="0" class="CBOX" style="margin-left:20px;">
			<% intIdx=intIdx+1 %>
			<span style="cursor:default" class="rbLink" onclick="fFILTRO.rb_negativado[<%=Cstr(intIdx)%>].click();">Apenas clientes não negativados</span>
			<br />
			<input type="radio" id="rb_negativado" name="rb_negativado" value="" class="CBOX" style="margin-left:20px;" checked>
			<% intIdx=intIdx+1 %>
			<span style="cursor:default" class="rbLink" onclick="fFILTRO.rb_negativado[<%=Cstr(intIdx)%>].click();">Ambos</span>
		</td>
	</tr>

<!--  UF  -->
	<tr>
		<td class="ME MD" align="left" nowrap><span class="PLTe">UF</span></td>
	</tr>
	<tr>
		<td class="ME MD MB" align="left">
			<select id="c_uf" name="c_uf" style="margin:1px 10px 6px 10px;" 
					onkeyup="if (window.event.keyCode==KEYCODE_DELETE) {this.options[0].selected=true;}">
				<% =UF_monta_itens_select(Null) %>
			</select>
		</td>
	</tr>


<!--  ORDENAÇÃO  -->
	<tr>
		<td class="ME MD" align="left" nowrap><span class="PLTe">ORDENAÇÃO DO RESULTADO</span></td>
	</tr>
	<tr>
		<td class="ME MD MB" align="left">
			<% intIdx=-1 %>
			<input type="radio" id="rb_ordenacao_saida" name="rb_ordenacao_saida" value="ORD_POR_CNPJ" class="CBOX" style="margin-left:20px;" checked>
			<% intIdx=intIdx+1 %>
			<span style="cursor:default" class="rbLink" onclick="fFILTRO.rb_ordenacao_saida[<%=Cstr(intIdx)%>].click();">CNPJ/CPF</span>
			<br />
			<input type="radio" id="rb_ordenacao_saida" name="rb_ordenacao_saida" value="ORD_POR_NOME" class="CBOX" style="margin-left:20px;">
			<% intIdx=intIdx+1 %>
			<span style="cursor:default" class="rbLink" onclick="fFILTRO.rb_ordenacao_saida[<%=Cstr(intIdx)%>].click();">Nome do Cliente</span>
		</td>
	</tr>
    <tr>
        <td  class="ME MD" align="left" nowrap>
            <span class="PLTe">NEGATIVADOS</span>
        </td>
    </tr>
    <tr>
        <td class="ME MD MB" align="left" >
            <table cellspacing="0" cellpadding="0" style="margin:0px 20px 6px 10px;">
                <tr class="MB" >
                    <td class="MTE N" align="center" style="width:110px" nowrap>
                        <span  >Mês </span>
                    </td>
                    <td class="MTE MD N" align="center" style="width:110px" nowrap>
                        <span  >Total</span>
                    </td>
                </tr>
                <tr class="MB" >
                    <td class="MTE N" align="center" style="width:110px" nowrap>
                        <span  ><%=mes_por_extenso(month(mes1 & "-" & 1),3) & "/" & year(mes1 & "-" & 1)%> </span>
                    </td>
                    <td class="MTE MD N" align="center" style="width:110px" nowrap>
                        <span  ><%=qtde1 %></span>
                    </td>
                </tr>
                <tr>
                    <td class="MTE N" align="center" nowrap>
                        <span  ><%=mes_por_extenso(month(mes2 & "-" & 1),3) & "/" & year(mes2 & "-" & 1)%> </span>
                    </td>
                    <td class="MTE MD N" align="center"  nowrap>
                        <span  ><%=qtde2 %></span> 
                    </td>
                </tr>
                <tr>
                    <td class="MTE MB N" align="center" nowrap>
                        <span  ><%=mes_por_extenso(month(mes3 & "-" & 1),3) & "/" & year(mes3 & "-" & 1)%> </span>
                    </td>
                    <td class="MTE MD MB N" align="center" nowrap>
                        <span  ><%=qtde3 %></span>
                    </td>
                </tr>
            </table>
        </td>
    </tr>
</table>

<!-- ************   SEPARADOR   ************ -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black;">
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
