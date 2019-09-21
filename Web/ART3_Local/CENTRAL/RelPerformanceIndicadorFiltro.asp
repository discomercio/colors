<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =================================================================
'	  R E L P E R F O R M A N C E I N D I C A D O R F I L T R O . A S P
'     =================================================================
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





' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________

' ____________________________________________________________________________
' FABRICANTE MONTA ITENS SELECT
'
function fabricante_monta_itens_select(byval id_default)
dim x, r, strResp, ha_default, v, i
	id_default = Trim("" & id_default)
    v = split(id_default, ", ")
	ha_default=False
	set r = cn.Execute("SELECT * FROM t_FABRICANTE ORDER BY fabricante")
	strResp = ""
	do while Not r.eof 
		x = Trim("" & r("fabricante"))
        strResp = strResp & "<option "
        for i=LBound(v) to UBound(v) 
		    if (id_default<>"") And (v(i)=x) then
		        strResp = strResp & "selected"
                ha_default=True
                exit for
		        end if
		   	next

		strResp = strResp & " value='" & x & "'>"
		strResp = strResp & Trim("" & r("fabricante")) & " - " & Trim("" & r("nome"))
		strResp = strResp & "</option>" & chr(13)
		r.MoveNext
		loop
		
	fabricante_monta_itens_select = strResp
	r.close
	set r=nothing
end function



' ____________________________________________________________________________
' VENDEDORES MONTA ITENS SELECT
'
function vendedores_monta_itens_select(byval id_default)
dim x, r, strResp, ha_default, strSql
	id_default = Trim("" & id_default)
	ha_default=False
	strSql = "SELECT DISTINCT" & _
				" t_USUARIO.usuario," & _
				" nome," & _
				" nome_iniciais_em_maiusculas" & _
			" FROM t_USUARIO" & _
				" INNER JOIN t_USUARIO_X_LOJA" & _
					" ON t_USUARIO.usuario=t_USUARIO_X_LOJA.usuario" & _
			 " WHERE" & _
				" (vendedor_loja <> 0)" & _
			 " ORDER BY" & _
				" t_USUARIO.usuario"
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

	if Not ha_default then
		strResp = "<OPTION SELECTED VALUE=''>&nbsp;</OPTION>" & chr(13) & strResp
		end if

	vendedores_monta_itens_select = strResp
	r.close
	set r=nothing
end function




' ____________________________________________________________________________
' INDICADORES MONTA ITENS SELECT
' LEMBRE-SE: O ORÇAMENTISTA É CONSIDERADO AUTOMATICAMENTE UM INDICADOR!!
function indicadores_monta_itens_select(byval id_default)
dim x, r, strResp, ha_default
	id_default = Trim("" & id_default)
	ha_default=False
	set r = cn.Execute("SELECT apelido, razao_social_nome_iniciais_em_maiusculas FROM t_ORCAMENTISTA_E_INDICADOR ORDER BY apelido")
	strResp = ""
	do while Not r.eof 
		x = UCase(Trim("" & r("apelido")))
		if (id_default<>"") And (id_default=x) then
			strResp = strResp & "<OPTION SELECTED"
			ha_default=True
		else
			strResp = strResp & "<OPTION"
			end if
		strResp = strResp & " VALUE='" & x & "'>"
		strResp = strResp & x & " - " & Trim("" & r("razao_social_nome_iniciais_em_maiusculas"))
		strResp = strResp & "</OPTION>" & chr(13)
		r.MoveNext
		loop

	if Not ha_default then
		strResp = "<OPTION SELECTED VALUE=''>&nbsp;</OPTION>" & chr(13) & strResp
		end if

	indicadores_monta_itens_select = strResp
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



<% if False then 'APENAS P/ HABILITAR O INTELLISENSE DURANTE O DESENVOLVIMENTO!! %>
<script src="../Global/jquery.js" language="JavaScript" type="text/javascript"></script>
<% end if %>

<script src="<%=URL_FILE__JQUERY%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_I18N%>" Language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">

    $(function () {
        $("#c_fabricante").change(function () {
            $("#spnCounterFabricante").text($("#c_fabricante :selected").length);
        });

        $("#spnCounterFabricante").text($("#c_fabricante :selected").length);
    });

function formata_ano(ano) {
var s_ano;
	
	s_ano = "" + ano;
	if (s_ano.length == 1) {
		s_ano = "0" + s_ano;
		}

	if (s_ano.length == 2) {
		if ((s_ano * 1) < 80) {
			s_ano = "20" + s_ano;
		}
		else {
			s_ano = "19" + s_ano;
			}
		}

	return s_ano * 1;
}

function limpaCampoFabricante(f) {
    $("#c_fabricante option:selected").removeAttr("selected");
    $("#spnCounterFabricante").text($("#c_fabricante :selected").length);
}

function fFILTROConfirma( f ) {
var s_mes_de, s_ano_de, s_mes_ate, s_ano_ate;
var strDtRefYYYYMMDD, strDtRefDDMMYYYY;
var i, blnFlagOk;

//	PERÍODO DE ENTREGA
	s_mes_de = trim(f.c_mes_inicio.value);
	s_ano_de = trim(f.c_ano_inicio.value);
	s_mes_ate = trim(f.c_mes_termino.value);
	s_ano_ate = trim(f.c_ano_termino.value);

	if ((s_mes_de == "") || (s_ano_de == "") || (s_mes_ate == "") || (s_ano_ate == "") )
	{
		alert("Preencha o período para consulta!");
		return;
	}

	if ((s_ano_ate + s_mes_ate) < (s_ano_de + s_mes_ate))
	{
		alert("Período inválido para consulta!");
		return;
	}

	if ((parseInt(s_ano_ate,10) - parseInt(s_ano_de,10)) > 1)
	{
		alert("Período para consulta superior a 1 ano!");
		return;
	}
	if ((parseInt(s_ano_ate,10) > parseInt(s_ano_de,10)) && (parseInt(s_mes_ate,10) >= parseInt(s_mes_de,10))) 
	{
		alert("Período para consulta superior a 1 ano!");
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


<body onload="fFILTRO.c_vendedor.focus();">
<center>

<form id="fFILTRO" name="fFILTRO" method="post" action="RelPerformanceIndicadorExec.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Relatório de Performance por Indicador</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>
<br>

<table class="Qx" cellSpacing="0">

<!--  VENDEDOR  -->
	<tr bgColor="#FFFFFF">
	<td class="MT" NOWRAP>
	<span class="PLTe">VENDEDOR</span>
	<br>
		<select id="c_vendedor" name="c_vendedor" style="margin-left:10px;margin-right:10px;margin-bottom:8px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
		<% =vendedores_monta_itens_select(Null) %>
		</select>
	</td>
	</tr>

<!--  INDICADOR  -->
	<tr bgColor="#FFFFFF">
	<td class="MDBE" NOWRAP>
	<span class="PLTe">INDICADOR</span>
	<br>
		<select id="c_indicador" name="c_indicador" style="margin:1px 10px 6px 10px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
		<% =indicadores_monta_itens_select(Null) %>
		</select>
	</td>
	</tr>

<!--  FABRICANTE  -->
	<tr>
		<td class="MDBE" NOWRAP align="left">
		    <span class="PLTe">FABRICANTE(S)</span>
            <br />
            <table style="padding:0px;">
                <tr>
                    <td align="left">
                        <select id="c_fabricante" name="c_fabricante" size="6" multiple style="min-width:220px;margin:1px 4px 6px 0px;">
		                <%=fabricante_monta_itens_select(get_default_valor_texto_bd(usuario, "RelPerformanceIndicadorFiltro|c_fabricante")) %>
		                </select>
                    </td>
                    <td style="text-align:left;vertical-align:top;">
				        <a name="bLimparFabricante" id="bLimparFabricante" href="javascript:limpaCampoFabricante(fFILTRO)" title="limpa o filtro 'Fabricante'">
							        <img src="../botao/botao_x_red.gif" style="vertical-align:bottom;margin-bottom:1px;" width="20" height="20" border="0"></a>
                        <br />
                        (<span class="Lbl" id="spnCounterFabricante"></span>)
                    </td>
                </tr>
            </table>
		</td>
	</tr>

<!--  PEDIDOS ENTREGUES ENTRE  -->
	<td class="MDBE" NOWRAP>
		<span class="PLTe" style="vertical-align:middle;cursor:default;margin-right:10px;" >PEDIDOS ENTREGUES ENTRE</span>
		<br>
		<!-- MÊS/ANO INÍCIO -->
		<table cellSpacing="0" cellPadding="0" style="margin-top:6px;margin-bottom:4px;">
		<tr bgColor="#FFFFFF">
		<td align="right" valign="middle">
		<span class="PLLc" style="color:#808080;"style="margin-left:10px;width:100px;text-align:right">&nbsp;Período inicial:</span>
		</td>
		<td style="width:10px;">&nbsp;</td>
		<td valign="middle">
		<span class="PLLc" style="color:#808080;" style="margin-left:60px;">&nbsp;Mês</span>
		<input class="PLBc" maxlength="2" style="width:48px;" name="c_mes_inicio" id="c_mes_inicio"
			onblur="if ( (trim(this.value)!='') && ( (converte_numero(this.value)<=0)||(converte_numero(this.value)>12) ) ) {alert('Mês inválido!'); this.focus();} else {if (converte_numero(this.value)>0) {while (this.value.length<2) this.value='0'+this.value;}}"
			onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fFILTRO.c_ano_inicio.focus(); filtra_numerico();" />
			&nbsp;
			<span class="PLLc" style="color:#808080;">&nbsp;&nbsp;Ano</span>
			<input class="PLBc" maxlength="4" style="width:60px;" name="c_ano_inicio" id="c_ano_inicio"
			onblur="if (trim(this.value)!='') {this.value=formata_ano(this.value);} if ((trim(this.value)!='')&&(this.value)<1980) {alert('Ano inválido!'); this.focus();}"
			onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fFILTRO.c_mes_termino.focus(); filtra_numerico();"
			/>
		</td>
		</tr>

		<!-- MÊS/ANO FIM -->
		<tr bgColor="#FFFFFF">
		<td align="right" valign="middle">
		<span class="PLLc" style="color:#808080;"style="margin-left:10px;width:100px;text-align:right">&nbsp;Período final:</span>
		</td>
		<td style="width:10px;">&nbsp;</td>
		<td valign="middle">
		<span class="PLLc" style="color:#808080;" style="margin-left:60px;">&nbsp;Mês</span>
		<input class="PLBc" maxlength="2" style="width:48px;" name="c_mes_termino" id="c_mes_termino"
			onblur="if ( (trim(this.value)!='') && ( (converte_numero(this.value)<=0)||(converte_numero(this.value)>12) ) ) {alert('Mês inválido!'); this.focus();} else {if (converte_numero(this.value)>0) {while (this.value.length<2) this.value='0'+this.value;}}"
			onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fFILTRO.c_ano_termino.focus(); filtra_numerico();" />
			&nbsp;
			<span class="PLLc" style="color:#808080;">&nbsp;&nbsp;Ano</span>
			<input class="PLBc" maxlength="4" style="width:60px;" name="c_ano_termino" id="c_ano_termino"
			onblur="if (trim(this.value)!='') {this.value=formata_ano(this.value);} if ((trim(this.value)!='')&&(this.value)<1980) {alert('Ano inválido!'); this.focus();}"
			onkeypress="if (digitou_enter(true)&&tem_info(this.value)) bCONFIRMA.focus(); filtra_numerico();"
			/>
		</td>
		</tr>
		</table>

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