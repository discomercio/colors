<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =====================================
'	  FinCadUnidadeNegocioRateioEdita.asp
'     =====================================
'
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


' _____________________________________________________________________________________________
'
'			I N I C I A L I Z A     P Á G I N A     A S P    N O    S E R V I D O R
' _____________________________________________________________________________________________


	On Error GoTo 0
	Err.Clear
	
'	OBTEM O ID
	dim s, strSql, usuario, id_selecionado, operacao_selecionada
	usuario = trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	REGISTRO A EDITAR
	id_selecionado = trim(request("id_selecionado"))
	operacao_selecionada = trim(request("operacao_selecionada"))

	if (operacao_selecionada<>OP_INCLUI) And (operacao_selecionada<>OP_CONSULTA) then Response.Redirect("aviso.asp?id=" & ERR_OPERACAO_NAO_ESPECIFICADA)

	if (operacao_selecionada<>OP_INCLUI) then
		if (id_selecionado="") Or (converte_numero(id_selecionado)=0) then Response.Redirect("aviso.asp?id=" & ERR_ID_NAO_INFORMADO)
		end if

'	CONECTA COM O BANCO DE DADOS
	dim cn, r, rs, msg_erro
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	if Not cria_recordset_otimista(r, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	if Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

	if (operacao_selecionada<>OP_INCLUI) then
		strSql = "SELECT " & _
					"*" & _
				" FROM t_FIN_UNIDADE_NEGOCIO_RATEIO" & _
				" WHERE" & _
					" (id = " & id_selecionado & ")"
		rs.Open strSql, cn
		if Err <> 0 then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
		if rs.EOF then Response.Redirect("aviso.asp?id=" & ERR_ID_NAO_CADASTRADO)
		end if

	dim intQtdeUnidadesNegocioRateio
	intQtdeUnidadesNegocioRateio = 0




' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________

function finPlanoContasContaRestantesMontaItensSelect(byval id_default)
dim x, r, s_sql, strResp, ha_default
	id_default = Trim("" & id_default)
	ha_default=False
	s_sql = "SELECT " & _
				"*" & _
			" FROM t_FIN_PLANO_CONTAS_CONTA" & _
			" WHERE" & _
				" (st_ativo <> 0)" & _
				" AND (" & _
						"(natureza + '|' + Convert(varchar(10), id)) NOT IN " & _
							"(" & _
								"SELECT natureza + '|' + Convert(varchar(10), id_plano_contas_conta) FROM t_FIN_UNIDADE_NEGOCIO_RATEIO" & _
							")" & _
					")" & _
			" ORDER BY" & _
				" natureza, id"
	set r = cn.Execute(s_sql)
	strResp = ""
	do while Not r.Eof 
		x = Trim("" & r("id"))
		if (converte_numero(id_default)<>0) And (converte_numero(id_default)=converte_numero(x)) then
			strResp = strResp & "<OPTION SELECTED"
			ha_default=True
		else
			strResp = strResp & "<OPTION"
			end if
		strResp = strResp & " VALUE='" & Trim("" & r("natureza")) & "|" & Trim("" & r("id")) & "'>"
		strResp = strResp & normaliza_codigo(x,TAM_PLANO_CONTAS__CONTA) & " (" & Trim("" & r("natureza")) & ") - " & Ucase(Trim("" & r("descricao")))
		strResp = strResp & "</OPTION>" & chr(13)
		r.MoveNext
		loop

	if Not ha_default then
		strResp = "<OPTION SELECTED VALUE=''>&nbsp;</OPTION>" & chr(13) & strResp
	else
		strResp = "<OPTION VALUE=''>&nbsp;</OPTION>" & chr(13) & strResp
		end if
		
	finPlanoContasContaRestantesMontaItensSelect = strResp
	r.close
	set r=nothing
end function
%>

<html>


<head>
	<title>CENTRAL ADMINISTRATIVA</title>
	</head>


<%
'		C L I E N T   S I D E   S C R I P T I N G
'
'      CCCCCCC   LLL        IIIII  EEEEEEEEE  NNN   NNN  TTTTTTTTT EEEEEEEEE
'     CCC   CCC  LLL         III   EEE        NNNN  NNN     TTT    EEE
'     CCC        LLL         III   EEE        NNNNN NNN     TTT    EEE
'     CCC        LLL         III   EEEEEE     NNN NNNNN     TTT    EEEEEE
'     CCC        LLL         III   EEE        NNN  NNNN     TTT    EEE
'     CCC   CCC  LLL   LLL   III   EEE        NNN   NNN     TTT    EEE
'      CCCCCCC   LLLLLLLLL  IIIII  EEEEEEEEE  NNN   NNN     TTT    EEEEEEEEE
'
%>

<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
function calcula_perc_total() {
var i, p, t;
	t = 0;
	for (i = 0; i < fCAD.c_perc_rateio.length; i++) {
		p = converte_numero(fCAD.c_perc_rateio[i].value);
		t += p;
	}
	fCAD.c_perc_total_rateio.value = formata_numero(t, 2);
}

function proximaLinha(indice) {
var intQtdeUnidadesNegocioRateio, intIndice;
	intQtdeUnidadesNegocioRateio = parseInt(fCAD.c_qtde_unidades_negocio_rateio.value);
	intIndice = parseInt(indice);
	if (intIndice >= intQtdeUnidadesNegocioRateio) {
		bATUALIZA.focus();
	}
	else {
		fCAD.c_perc_rateio[intIndice + 1].focus();
	}
}

function RemoveRegistro( f ) {
var b;
	b=window.confirm('Confirma a exclusão?');
	if (b){
		f.operacao_selecionada.value=OP_EXCLUI;
		dREMOVE.style.visibility="hidden";
		window.status = "Aguarde ...";
		f.submit();
		}
}

function AtualizaRegistro(f) {
var i, p, perc_total;
	if (trim(f.c_plano_contas_conta.value)=="") {
		alert('Selecione o plano de conta!!');
		f.c_plano_contas_conta.focus();
		return;
	}

	perc_total = 0;
	for (i = 0; i < f.c_perc_rateio.length; i++) {
		p = converte_numero(f.c_perc_rateio[i].value);
		if (p < 0) {
			alert("O percentual não pode ser negativo!!");
			f.c_perc_rateio[i].focus();
			return;
		}
		perc_total = perc_total + p;
	}

	if (perc_total < 100) {
		alert("A soma do rateio não atinge 100%");
		return;
	}
	else if (perc_total > 100) {
		alert("A soma do rateio excede 100%");
		return;
	}
	
//	PARA O CASO DE TER CLICADO NO BOTÃO BACK APÓS TER CLICADO NA OPERAÇÃO EXCLUIR
	f.operacao_selecionada.value=f.operacao_selecionada_original.value;
	dATUALIZA.style.visibility="hidden";
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

<style TYPE="text/css">
#rb_st_ativo {
	margin: 0pt 2pt 1pt 15pt;
	}
#c_plano_contas_conta 
{
	margin-bottom:4pt;
}
#c_unidade_negocio_descricao
{
	width:300px;
	text-align:right;
}
#c_perc_rateio
{
	margin-left:8px;
	width: 60px;
	text-align:right;
	padding-right:4px;
}
#c_perc_total_rateio
{
	margin-left:8px;
	width: 60px;
	text-align:right;
	padding-right:4px;
}
#lblPercentual
{
	margin-left:3px;
}
</style>


<%	if operacao_selecionada=OP_INCLUI then
		s = "fCAD.c_plano_contas_conta.focus()"
	else
		s = "focus()"
		end if
%>
<body onLoad="calcula_perc_total(); <%=s%>">
<center>



<!--  FORMULÁRIO DE CADASTRO  -->

<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
<%	if operacao_selecionada=OP_INCLUI then
		s = "Unidades de Negócio: Cadastro de Novo Rateio"
	else
		s = "Unidades de Negócio: Consulta/Edição de Rateio"
		end if
%>
	<td align="CENTER" vAlign="BOTTOM"><p class="PEDIDO"><%=s%><br><span class="C">&nbsp;</span></p></td>
</tr>
</table>
<br>


<!--  CAMPOS DO CADASTRO  -->
<form id="fCAD" name="fCAD" METHOD="POST" ACTION="FinCadUnidadeNegocioRateioAtualiza.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<INPUT type=HIDDEN name='operacao_selecionada_original' id="operacao_selecionada_original" value='<%=operacao_selecionada%>'>
<INPUT type=HIDDEN name='operacao_selecionada' id="operacao_selecionada" value='<%=operacao_selecionada%>'>
<INPUT type=HIDDEN name='id_selecionado' id="id_selecionado" value="<%=id_selecionado%>">

<!--  ASSEGURA CRIAÇÃO DE UM ARRAY DE CAMPOS, MESMO QUANDO HOUVER SOMENTE 1 LINHA!! -->
<input type="HIDDEN" id="c_unidade_negocio_id" name="c_unidade_negocio_id" value=''>
<input type="HIDDEN" id="c_unidade_negocio_descricao" name="c_unidade_negocio_descricao" value=''>
<input type="HIDDEN" id="c_perc_rateio" name="c_perc_rateio" value=''>

<!-- ************   PLANO DE CONTAS   ************ -->
<table width="649" class="Q" cellSpacing="0">
	<tr>
		<td>
			<p class="R">PLANO DE CONTA</p>
			<% if operacao_selecionada = OP_INCLUI then %>
			<p class="C">
				<select id="c_plano_contas_conta" name="c_plano_contas_conta" style="margin-left:4px;margin-top:4px;" onkeypress="if (digitou_enter(true)) fCAD.c_perc_rateio[1].focus();" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
				<%=finPlanoContasContaRestantesMontaItensSelect("")%>
				</select>
			</p>
			<% else 
				strSql = "SELECT" & _
							" descricao" & _
						" FROM t_FIN_PLANO_CONTAS_CONTA" & _
						" WHERE" & _
							" (id = " & Trim("" & rs("id_plano_contas_conta")) & ")" & _
							" AND (natureza = '" & Trim("" & rs("natureza")) & "')"
				if r.State <> 0 then r.Close
				r.Open strSql, cn
				if r.Eof then
					s = normaliza_codigo(Trim("" & rs("id_plano_contas_conta")),TAM_PLANO_CONTAS__CONTA) & " (" & Trim("" & rs("natureza")) & ")" & " - PLANO DE CONTA NÃO CADASTRADO"
				else
					s = normaliza_codigo(Trim("" & rs("id_plano_contas_conta")),TAM_PLANO_CONTAS__CONTA) & " (" & Trim("" & rs("natureza")) & ")" & " - " & Ucase(Trim("" & r("descricao")))
					end if
			%>
				<p class="C"><%=s%></p>
				<INPUT type=HIDDEN id="c_plano_contas_conta" name="c_plano_contas_conta" value="<%=Trim("" & rs("natureza")) & "|" & Trim("" & rs("id_plano_contas_conta"))%>">
			<% end if %>
		</td>
	</tr>
</table>

<!-- ************   STATUS ATIVO   ************ -->
<table width="649" class="QS" cellSpacing="0">
	<tr>
<%
	dim st_ativo
	st_ativo=false
	if operacao_selecionada=OP_CONSULTA then
		if Cstr(rs("st_ativo")) = Cstr(COD_FIN_ST_ATIVO__ATIVO) then st_ativo=true
	elseif operacao_selecionada=OP_INCLUI then
		st_ativo=true
		end if
%>
		<td width="100%">
		<p class="R">STATUS</p>
		<p class="C">
			<input type="RADIO" id="rb_st_ativo" name="rb_st_ativo" 
				value="<%=COD_FIN_ST_ATIVO__INATIVO%>" 
				class="TA" <%if Not st_ativo then Response.Write(" checked")%>
				><span onclick="fCAD.rb_st_ativo[0].click();" 
				style="cursor:default;color:<%=finStAtivoCor(COD_FIN_ST_ATIVO__INATIVO)%>;"
				><%=finStAtivoDescricao(COD_FIN_ST_ATIVO__INATIVO)%></span
				>&nbsp;</p>
		<p class="C">
			<input type="RADIO" id="rb_st_ativo" name="rb_st_ativo" 
				value="<%=COD_FIN_ST_ATIVO__ATIVO%>" 
				class="TA" <%if st_ativo then Response.Write(" checked")%>
				><span onclick="fCAD.rb_st_ativo[1].click();" 
				style="cursor:default;color:<%=finStAtivoCor(COD_FIN_ST_ATIVO__ATIVO)%>;"
				><%=finStAtivoDescricao(COD_FIN_ST_ATIVO__ATIVO)%></span
				>&nbsp;</p>
		</td>
	</tr>
</table>

<!-- ************   RATEIO   ************ -->
<table width="649" class="QS" cellSpacing="0" cellpadding="3">
	<tr>
		<td width="100%" colspan="2">
		<p class="R">RATEIO</p>
		</td>
	</tr>

<% if operacao_selecionada=OP_CONSULTA then %>
	<%
	'	UNIDADES DE NEGÓCIO JÁ CADASTRADAS NO RATEIO
		strSql = _
			"SELECT" & _
				" tFUN.id," & _
				" tFUN.descricao," & _
				" tFUNRI.perc_rateio" & _
			" FROM t_FIN_UNIDADE_NEGOCIO_RATEIO_ITEM tFUNRI" & _
				" INNER JOIN t_FIN_UNIDADE_NEGOCIO tFUN ON (tFUNRI.id_unidade_negocio=tFUN.id)" & _
			" WHERE" &_
				" (tFUNRI.id_rateio = " & id_selecionado & ")" & _
			" ORDER BY" & _
				" tFUN.descricao"
		if r.State <> 0 then r.Close
		r.Open strSql, cn
		do while Not r.Eof
			intQtdeUnidadesNegocioRateio = intQtdeUnidadesNegocioRateio + 1
	%>
		<tr>
			<input type="hidden" id="c_unidade_negocio_id" name="c_unidade_negocio_id" value="<%=Trim("" & r("id"))%>" />
			<td width="60%" align="right">
			<input type="text" id="c_unidade_negocio_descricao" name="c_unidade_negocio_descricao" class="TA" value="<%=Ucase(Trim("" & r("descricao")))%>" readonly tabindex="-1" />
			</td>
			<td width="40%">
			<input type="text" id="c_perc_rateio" name="c_perc_rateio"
				value="<%=formata_perc(r("perc_rateio"))%>"
				onfocus="this.select();"
				onkeypress="if (digitou_enter(true)&&tem_info(this.value)) {proximaLinha(<%=Cstr(intQtdeUnidadesNegocioRateio)%>);} filtra_percentual();"
				onkeyup="calcula_perc_total();"
				onblur="this.value=formata_numero(this.value,2); if ((converte_numero(this.value)>100)||(converte_numero(this.value)<0)) {alert('Percentual inválido!!');this.focus();} calcula_perc_total();"
				/><span id="lblPercentual" name="lblPercentual">%</span>
			</td>
		</tr>
		<%	r.MoveNext
			loop
		%>
	<%
	'	UNIDADES DE NEGÓCIO AINDA NÃO CADASTRADAS NO RATEIO
		strSql = _
			"SELECT" & _
				" id," & _
				" descricao" & _ 
			" FROM t_FIN_UNIDADE_NEGOCIO" & _
			" WHERE" & _
				" (st_ativo <> 0)" & _
				" AND" & _
					"(" & _
						"id NOT IN " & _
							"(" & _
								"SELECT id_unidade_negocio FROM t_FIN_UNIDADE_NEGOCIO_RATEIO_ITEM WHERE (id_rateio = " & id_selecionado & ")" & _
							")" & _
					")" & _
			" ORDER BY" & _
				" descricao"
		if r.State <> 0 then r.Close
		r.Open strSql, cn
		do while Not r.Eof
			intQtdeUnidadesNegocioRateio = intQtdeUnidadesNegocioRateio + 1
	%>
		<tr>
			<input type="hidden" id="c_unidade_negocio_id" name="c_unidade_negocio_id" value="<%=Trim("" & r("id"))%>" />
			<td width="60%" align="right">
			<input type="text" id="c_unidade_negocio_descricao" name="c_unidade_negocio_descricao" class="TA" value="<%=Ucase(Trim("" & r("descricao")))%>" readonly tabindex="-1" />
			</td>
			<td width="40%">
			<input type="text" id="c_perc_rateio" name="c_perc_rateio"
				value=""
				onkeypress="if (digitou_enter(true)&&tem_info(this.value)) {proximaLinha(<%=Cstr(intQtdeUnidadesNegocioRateio)%>);} filtra_percentual();"
				onkeyup="calcula_perc_total();"
				onblur="this.value=formata_numero(this.value,2); if ((converte_numero(this.value)>100)||(converte_numero(this.value)<0)) {alert('Percentual inválido!!');this.focus();} calcula_perc_total();"
				/><span id="lblPercentual" name="lblPercentual">%</span>
			</td>
		</tr>
		<%	r.MoveNext
			loop
		%>

<% else %>
	<%
	'	UNIDADES DE NEGÓCIO
		strSql = _
			"SELECT" & _
				" id," & _
				" descricao" & _ 
			" FROM t_FIN_UNIDADE_NEGOCIO" & _
			" WHERE" & _
				" (st_ativo <> 0)" & _
			" ORDER BY" & _
				" descricao"
		if r.State <> 0 then r.Close
		r.Open strSql, cn
		do while Not r.Eof
			intQtdeUnidadesNegocioRateio = intQtdeUnidadesNegocioRateio + 1
	%>
		<tr>
			<input type="hidden" id="c_unidade_negocio_id" name="c_unidade_negocio_id" value="<%=Trim("" & r("id"))%>" />
			<td width="60%" align="right">
			<input type="text" id="c_unidade_negocio_descricao" name="c_unidade_negocio_descricao" class="TA" value="<%=Ucase(Trim("" & r("descricao")))%>" readonly tabindex="-1" />
			</td>
			<td width="40%">
			<input type="text" id="c_perc_rateio" name="c_perc_rateio" value="" 
				onkeypress="if (digitou_enter(true)&&tem_info(this.value)) {proximaLinha(<%=Cstr(intQtdeUnidadesNegocioRateio)%>);} filtra_percentual();"
				onkeyup="calcula_perc_total();"
				onblur="this.value=formata_numero(this.value,2); if ((converte_numero(this.value)>100)||(converte_numero(this.value)<0)) {alert('Percentual inválido!!');this.focus();} calcula_perc_total();"
				/><span id="lblPercentual" name="lblPercentual">%</span>
			</td>
		</tr>
		<%	r.MoveNext
			loop
		%>

<% end if %>

	<!-- TOTAL -->
	<tr>
		<td colspan="2">
			<hr style="width:520px;height:2px;" />
		</td>
	</tr>
	<tr>
		<td width="60%" align="right">
			<p class="N" style='color:Navy'>TOTAL</p>
		</td>
		<td width="40%">
			<input type="text" name="c_perc_total_rateio" id="c_perc_total_rateio" readonly tabindex=-1 value='' style='color:Navy' /><span id="lblPercentual" name="lblPercentual">%</span>
		</td>
	</tr>
	<tr><td style="height:8px;"></td></tr>
</table>

<input type="hidden" id="c_qtde_unidades_negocio_rateio" name="c_qtde_unidades_negocio_rateio" value="<%=Cstr(intQtdeUnidadesNegocioRateio)%>" />


<!-- ************   SEPARADOR   ************ -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>
<br>


<table class="notPrint" width="649" cellSpacing="0">
<tr>
	<td><a href="javascript:history.back()" title="cancela as alterações no cadastro">
		<img src="../botao/cancelar.gif" width="176" height="55" border="0"></a></td>
	<%
	s = ""
	if operacao_selecionada=OP_CONSULTA then
		s = "<td align='CENTER'>" & chr(13) & _
				"		<div name='dREMOVE' id='dREMOVE'>" & chr(13) & _
					"			<a href='javascript:RemoveRegistro(fCAD)' title='exclui do banco de dados'>" & chr(13) & _
						"				<img src='../botao/remover.gif' width=176 height=55 border=0>" & chr(13) & _
					"			</a>" & chr(13) & _
				"		</div>" & chr(13) & _
			"	</td>" & chr(13)
		end if
	%><%=s%>
	<td align="RIGHT"><div name="dATUALIZA" id="dATUALIZA">
		<a name="bATUALIZA" id="bATUALIZA" href="javascript:AtualizaRegistro(fCAD)" title="atualiza o cadastro">
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
	if r.State <> 0 then r.Close
	set r = nothing

	if rs.State <> 0 then rs.Close
	set rs = nothing

	cn.Close
	set cn = nothing
%>