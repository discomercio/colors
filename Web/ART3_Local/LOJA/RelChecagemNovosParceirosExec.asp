<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ===================================================================================
'     RelChecagemNovosParceirosExec.asp
'     ===================================================================================
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

	dim usuario, loja
	usuario = Trim(Session("usuario_atual"))
	loja = Trim(Session("loja_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 
	If (loja = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	CONECTA COM O BANCO DE DADOS
	dim cn, rs, msg_erro
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	if Not operacao_permitida(OP_LJA_REL_CHECAGEM_NOVOS_PARCEIROS, s_lista_operacoes_permitidas) then 
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	dim alerta
	alerta = ""

	dim c_indicador
	c_indicador = Trim(Request.Form("c_indicador"))
	



' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________

' _____________________________________
' CONSULTA EXECUTA
'
sub consulta_executa
dim r
dim x
dim cab, cab_table
dim s_sql, s_where
dim n_reg, n_reg_total
dim intLargApelido, intLargNome, intLargTelefone
dim strDdd, strTelefone, strDddCel, strTelCel, strListaTelefones

'	CRITÉRIOS DE RESTRIÇÃO
	s_where = "(checado_status=0)" & _
			  " AND " & _
			  "(CONVERT(smallint, loja) = " & loja & ")"
		
'	FILTRO: INDICADOR
	if c_indicador <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (apelido = '" & c_indicador & "')"
		end if
		
'	MONTA SQL DE CONSULTA
	s_sql = "SELECT " & _
				"apelido, " & _
				"razao_social_nome_iniciais_em_maiusculas, " & _
				"ddd, telefone, " & _
				"ddd_cel, tel_cel, " & _
				"nextel" & _
			" FROM t_ORCAMENTISTA_E_INDICADOR" & _
			" WHERE " & _
				s_where & _
			" ORDER BY" & _
				" apelido"
				
				
  ' CABEÇALHO
	intLargApelido = 80
	intLargNome = 250
	intLargTelefone = 95
		
	cab_table = "<TABLE cellSpacing=0 class='MB'>" & chr(13)
	
	cab = _
		"	<TR style='background:azure' NOWRAP>" & chr(13) & _
		"		<TD class='MDTE' valign='bottom' NOWRAP><P style='width:" & CStr(intLargApelido) & "px' class='R'>Apelido</P></TD>" &  chr(13) & _
		"		<TD class='MTD' valign='bottom'><P style='width:" & CStr(intLargNome) & "px' class='R'>Nome</P></TD>" & chr(13) & _
		"		<TD class='MTD' valign='bottom'><P style='width:" & CStr(intLargTelefone) & "px' class='R'>Telefone</P></TD>" & chr(13) & _
		"	</TR>" & chr(13)
	
'	LAÇO P/ LEITURA DO RECORDSET
	x = cab_table & _
		cab
	n_reg = 0
	n_reg_total = 0
	set r = cn.execute(s_sql)
	do while Not r.Eof
	
		n_reg = n_reg + 1
		n_reg_total = n_reg_total + 1

		x = x & "	<TR>" & chr(13)
		
	'>  INDICADOR (APELIDO)
		x = x & "		<TD class='MDTE' valign='top' style='width:" & CStr(intLargApelido) & "px'>" & _
							"<P class='C'>" & _
								"<a href='javascript:fRELConcluir(" & chr(34) & r("apelido") & chr(34) & ")' title='clique para consultar o cadastro'>" & _
								Trim("" & r("apelido")) & _
								"</a>" & _
							"</P>" & _
						"</TD>" & chr(13)

	'>  INDICADOR (NOME)
		x = x & "		<TD class='MTD' valign='top' style='width:" & CStr(intLargNome) & "px'>" & _
							"<P class='Cn'>" & _
								"<a href='javascript:fRELConcluir(" & chr(34) & r("apelido") & chr(34) & ")' title='clique para consultar o cadastro'>" & _
								Trim("" & r("razao_social_nome_iniciais_em_maiusculas")) & _
								"</a>" & _
							"</P>" & _
						"</TD>" & chr(13)

	'>  TELEFONE
		strDdd = Trim("" & r("ddd"))
		strTelefone = Trim("" & r("telefone"))
		strDddCel = Trim("" & r("ddd_cel"))
		strTelCel = Trim("" & r("tel_cel"))
		if strTelefone <> "" then strTelefone = formata_ddd_telefone_ramal(strDdd, strTelefone, "")
		if strTelCel <> "" then strTelCel = formata_ddd_telefone_ramal(strDddCel, strTelCel, "")
		strListaTelefones = strTelefone
		if (strListaTelefones <> "") And (strTelCel <> "") then strListaTelefones = strListaTelefones & "<br>"
		strListaTelefones = strListaTelefones & strTelCel
		if Trim("" & r("nextel")) <> "" then
			if strListaTelefones <> "" then strListaTelefones = strListaTelefones & "<br>"
			strListaTelefones = strListaTelefones & Trim("" & r("nextel"))
			end if
		if strListaTelefones = "" then strListaTelefones = "&nbsp;"
		x = x & "		<TD class='MTD' valign='top' style='width:" & CStr(intLargTelefone) & "px'>" & _
							"<P class='Cn'>" & _
								"<a href='javascript:fRELConcluir(" & chr(34) & r("apelido") & chr(34) & ")' title='clique para consultar o cadastro'>" & _
								strListaTelefones & _
								"</a>" & _
							"</P>" & _
						"</TD>" & chr(13)
						
		x = x & "	</TR>" & chr(13)
		
		if (n_reg mod 100) = 0 then
			Response.Write x
			x = ""
			end if
		
		r.MoveNext
		loop


  ' MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	if n_reg_total = 0 then
		x = cab_table & _
			cab & _
			"	<TR NOWRAP>" & chr(13) & _
			"		<TD class='MT' colspan=3><P class='ALERTA'>&nbsp;NENHUM INDICADOR ENCONTRADO&nbsp;</P></TD>" & chr(13) & _
			"	</TR>" & chr(13)
		end if

  ' FECHA TABELA
	x = x & "</TABLE>" & chr(13)
	
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
	<title>LOJA</title>
	</head>



<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
window.status='Aguarde, executando a consulta ...';

function fRELConcluir(s_id){
	window.status = "Aguarde ...";
	fREL.id_selecionado.value=s_id;
	fREL.submit(); 
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


<% if alerta <> "" then %>
<!-- ************************************************************ -->
<!-- **********  PÁGINA PARA EXIBIR MENSAGENS DE ERRO  ********** -->
<!-- ************************************************************ -->
<body onload="bVOLTAR.focus();">
<center>
<br>
<!--  T E L A  -->
<p class="T">A V I S O</p>
<div class="MtAlerta" style="width:600px;font-weight:bold;" align="center"><p style='margin:5px 2px 5px 2px;'><%=alerta%></p></div>
<br><br>
<p class="TracoBottom"></p>
<table cellSpacing="0">
<tr>
	<td align="center"><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()"><img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
</tr>
</table>
</center>
</body>



<% else %>
<!-- ***************************************************** -->
<!-- **********  PÁGINA PARA EXIBIR RESULTADO   ********** -->
<!-- ***************************************************** -->
<body onload="window.status='Concluído';" link=#000000 alink=#000000 vlink=#000000>

<center>

<form id="fREL" name="fREL" method="post" action="OrcamentistaEIndicadorEdita.asp" >
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="id_selecionado" id="id_selecionado" value=''>
<input type="hidden" name="operacao_selecionada" id="operacao_selecionada" value='<%=OP_CONSULTA%>'>
<input type="hidden" name="c_indicador" id="c_indicador" value='<%=c_indicador%>'>
<% 'NO CASO DE TER FILTRADO P/ APENAS 1 INDICADOR, APÓS SUA CHECAGEM NÃO DEVE RETORNAR P/ A LISTAGEM
	if c_indicador = "" then %>
<input type="hidden" name="pagina_relatorio_originou_edicao" id="pagina_relatorio_originou_edicao" value='RelChecagemNovosParceirosExec.asp'>
<% else %>
<input type="hidden" name="pagina_relatorio_originou_edicao" id="pagina_relatorio_originou_edicao" value='Resumo.asp'>
<% end if %>

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Checagem de Novos Parceiros</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>

<!--  RELATÓRIO  -->
<br>
<% consulta_executa %>

<!-- ************   SEPARADOR   ************ -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>
<br>


<table class="notPrint" width="649" cellSpacing="0">
<tr>
	<td align="center"><a name="bVOLTA" id="bVOLTA" href="resumo.asp?<%=MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="volta para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
</tr>
</table>
</form>

</center>
</body>

<% end if %>

</html>


<%
	if rs.State <> 0 then rs.Close
	set rs = nothing

'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>
