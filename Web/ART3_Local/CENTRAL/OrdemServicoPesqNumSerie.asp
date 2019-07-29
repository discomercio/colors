<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =======================================================
'	  O R D E M S E R V I C O P E S Q N U M S E R I E . A S P
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


' _____________________________________________________________________________________________
'
'			I N I C I A L I Z A     P Á G I N A     A S P    N O    S E R V I D O R
' _____________________________________________________________________________________________

	On Error GoTo 0
	Err.Clear
	
	Server.ScriptTimeout = MAX_SERVER_SCRIPT_TIMEOUT_EM_SEG
	
'	OBTEM USUÁRIO
	dim usuario
	usuario = trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	CONECTA COM O BANCO DE DADOS
	dim cn, rs, msg_erro
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

'	PARÂMETROS DE PESQUISA
	dim c_num_serie
	c_num_serie = Trim(Request.Form("c_num_serie"))
	
	dim s_tabela, id_OS_unica_encontrada, qtde_OS_encontradas
	dim blnPesquisaAutomaticaPorSemelhanca

	blnPesquisaAutomaticaPorSemelhanca = False
	s_tabela=executa_consulta(id_OS_unica_encontrada, qtde_OS_encontradas)

	if id_OS_unica_encontrada<>"" then
		Response.Redirect("OrdemServico.asp?num_OS=" & id_OS_unica_encontrada & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")))
		end if




' ________________________________
' E X E C U T A _ C O N S U L T A
'
function executa_consulta(byref id_registro_unico_encontrado, byref qtde_registros)
dim consulta, consulta_base
dim x, cab, s_ult_id, s_num_serie_aux, s_num_pedido
dim blnUsuarioDesejaPesqPorSemelhanca

	id_registro_unico_encontrado = ""
	qtde_registros = 0
	
  ' CABEÇALHO
	cab="<TABLE class='Q' cellSpacing=0>" & chr(13) & _
		"	<TR style='background:azure;'>" & chr(13) & _
		"		<TD class='MD MB' style='width:75px;'><P class='Rc' style='margin-right:2pt;'>DATA</P></TD>" & chr(13) & _
		"		<TD class='MD MB' style='width:50px;'><P class='R' style='margin-right:2pt;'>O.S.</P></TD>" & chr(13) & _
		"		<TD class='MD MB' style='width:75px;'><P class='R' style='margin-left:2pt'>PEDIDO</P></TD>" & chr(13) & _
		"		<TD class='MD MB' style='width:100px;'><P class='R' style='margin-left:2pt'>Nº SÉRIE</P></TD>" & chr(13) & _
		"		<TD class='MB' style='width:240px;'><P class='R' style='margin-left:2pt'>PRODUTO</P></TD>" & chr(13) & _
		"	</TR>" & chr(13)

	'O USUÁRIO PODE DIGITAR O CARACTER ESPECIAL ASTERISCO ('*') P/ INDICAR QUE
	'DESEJA (E COMO DESEJA) PESQUISAR POR SEMELHANÇA.
	blnUsuarioDesejaPesqPorSemelhanca = False
	s_num_serie_aux = c_num_serie
	if (Instr(s_num_serie_aux, "*") > 0) Or (Instr(s_num_serie_aux, BD_CURINGA_TODOS) > 0) then
		blnUsuarioDesejaPesqPorSemelhanca = True
		s_num_serie_aux = Replace(s_num_serie_aux, "*", BD_CURINGA_TODOS)
		end if

	consulta_base = "SELECT DISTINCT" & _
						" a.ordem_servico," & _
						" a.data," & _
						" a.pedido," & _
						" a.fabricante," & _
						" a.produto," & _
						" a.descricao," & _
						" a.descricao_html," & _
						" b.num_serie" & _
					" FROM t_ORDEM_SERVICO a INNER JOIN t_ORDEM_SERVICO_ITEM b" & _
						" ON (a.ordem_servico=b.ordem_servico)" & _
					" WHERE"

	if blnUsuarioDesejaPesqPorSemelhanca then
		consulta = consulta_base & _
					" (num_serie LIKE '" & s_num_serie_aux & "')" & _
					" ORDER BY a.data, a.ordem_servico"
		if rs.State <> 0 then rs.Close
		rs.open consulta, cn
	else
		consulta = consulta_base & _
					" (num_serie = '" & c_num_serie & "')" & _
					" ORDER BY a.data, a.ordem_servico"
		if rs.State <> 0 then rs.Close
		rs.open consulta, cn
		if rs.Eof then
			'SE NÃO ENCONTROU O Nº DE SÉRIE EXATO, PESQUISA AUTOMATICAMENTE POR SEMELHANÇA
			'TUDO QUE TERMINE C/ A TEXTO FORNECIDO
			blnPesquisaAutomaticaPorSemelhanca = True
			s_num_serie_aux = BD_CURINGA_TODOS & c_num_serie
			consulta = consulta_base & _
						" (num_serie LIKE '" & s_num_serie_aux & "')" & _
						" ORDER BY a.data, a.ordem_servico"
			if rs.State <> 0 then rs.Close
			rs.open consulta, cn
			end if
		end if
	
	x=cab
	while not rs.eof 
	  ' CONTAGEM
		qtde_registros = qtde_registros + 1

		x = x & "	<TR NOWRAP >" & chr(13)

	 '> DATA
		x = x & "		<TD class='MDB' NOWRAP valign='top'><P class='Cc' style='margin-left:2pt'>" & _
				formata_data(rs("data")) & "</P></TD>" & chr(13)

	 '> O.S.
		x = x & "		<TD class='MDB' valign='top'><P class='C' style='margin-left:2pt'>" & _
				"<a href='javascript:fOSConcluir(" & chr(34) & rs("ordem_servico") & chr(34) & _
				")' title='clique para consultar a Ordem de Serviço'>" & _
				formata_num_OS_tela(Trim("" & rs("ordem_servico"))) & "</a></P></TD>" & chr(13)

	 '> PEDIDO
		s_num_pedido = Trim("" & rs("pedido"))
		if s_num_pedido = "" then s_num_pedido = "&nbsp;"
		x = x & "		<TD class='MDB' NOWRAP valign='top'><P class='C' style='margin-left:2pt'>" & _
				s_num_pedido & "</P></TD>" & chr(13)

	 '> Nº SÉRIE
		x = x & "		<TD class='MDB' NOWRAP valign='top'><P class='C' style='margin-left:2pt'>" & _
				Trim("" & rs("num_serie")) & "</P></TD>" & chr(13)

	 '> PRODUTO
		x = x & "		<TD class='MB' NOWRAP valign='top'><P class='C' style='margin-left:2pt'>" & _
				produto_formata_descricao_em_html(Trim("" & rs("descricao_html"))) & "</P></TD>" & chr(13)

		x = x & "	</TR>" & chr(13)

		s_ult_id = Trim("" & rs("ordem_servico"))
		
		rs.MoveNext
		wend
	

  ' SE FOI ENCONTRADO APENAS UM ÚNICO REGISTRO POR COMPARAÇÃO EXATA, RETORNA SEU ID
	if Not blnPesquisaAutomaticaPorSemelhanca then
		if qtde_registros = 1 then id_registro_unico_encontrado = s_ult_id
		end if

  ' MOSTRA TOTAL DE O.S.
	x = x & "	<TR NOWRAP style='background: #FFFFDD'>" & chr(13) & _
			"		<TD COLSPAN='5' NOWRAP><P class='Cd'>" & "TOTAL:&nbsp;&nbsp;&nbsp;" & formata_inteiro(qtde_registros) & "</P></TD>" & chr(13) & _
			"	</TR>" & chr(13)

  ' FECHA TABELA
	x = x & "</TABLE>" & chr(13)
	
	executa_consulta = x

End function

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
window.status='Aguarde, executando consulta ...';

function fOSConcluir(s_id){
	window.status = "Aguarde ...";
	fOS.num_OS.value=s_id;
	fOS.submit(); 
}
</script>

<link href="<%=URL_FILE__E_CSS%>" Rel="stylesheet" Type="text/css">
<link href="<%=URL_FILE__EPRINTER_CSS%>" Rel="stylesheet" Type="text/css" media="print">



<body onload="window.status='Concluído';" link=#000000 alink=#000000 vlink=#000000>

<!--  I D E N T I F I C A Ç Ã O  -->
<table width="100%" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom" nowrap><span class="PEDIDO">Pesquisa de Ordem de Serviço pelo Nº Série</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>


<!--  TABELA COM RESULTADO  -->
<br>
<center>
<form method="post" action="OrdemServico.asp" id="fOS" name="fOS">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="c_num_serie" id="c_num_serie" value='<%=c_num_serie%>'>
<input type="hidden" name="num_OS" id="num_OS" value='<%=id_OS_unica_encontrada%>'>

<%if blnPesquisaAutomaticaPorSemelhanca And (qtde_OS_encontradas > 0) then%>
<table>
<tr><td>
	<span class="N" style="color:red">A pesquisa pelo valor exato não obteve resultados.
	<br>O resultado apresentado foi obtido pela pesquisa por valores semelhantes.
	<br>Nº Série: '<i><%=c_num_serie%></i>'</span><span class="Rc">
</td></tr>
</table>
<br>
<%end if%>
<% =s_tabela %>
</form>

<p class="TracoBottom"></p>

<table class="notPrint" cellSpacing="0">
<tr>
	<td align="center"><a href="javascript:history.back();" title="volta para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
</tr>
</table>

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
