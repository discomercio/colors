<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =================================================================
'	  RelFarolProdutoCompradoGravaDados.asp
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

	class cl_REL_FAROL_PROD_COMPRADOS_GRAVA_DADOS
		dim fabricante
		dim produto
		dim qtde_comprada
		dim qtde_comprada_original
		end class
	
	dim s, usuario, msg_erro, s_log
	s_log = ""

	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	dim alerta
	alerta=""

'	OBTÉM FILTROS
	dim c_filtro_fabricante
	c_filtro_fabricante = Trim(Request.Form("c_filtro_fabricante"))

'	OBTÉM DADOS DO FORMULÁRIO
	dim i, j, n, s_produto
	dim blnAchou

'	OBTÉM A QTDE COMPRADA ESPECIFICADA P/ CADA PRODUTO
	dim v_qtde_comprada
	redim v_qtde_comprada(0)
	set v_qtde_comprada(Ubound(v_qtde_comprada)) = New cl_REL_FAROL_PROD_COMPRADOS_GRAVA_DADOS
	v_qtde_comprada(UBound(v_qtde_comprada)).produto = ""
	
	n = Request.Form("c_prod").Count
	for i = 1 to n
		s_produto = Trim(Request.Form("c_prod")(i))
		if s_produto <> "" then
			if Trim(v_qtde_comprada(Ubound(v_qtde_comprada)).produto) <> "" then
				redim preserve v_qtde_comprada(UBound(v_qtde_comprada)+1)
				set v_qtde_comprada(Ubound(v_qtde_comprada)) = New cl_REL_FAROL_PROD_COMPRADOS_GRAVA_DADOS
				end if
			with v_qtde_comprada(Ubound(v_qtde_comprada))
				.produto = s_produto
				.fabricante = Trim(Request.Form("c_fabr")(i))
				.qtde_comprada = converte_numero(retorna_so_digitos(Trim(Request.Form("c_qtde")(i))))
				.qtde_comprada_original = converte_numero(retorna_so_digitos(Trim(Request.Form("c_qtde_original")(i))))
				end with
			end if
		next


'	GRAVA OS DADOS
'	==============
	if alerta = "" then
	'	~~~~~~~~~~~~~
		cn.BeginTrans
	'	~~~~~~~~~~~~~
		If Not cria_recordset_pessimista(rs, msg_erro) then 
		'	~~~~~~~~~~~~~~~~
			cn.RollbackTrans
		'	~~~~~~~~~~~~~~~~
			Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
			end if
			
		for i=Lbound(v_qtde_comprada) to Ubound(v_qtde_comprada)
			if (v_qtde_comprada(i).produto <> "") And (v_qtde_comprada(i).qtde_comprada <> v_qtde_comprada(i).qtde_comprada_original) then
				s = "SELECT " & _
						"*" & _
					" FROM t_PRODUTO" & _
					" WHERE" & _
						" (fabricante = '" & v_qtde_comprada(i).fabricante & "')" & _
						" AND (produto = '" & v_qtde_comprada(i).produto & "')"
				if rs.State <> 0 then rs.Close
				rs.Open s, cn
				if rs.Eof then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Produto (" & v_qtde_comprada(i).fabricante & ")" & v_qtde_comprada(i).produto & " não foi encontrado."
				else
				'	HOUVE ALTERAÇÃO?
					if CLng(rs("farol_qtde_comprada")) <> CLng(v_qtde_comprada(i).qtde_comprada) then
					'	INFORMAÇÕES PARA O LOG
						if s_log <> "" then s_log = s_log & "; "
						s_log = s_log & "(" & Trim("" & rs("fabricante")) & ")" & Trim("" & rs("produto")) & ": " & rs("farol_qtde_comprada") & "=>" & v_qtde_comprada(i).qtde_comprada
					'	ATUALIZA OS DADOS
						rs("farol_qtde_comprada") = v_qtde_comprada(i).qtde_comprada
						rs("farol_qtde_comprada_usuario_ult_atualiz") = usuario
						rs("farol_qtde_comprada_dt_hr_ult_atualiz") = Now
						rs.Update
						if Err <> 0 then
							alerta=texto_add_br(alerta)
							alerta=alerta & Cstr(Err) & ": " & Err.Description
							end if
						end if
					end if
				end if
			
		'	SE HOUVE ERRO, CANCELA O LAÇO
			if alerta <> "" then exit for
			next
		
		if alerta = "" then
			if s_log <> "" then
				grava_log usuario, "", "", "", OP_LOG_FAROL_PRODUTO_COMPRADO_GRAVA_DADOS, s_log
				end if
			end if

	'	FINALIZA TRANSAÇÃO
	'	==================
		if alerta = "" then
		'	~~~~~~~~~~~~~~
			cn.CommitTrans
		'	~~~~~~~~~~~~~~
			if Err<>0 then
				alerta=Cstr(Err) & ": " & Err.Description
				end if
		else
		'	~~~~~~~~~~~~~~~~
			cn.RollbackTrans
		'	~~~~~~~~~~~~~~~~
			end if
		end if
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
function fRetornar(f) {
	f.action = "RelFarolProdutoCompradoFiltro.asp?url_back=X";
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
<body onload="window.status='Concluído';bVOLTAR.focus();" link=#000000 alink=#000000 vlink=#000000>

<center>

<form id="f" name="f" method="post">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<!-- FILTROS -->
<input type="hidden" name="c_filtro_fabricante" id="c_filtro_fabricante" value="<%=c_filtro_fabricante%>" />


<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">

<tr>
	<td align="right" valign="bottom"><p class="PEDIDO">Produtos Comprados (Farol)<span class="C">&nbsp;</span></p></td>
</tr>
</table>
<br>
<br>

<div class='MtAviso' style="width:649px;font-weight:bold;border:1pt solid black;" align="center"><p style='margin:5px 2px 5px 2px;'>Dados gravados com sucesso!</p></div>
<br>

<!-- ************   SEPARADOR   ************ -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>

<!-- ************   LINKS: PÁGINA INICIAL / ENCERRA SESSÃO   ************ -->
<table width="649" cellPadding="0" CellSpacing="0">
<tr><td align="right"><span class="Rc">
	<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
	<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
	</span>
</td></tr>
</table>

<!-- ************   BOTÕES   ************ -->
<table class="notPrint" width="649" cellSpacing="0">
<tr>
	<td align="center"><a name="bVOLTAR" id="bVOLTAR" href="javascript:fRetornar(f)" title="Retornar para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
</tr>
</table>
</form>

</center>
</body>

<% end if %>

</html>


<%
'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>