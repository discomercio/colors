<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =============================================================
'	  TabelaComissaoVendedorConfirma.asp
'     =============================================================
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

	dim s, s_aux, strUrlBotaoVoltar, s_log, s_log_tabela_antiga, s_log_tabela_nova, usuario, msg_erro
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs
	if Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	dim i, n, strSql, c_id_perfil, operacao_selecionada, blnHaDados

	strUrlBotaoVoltar = "javascript:history.back()"
	
	dim alerta
	alerta = ""
	
'	OBTÉM DADOS DIGITADOS NO FORMULÁRIO
	operacao_selecionada = Trim(Request("operacao_selecionada"))
	if operacao_selecionada = "" then
		alerta=texto_add_br(alerta)
		alerta=alerta & "Nenhuma operação foi selecionada."
		end if
		
	c_id_perfil = Trim(Request.Form("c_id_perfil"))
	if c_id_perfil = "" then
		alerta=texto_add_br(alerta)
		alerta=alerta & "Perfil não foi informado."
		end if

	dim vArray
	redim vArray(0)
	set vArray(0) = New cl_DUAS_COLUNAS
	n = Request.Form("c_perc_desconto").Count
	for i = 1 to n
		s=Trim(Request.Form("c_perc_desconto")(i))
		if s <> "" then
			if Trim(vArray(ubound(vArray)).c1) <> "" then
				redim preserve vArray(ubound(vArray)+1)
				set vArray(ubound(vArray)) = New cl_DUAS_COLUNAS
				end if
			with vArray(ubound(vArray))
				s = Trim(Request.Form("c_perc_desconto")(i))
				.c1 = converte_numero(s)
				s = Trim(Request.Form("c_perc_comissao")(i))
				.c2 = converte_numero(s)
				end with
			end if
		next
	
	for i = Lbound(vArray) to (Ubound(vArray)-1)
		if vArray(i).c1 = vArray(i+1).c1 then
			alerta=texto_add_br(alerta)
			alerta=alerta & "O percentual de desconto de " & formata_perc_desc(vArray(i).c1) & " aparece em duplicidade."
		elseif vArray(i).c1 > vArray(i+1).c1 then
			alerta=texto_add_br(alerta)
			alerta=alerta & "Percentuais de desconto não estão em ordem crescente."
			end if
		next

	blnHaDados = False
	for i = Lbound(vArray) to (Ubound(vArray))
		if (vArray(i).c1 <> 0) Or (vArray(i).c2 <> 0) then blnHaDados = True
		next
		
	if Not blnHaDados then
		alerta=texto_add_br(alerta)
		alerta=alerta & "Nenhum percentual de comissão foi definido."
		end if
		
	dim erro_fatal
	erro_fatal=False

	if alerta = "" then
		if operacao_selecionada = OP_EXCLUI then
		'	~~~~~~~~~~~~~
			cn.BeginTrans
		'	~~~~~~~~~~~~~
		'	INFORMAÇÕES P/ O LOG
			s_log = ""
			s_log_tabela_antiga = ""
			strSql = "SELECT " & _
						"*" & _
					" FROM t_PERCENTUAL_COMISSAO_VENDEDOR" & _
					" WHERE" & _
						"(id_perfil = '" & c_id_perfil & "')" & _
					" ORDER BY" & _
						" perc_desconto"
			set rs = cn.execute(strSql)
			do while Not rs.Eof
				if s_log_tabela_antiga <> "" then s_log_tabela_antiga = s_log_tabela_antiga & " "
				s_log_tabela_antiga = s_log_tabela_antiga & formata_perc_desc(rs("perc_desconto")) & "=" & formata_perc_desc(rs("perc_comissao"))
				rs.MoveNext
				loop
			
			strSql = "DELETE FROM t_PERCENTUAL_COMISSAO_VENDEDOR WHERE (id_perfil = '" & c_id_perfil & "')"
			cn.Execute(strSql)
			if Err <> 0 then
				erro_fatal=True
				alerta = "FALHA AO EXCLUIR DADOS ANTERIORES DA TABELA t_PERCENTUAL_COMISSAO_VENDEDOR (" & Cstr(Err) & ": " & Err.Description & ")."
				end if
			
			if s_log_tabela_antiga <> "" then
				s_log = "Exclusão da tabela de comissão do vendedor para o perfil " & x_perfil_apelido(c_id_perfil) & ": " & s_log_tabela_antiga
				end if
			
			if s_log <> "" then grava_log usuario, "", "", "", OP_LOG_TABELA_COMISSAO_VENDEDOR, s_log
			
		'	~~~~~~~~~~~~~~
			cn.CommitTrans
		'	~~~~~~~~~~~~~~

			if Err=0 then 
				strUrlBotaoVoltar = "TabelaComissaoVendedorFiltro.asp" & "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))
			else
				alerta=Cstr(Err) & ": " & Err.Description
				end if

		else
		'	~~~~~~~~~~~~~
			cn.BeginTrans
		'	~~~~~~~~~~~~~
		'	INFORMAÇÕES P/ O LOG
			s_log = ""
			s_log_tabela_antiga = ""
			strSql = "SELECT " & _
						"*" & _
					" FROM t_PERCENTUAL_COMISSAO_VENDEDOR" & _
					" WHERE" & _
						"(id_perfil = '" & c_id_perfil & "')" & _
					" ORDER BY" & _
						" perc_desconto"
			set rs = cn.execute(strSql)
			do while Not rs.Eof
				if s_log_tabela_antiga <> "" then s_log_tabela_antiga = s_log_tabela_antiga & " "
				s_log_tabela_antiga = s_log_tabela_antiga & formata_perc_desc(rs("perc_desconto")) & "=" & formata_perc_desc(rs("perc_comissao"))
				rs.MoveNext
				loop
			
			strSql = "DELETE FROM t_PERCENTUAL_COMISSAO_VENDEDOR WHERE (id_perfil = '" & c_id_perfil & "')"
			cn.Execute(strSql)
			if Err <> 0 then
				erro_fatal=True
				alerta = "FALHA AO EXCLUIR DADOS ANTERIORES DA TABELA t_PERCENTUAL_COMISSAO_VENDEDOR (" & Cstr(Err) & ": " & Err.Description & ")."
				end if
	
			s_log_tabela_nova = ""
			for i=Lbound(vArray) to Ubound(vArray)
				if s_log_tabela_nova <> "" then s_log_tabela_nova = s_log_tabela_nova & " "
				s_log_tabela_nova = s_log_tabela_nova & formata_perc_desc(vArray(i).c1) & "=" & formata_perc_desc(vArray(i).c2)
				strSql = "INSERT INTO t_PERCENTUAL_COMISSAO_VENDEDOR (" & _
							"id_perfil," & _
							"perc_desconto," & _
							"perc_comissao" & _
						") VALUES (" & _
							"'" & c_id_perfil & "'," & _
							bd_formata_numero(vArray(i).c1) & "," & _
							bd_formata_numero(vArray(i).c2) & _
							")"
				cn.Execute(strSql)
				next
	
		'	INFORMAÇÕES P/ O LOG
			if s_log_tabela_antiga <> s_log_tabela_nova then
				if operacao_selecionada = OP_INCLUI then
					s_log = "Cadastramento da tabela de comissão do vendedor para o perfil " & x_perfil_apelido(c_id_perfil) & ": " & s_log_tabela_nova
				else
					if s_log_tabela_antiga = "" then s_log_tabela_antiga = "(vazia)"
					if s_log_tabela_nova = "" then s_log_tabela_nova = "(vazia)"
					s_log_tabela_antiga = "Tabela antiga: " & s_log_tabela_antiga
					s_log_tabela_nova = "Tabela nova: " & s_log_tabela_nova
					s_log = "Edição da tabela de comissão do vendedor para o perfil " & x_perfil_apelido(c_id_perfil) & ": " & s_log_tabela_antiga & "; " & s_log_tabela_nova
					end if
				end if
	
			if s_log <> "" then grava_log usuario, "", "", "", OP_LOG_TABELA_COMISSAO_VENDEDOR, s_log
			
		'	~~~~~~~~~~~~~~
			cn.CommitTrans
		'	~~~~~~~~~~~~~~

			if Err=0 then 
				strUrlBotaoVoltar = "TabelaComissaoVendedorFiltro.asp" & "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))
			else
				alerta=Cstr(Err) & ": " & Err.Description
				end if
			
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

<!-- ************************************************************ -->
<!-- **********  PÁGINA PARA EXIBIR MENSAGENS DE ERRO  ********** -->
<!-- ************************************************************ -->
<body onload="bVOLTAR.focus();">
<center>
<br>
<!--  T E L A  -->
<p class="T">A V I S O</p>
<% 
	s = ""
	s_aux="'MtAviso'"
	if alerta <> "" then
		s = "<p style='margin:5px 2px 5px 2px;'>" & alerta & "</p>"
		s_aux="'MtAlerta'"
	else
		select case operacao_selecionada
			case OP_INCLUI
				s = "TABELA DE COMISSÃO CADASTRADA COM SUCESSO."
			case OP_CONSULTA, OP_ALTERA
				s = "TABELA DE COMISSÃO ALTERADA COM SUCESSO."
			case OP_EXCLUI
				s = "TABELA DE COMMISSÃO EXCLUÍDA COM SUCESSO."
			end select
		if s <> "" then s="<P style='margin:5px 2px 5px 2px;'>" & s & "</P>"
		end if
%>
<div class=<%=s_aux%> style="width:400px;font-weight:bold;" align="center"><%=s%></div>
<br><br>
<p class="TracoBottom"></p>
<table cellSpacing="0">
<tr>
	<td align="center"><a name="bVOLTAR" id="bVOLTAR" href="<%=strUrlBotaoVoltar%>"><img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
</tr>
</table>
</center>
</body>

</html>


<%
'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>