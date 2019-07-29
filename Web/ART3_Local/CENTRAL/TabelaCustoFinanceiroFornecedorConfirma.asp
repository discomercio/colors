<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =============================================================
'	  TabelaCustoFinanceiroFornecedorConfirma.asp
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

	dim i, n, strSql, c_fabricante, operacao_selecionada

	strUrlBotaoVoltar = "javascript:history.back()"
	
	dim alerta
	alerta = ""
	
'	OBTÉM DADOS DIGITADOS NO FORMULÁRIO
	operacao_selecionada = Trim(Request("operacao_selecionada"))
	if operacao_selecionada = "" then
		alerta=texto_add_br(alerta)
		alerta=alerta & "Nenhuma operação foi selecionada."
		end if
	
	c_fabricante = Trim(Request.Form("c_fabricante"))
	if c_fabricante = "" then
		alerta=texto_add_br(alerta)
		alerta=alerta & "Fornecedor não foi informado."
		end if

	dim vArraySemEntrada
	redim vArraySemEntrada(0)
	set vArraySemEntrada(0) = New cl_DUAS_COLUNAS
	n = Request.Form("c_sem_entrada_qtde_parcelas").Count
	for i = 1 to n
		s=Trim(Request.Form("c_sem_entrada_coeficiente")(i))
		if converte_numero(s) <> 0 then
			if Trim(vArraySemEntrada(ubound(vArraySemEntrada)).c1) <> "" then
				redim preserve vArraySemEntrada(ubound(vArraySemEntrada)+1)
				set vArraySemEntrada(ubound(vArraySemEntrada)) = New cl_DUAS_COLUNAS
				end if
			with vArraySemEntrada(ubound(vArraySemEntrada))
				s = Trim(Request.Form("c_sem_entrada_qtde_parcelas")(i))
				.c1 = converte_numero(s)
				s = Trim(Request.Form("c_sem_entrada_coeficiente")(i))
				.c2 = converte_numero(s)
				end with
			end if
		next
	
	for i = Lbound(vArraySemEntrada) to (Ubound(vArraySemEntrada)-1)
		if vArraySemEntrada(i).c1 = vArraySemEntrada(i+1).c1 then
			alerta=texto_add_br(alerta)
			alerta=alerta & "A quantidade de parcelas 0+" & Cstr(vArraySemEntrada(i).c1) & " aparece em duplicidade."
		elseif vArraySemEntrada(i).c1 > vArraySemEntrada(i+1).c1 then
			alerta=texto_add_br(alerta)
			alerta=alerta & "As opções de parcelamento sem entrada não estão em ordem crescente."
			end if
		next

	for i = Lbound(vArraySemEntrada) to Ubound(vArraySemEntrada)
		if vArraySemEntrada(i).c2 = 0 then 
			alerta=texto_add_br(alerta)
			alerta=alerta & "Não foi informado o coeficiente para a opção 0+" & vArraySemEntrada(i).c1
			end if
		next

	dim vArrayComEntrada
	redim vArrayComEntrada(0)
	set vArrayComEntrada(0) = New cl_DUAS_COLUNAS
	n = Request.Form("c_com_entrada_qtde_parcelas").Count
	for i = 1 to n
		s=Trim(Request.Form("c_com_entrada_coeficiente")(i))
		if converte_numero(s) <> 0 then
			if Trim(vArrayComEntrada(ubound(vArrayComEntrada)).c1) <> "" then
				redim preserve vArrayComEntrada(ubound(vArrayComEntrada)+1)
				set vArrayComEntrada(ubound(vArrayComEntrada)) = New cl_DUAS_COLUNAS
				end if
			with vArrayComEntrada(ubound(vArrayComEntrada))
				s = Trim(Request.Form("c_com_entrada_qtde_parcelas")(i))
				.c1 = converte_numero(s)
				s = Trim(Request.Form("c_com_entrada_coeficiente")(i))
				.c2 = converte_numero(s)
				end with
			end if
		next
	
	for i = Lbound(vArrayComEntrada) to (Ubound(vArrayComEntrada)-1)
		if vArrayComEntrada(i).c1 = vArrayComEntrada(i+1).c1 then
			alerta=texto_add_br(alerta)
			alerta=alerta & "A quantidade de parcelas 1+" & Cstr(vArrayComEntrada(i).c1) & " aparece em duplicidade."
		elseif vArrayComEntrada(i).c1 > vArrayComEntrada(i+1).c1 then
			alerta=texto_add_br(alerta)
			alerta=alerta & "As opções de parcelamento com entrada não estão em ordem crescente."
			end if
		next

	for i = Lbound(vArrayComEntrada) to Ubound(vArrayComEntrada)
		if vArrayComEntrada(i).c2 = 0 then
			alerta=texto_add_br(alerta)
			alerta=alerta & "Não foi informado o coeficiente para a opção 1+" & vArrayComEntrada(i).c1
			end if
		next

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
			
		'	TABELA SEM ENTRADA
			strSql = "SELECT " & _
						"*" & _
					" FROM t_PERCENTUAL_CUSTO_FINANCEIRO_FORNECEDOR" & _
					" WHERE" & _
						"(fabricante = '" & c_fabricante & "')" & _
						" AND (tipo_parcelamento = '" & COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__SEM_ENTRADA & "')" & _
					" ORDER BY" & _
						" qtde_parcelas"
			set rs = cn.execute(strSql)
			do while Not rs.Eof
				if s_log_tabela_antiga <> "" then s_log_tabela_antiga = s_log_tabela_antiga & " "
				s_log_tabela_antiga = s_log_tabela_antiga & "0+" & Cstr(rs("qtde_parcelas")) & "=" & formata_coeficiente_custo_financ_fornecedor(rs("coeficiente"))
				rs.MoveNext
				loop

		'	TABELA COM ENTRADA
			strSql = "SELECT " & _
						"*" & _
					" FROM t_PERCENTUAL_CUSTO_FINANCEIRO_FORNECEDOR" & _
					" WHERE" & _
						"(fabricante = '" & c_fabricante & "')" & _
						" AND (tipo_parcelamento = '" & COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__COM_ENTRADA & "')" & _
					" ORDER BY" & _
						" qtde_parcelas"
			set rs = cn.execute(strSql)
			do while Not rs.Eof
				if s_log_tabela_antiga <> "" then s_log_tabela_antiga = s_log_tabela_antiga & " "
				s_log_tabela_antiga = s_log_tabela_antiga & "1+" & Cstr(rs("qtde_parcelas")) & "=" & formata_coeficiente_custo_financ_fornecedor(rs("coeficiente"))
				rs.MoveNext
				loop
			
		'	APAGA OS DADOS (SEM ENTRADA E COM ENTRADA)
			strSql = "DELETE FROM t_PERCENTUAL_CUSTO_FINANCEIRO_FORNECEDOR WHERE (fabricante = '" & c_fabricante & "')"
			cn.Execute(strSql)
			if Err <> 0 then
				erro_fatal=True
				alerta = "FALHA AO EXCLUIR DADOS ANTERIORES DA TABELA t_PERCENTUAL_CUSTO_FINANCEIRO_FORNECEDOR (" & Cstr(Err) & ": " & Err.Description & ")."
				end if
			
			if s_log_tabela_antiga <> "" then
				s_log = "Exclusão da tabela de custo financeiro para o fornecedor " & c_fabricante & " - " & x_fabricante(c_fabricante) & ": " & s_log_tabela_antiga
				end if
				
			if s_log <> "" then grava_log usuario, "", "", "", OP_LOG_TABELA_CUSTO_FINANCEIRO_FORNECEDOR, s_log
			
		'	~~~~~~~~~~~~~~
			cn.CommitTrans
		'	~~~~~~~~~~~~~~

			if Err=0 then 
				strUrlBotaoVoltar = "TabelaCustoFinanceiroFornecedorFiltro.asp" & "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))
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
			
		'	TABELA SEM ENTRADA
			strSql = "SELECT " & _
						"*" & _
					" FROM t_PERCENTUAL_CUSTO_FINANCEIRO_FORNECEDOR" & _
					" WHERE" & _
						"(fabricante = '" & c_fabricante & "')" & _
						" AND (tipo_parcelamento = '" & COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__SEM_ENTRADA & "')" & _
					" ORDER BY" & _
						" qtde_parcelas"
			set rs = cn.execute(strSql)
			do while Not rs.Eof
				if s_log_tabela_antiga <> "" then s_log_tabela_antiga = s_log_tabela_antiga & " "
				s_log_tabela_antiga = s_log_tabela_antiga & "0+" & Cstr(rs("qtde_parcelas")) & "=" & formata_coeficiente_custo_financ_fornecedor(rs("coeficiente"))
				rs.MoveNext
				loop
			
		'	TABELA COM ENTRADA
			strSql = "SELECT " & _
						"*" & _
					" FROM t_PERCENTUAL_CUSTO_FINANCEIRO_FORNECEDOR" & _
					" WHERE" & _
						"(fabricante = '" & c_fabricante & "')" & _
						" AND (tipo_parcelamento = '" & COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__COM_ENTRADA & "')" & _
					" ORDER BY" & _
						" qtde_parcelas"
			set rs = cn.execute(strSql)
			do while Not rs.Eof
				if s_log_tabela_antiga <> "" then s_log_tabela_antiga = s_log_tabela_antiga & " "
				s_log_tabela_antiga = s_log_tabela_antiga & "1+" & Cstr(rs("qtde_parcelas")) & "=" & formata_coeficiente_custo_financ_fornecedor(rs("coeficiente"))
				rs.MoveNext
				loop
			
		'	APAGA OS DADOS (SEM ENTRADA E COM ENTRADA)
			strSql = "DELETE FROM t_PERCENTUAL_CUSTO_FINANCEIRO_FORNECEDOR WHERE (fabricante = '" & c_fabricante & "')"
			cn.Execute(strSql)
			if Err <> 0 then
				erro_fatal=True
				alerta = "FALHA AO EXCLUIR DADOS ANTERIORES DA TABELA t_PERCENTUAL_CUSTO_FINANCEIRO_FORNECEDOR (" & Cstr(Err) & ": " & Err.Description & ")."
				end if
	
			s_log_tabela_nova = ""
			
		'	TABELA SEM ENTRADA
			for i=Lbound(vArraySemEntrada) to Ubound(vArraySemEntrada)
				if s_log_tabela_nova <> "" then s_log_tabela_nova = s_log_tabela_nova & " "
				s_log_tabela_nova = s_log_tabela_nova & "0+" & Cstr(vArraySemEntrada(i).c1) & "=" & formata_coeficiente_custo_financ_fornecedor(vArraySemEntrada(i).c2)
				strSql = "INSERT INTO t_PERCENTUAL_CUSTO_FINANCEIRO_FORNECEDOR (" & _
							"fabricante," & _
							"tipo_parcelamento," & _
							"qtde_parcelas," & _
							"coeficiente" & _
						") VALUES (" & _
							"'" & c_fabricante & "'," & _
							"'" & COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__SEM_ENTRADA & "'," & _
							bd_formata_numero(vArraySemEntrada(i).c1) & "," & _
							bd_formata_numero(vArraySemEntrada(i).c2) & _
							")"
				cn.Execute(strSql)
				next

		'	TABELA COM ENTRADA
			for i=Lbound(vArrayComEntrada) to Ubound(vArrayComEntrada)
				if s_log_tabela_nova <> "" then s_log_tabela_nova = s_log_tabela_nova & " "
				s_log_tabela_nova = s_log_tabela_nova & "1+" & Cstr(vArrayComEntrada(i).c1) & "=" & formata_coeficiente_custo_financ_fornecedor(vArrayComEntrada(i).c2)
				strSql = "INSERT INTO t_PERCENTUAL_CUSTO_FINANCEIRO_FORNECEDOR (" & _
							"fabricante," & _
							"tipo_parcelamento," & _
							"qtde_parcelas," & _
							"coeficiente" & _
						") VALUES (" & _
							"'" & c_fabricante & "'," & _
							"'" & COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__COM_ENTRADA & "'," & _
							bd_formata_numero(vArrayComEntrada(i).c1) & "," & _
							bd_formata_numero(vArrayComEntrada(i).c2) & _
							")"
				cn.Execute(strSql)
				next
	
		'	INFORMAÇÕES P/ O LOG
			if s_log_tabela_antiga <> s_log_tabela_nova then
				if operacao_selecionada = OP_INCLUI then
					s_log = "Cadastramento da tabela de custo financeiro para o fornecedor " & c_fabricante & " - " & x_fabricante(c_fabricante) & ": " & s_log_tabela_nova
				else
					if s_log_tabela_antiga = "" then s_log_tabela_antiga = "(vazia)"
					if s_log_tabela_nova = "" then s_log_tabela_nova = "(vazia)"
					s_log_tabela_antiga = "Tabela antiga: " & s_log_tabela_antiga
					s_log_tabela_nova = "Tabela nova: " & s_log_tabela_nova
					s_log = "Edição da tabela de custo financeiro para o fornecedor " & c_fabricante & " - " & x_fabricante(c_fabricante) & ": " & s_log_tabela_antiga & "; " & s_log_tabela_nova
					end if
				end if
	
			if s_log <> "" then grava_log usuario, "", "", "", OP_LOG_TABELA_CUSTO_FINANCEIRO_FORNECEDOR, s_log
			
		'	~~~~~~~~~~~~~~
			cn.CommitTrans
		'	~~~~~~~~~~~~~~

			if Err=0 then 
				strUrlBotaoVoltar = "TabelaCustoFinanceiroFornecedorFiltro.asp" & "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))
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
		s = "<P style='margin:5px 2px 5px 2px;'>" & alerta & "</P>"
		s_aux="'MtAlerta'"
	else
		select case operacao_selecionada
			case OP_INCLUI
				s = "TABELA DE CUSTO FINANCEIRO CADASTRADA COM SUCESSO."
			case OP_CONSULTA, OP_ALTERA
				s = "TABELA DE CUSTO FINANCEIRO ALTERADA COM SUCESSO."
			case OP_EXCLUI
				s = "TABELA DE CUSTO FINANCEIRO EXCLUÍDA COM SUCESSO."
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