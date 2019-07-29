<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<% Response.Expires=-1 %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<%
'     ========================================================
'	  JsonConsultaSenhaDescontoBD.asp
'     ========================================================
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
	dim strSql, strResp, msg_erro
	
'	CONECTA COM O BANCO DE DADOS
	dim cn, rs
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

	dim id_cliente
	id_cliente = Trim(Request("id_cliente"))
	
	dim s_loja
	s_loja = UCase(Trim(Request("loja")))
	
	if (s_loja = "") Or (converte_numero(s_loja)=0) then
		Response.End
		end if
	
'	PESQUISA OS INDICADORES DA LOJA INFORMADA
	strSql = _
			"SELECT " & _
				"*" & _
			" FROM t_DESCONTO" & _
			" WHERE" & _
				" (usado_status=0)" & _
				" AND (cancelado_status=0)" & _
				" AND (id_cliente='" & id_cliente & "')" & _
				" AND (loja='" & s_loja & "')" & _
				" AND (data >= " & bd_formata_data_hora(Now-converte_min_to_dec(TIMEOUT_DESCONTO_EM_MIN)) & ")" & _
			" ORDER BY" & _
				" fabricante," & _
				" produto," & _
				" data DESC"
'	EXECUTA A CONSULTA
	if rs.State <> 0 then rs.Close
	rs.open strSql, cn
'	ENCONTROU DADOS
	strResp = ""
	do while Not rs.Eof
		if strResp <> "" then strResp = strResp & ", "
		strResp = strResp & _
					"{" & _
					"""id"":" & """" & Trim("" & rs("id")) & """," & _
					"""fabricante"":" & """" & Trim("" & rs("fabricante")) & """," & _
					"""produto"":" & """" & Trim("" & rs("produto")) & """," & _
					"""desc_max"":" & js_formata_numero(rs("desc_max")) & "," & _
					"""data"":" & """" & formata_data_hora(rs("data")) & """," & _
					"""id_cliente"":" & """" & Trim("" & rs("id_cliente")) & """," & _
					"""cnpj_cpf"":" & """" & Trim("" & rs("cnpj_cpf")) & """," & _
					"""loja"":" & """" & Trim("" & rs("loja")) & """," & _
					"""autorizador"":" & """" & Trim("" & rs("autorizador")) & """," & _
					"""supervisor_autorizador"":" & """" & Trim("" & rs("supervisor_autorizador")) & """" & _
					"}"
		rs.MoveNext
		loop

	strResp = "{" & _
					"""item"":[" & _
					strResp & _
					"]" & _
				"}"

	Response.Write strResp

'	FECHA CONEXAO COM O BANCO DE DADOS
	if rs.State <> 0 then rs.Close
	set rs = nothing

	cn.Close
	set cn = nothing
%>
