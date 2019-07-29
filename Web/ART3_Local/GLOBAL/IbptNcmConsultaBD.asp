<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<% Response.Expires=-1 %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<%
'     =========================================
'	  IbptNcmConsultaBD.asp
'     =========================================
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

	dim i, s_ncm, s_item, v_aux, v_ncm
	s_ncm = Trim(Request("ncm"))
	s_ncm = Replace(s_ncm, "|", ",")
	s_ncm = Replace(s_ncm, ";", ",")
	s_ncm = Replace(s_ncm, " ", ",")
	v_aux = Split(s_ncm, ",")
	redim v_ncm(0)
	for i=Lbound(v_aux) to Ubound(v_aux)
		if Trim(v_aux(i)) <> "" then
			if Trim(v_ncm(Ubound(v_ncm))) <> "" then
				redim preserve v_ncm(Ubound(v_ncm)+1)
				end if
			v_ncm(Ubound(v_ncm)) = Trim(v_aux(i))
			end if
		next
	
'	EXECUTA A CONSULTA
	strResp = ""
	for i=Lbound(v_ncm) to Ubound(v_ncm)
		if Trim(v_ncm(i)) <> "" then
			s_item = ""
			strSql = "SELECT " & _
						"*" & _
					" FROM t_IBPT" & _
					" WHERE" & _
						" (codigo = '" & Trim(v_ncm(i)) & "')" & _
						" AND (tabela = '0')" & _
					" ORDER BY" & _
						" codigo," & _
						" ex"
			if rs.State <> 0 then rs.Close
			rs.open strSql, cn
			if rs.Eof then
				s_item = _
						"{" & _
						"'ncm':'" & Trim(v_ncm(i)) & "'" & _
						" , " & _
						"'cadastrado':false" & _
						" , " & _
						"'ex':''" & _
						" , " & _
						"'percAliqNac':0.0" & _
						" , " & _
						"'percAliqImp':0.0" & _
						"}"
			else
				do while Not rs.Eof
					s_item = s_item & _
							"{" & _
							"'ncm':'" & Trim(v_ncm(i)) & "'" & _
							" , " & _
							"'cadastrado':true" & _
							" , " & _
							"'ex':'" & Trim("" & rs("ex")) & "'" & _
							" , " & _
							"'percAliqNac':" & js_formata_numero(rs("percAliqNac")) & _
							" , " & _
							"'percAliqImp':" & js_formata_numero(rs("percAliqImp")) & _
							"}"
					rs.MoveNext
					if Not rs.Eof then s_item = s_item & ","
					loop
				end if
			
			if strResp <> "" then strResp = strResp & ","
			strResp = strResp & s_item
			end if
		next


	strResp = "{'resposta': [" & strResp & "]}"
	strResp = Replace(strResp, "'", chr(34))
	
'	RETORNA RESPOSTA
	Response.Write strResp

'	FECHA CONEXAO COM O BANCO DE DADOS
	if rs.State <> 0 then rs.Close
	set rs = nothing

	cn.Close
	set cn = nothing

%>
