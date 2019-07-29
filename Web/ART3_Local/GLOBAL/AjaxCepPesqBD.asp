<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<% Response.Expires=-1 %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<%
'     =========================================
'	  A J A X C E P P E S Q B D . A S P
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
	
	Const OPCAO_PESQUISA_POR_CEP = "POR_CEP"
	Const OPCAO_PESQUISA_POR_ENDERECO = "POR_END"
	
'	OBTEM O ID
	dim strSql, strResp, strEspaco, strLogradouro, strBairro, msg_erro
	
'	CONECTA COM O BANCO DE DADOS
	dim cn, rs
	If Not bdd_cep_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

	dim s_cep, s_endereco, s_uf, s_localidade, s_opcao_pesquisa_por
	s_cep = retorna_so_digitos(Trim(Request("cep")))
	s_endereco = Trim(Request("endereco"))
	s_uf = UCase(Trim(Request("uf")))
	s_localidade = UCase(Trim(Request("localidade")))
	s_opcao_pesquisa_por=Trim(Request("opcao"))
	
	if s_opcao_pesquisa_por = OPCAO_PESQUISA_POR_CEP then
	'	A PESQUISA É POR CEP
		if (len(s_cep) <> 5) And (len(s_cep) <> 8) then
			Response.End
			end if
	elseif s_opcao_pesquisa_por = OPCAO_PESQUISA_POR_ENDERECO then
	'	A PESQUISA É POR ENDEREÇO (UF OBRIGATÓRIO NESTE CASO)
		if (s_uf = "") Or (Not uf_ok(s_uf)) Or (s_localidade = "") then
			Response.End
			end if
	else
	'	PARÂMETROS INVÁLIDOS
		Response.End
		end if
	
	if s_opcao_pesquisa_por = OPCAO_PESQUISA_POR_CEP then
	'	PESQUISA POR CEP?
		strSql = "SELECT" & _
					" 'LOGRADOURO' AS tabela_origem," & _
					" Logr.CEP_DIG AS cep," & _
					" Logr.UFE_SG AS uf," & _
					" Loc.LOC_NOSUB AS localidade," & _
					" Bai.BAI_NO AS bairro_extenso," & _
					" Bai.BAI_NO_ABREV AS bairro_abreviado," & _
					" Logr.LOG_TIPO_LOGRADOURO AS logradouro_tipo," & _
					" Logr.LOG_NO AS logradouro_nome," & _
					" Logr.LOG_COMPLEMENTO AS logradouro_complemento" & _
				" FROM LOG_LOGRADOURO Logr" & _
					" LEFT JOIN LOG_BAIRRO Bai ON (Logr.BAI_NU_SEQUENCIAL_INI = Bai.BAI_NU_SEQUENCIAL)" & _
					" LEFT JOIN LOG_LOCALIDADE Loc ON (Logr.LOC_NU_SEQUENCIAL = Loc.LOC_NU_SEQUENCIAL)" & _
				" WHERE"
				
		if len(s_cep) = 5 then
			strSql = strSql & _
						" (Logr.CEP_DIG LIKE '" & s_cep & BD_CURINGA_TODOS & "')"
		else
			strSql = strSql & _
						" (Logr.CEP_DIG = '" & s_cep & "')"
			end if
		
		strSql = strSql & _
				" UNION " & _
				"SELECT" & _
					" 'LOCALIDADE' AS tabela_origem," & _
					" CEP_DIG AS cep," & _
					" UFE_SG AS uf," & _
					" LOC_NOSUB AS localidade," & _
					" '' AS bairro_extenso," & _
					" '' AS bairro_abreviado," & _
					" '' AS logradouro_tipo," & _
					" '' AS logradouro_nome," & _
					" '' AS logradouro_complemento" & _
					" FROM LOG_LOCALIDADE " & _
					" WHERE"
					
		if len(s_cep) = 5 then
			strSql = strSql & _
						" (CEP_DIG LIKE '" & s_cep & BD_CURINGA_TODOS & "')"
		else
			strSql = strSql & _
						" (CEP_DIG = '" & s_cep & "')"
			end if
		
	'	CONSULTA DADOS DA TABELA ANTIGA, POIS ELA É MANTIDA P/ MANTER FUNCIONANDO O CADASTRAMENTO MANUAL DE CEP'S
		strSql = strSql & _
				" UNION " & _
				"SELECT" & _
					" 'LOGRADOURO' AS tabela_origem," & _
					" cep8_log" & SQL_COLLATE_CASE_ACCENT & " AS cep," & _
					" uf_log" & SQL_COLLATE_CASE_ACCENT & " AS uf," & _
					" nome_local" & SQL_COLLATE_CASE_ACCENT & " AS localidade," & _
					" extenso_bai" & SQL_COLLATE_CASE_ACCENT & " AS bairro_extenso," & _
					" abrev_bai" & SQL_COLLATE_CASE_ACCENT & " AS bairro_abreviado," & _
					" abrev_tipo" & SQL_COLLATE_CASE_ACCENT & " AS logradouro_tipo," & _
					" nome_log" & SQL_COLLATE_CASE_ACCENT & " AS logradouro_nome," & _
					" comple_log" & SQL_COLLATE_CASE_ACCENT & " AS logradouro_complemento" & _
				" FROM t_CEP_LOGRADOURO " & _
				" WHERE"
				
		if len(s_cep) = 5 then
			strSql = strSql & _
						" (cep8_log LIKE '" & s_cep & BD_CURINGA_TODOS & "')"
		else
			strSql = strSql & _
						" (cep8_log = '" & s_cep & "')"
			end if

		strSql = strSql & _
				" ORDER BY cep"
	
	else
	'	PESQUISA POR ENDEREÇO
		strSql = "SELECT TOP 300" & _
					" 'LOGRADOURO' AS tabela_origem," & _
					" Logr.CEP_DIG AS cep," & _
					" Logr.UFE_SG AS uf," & _
					" Loc.LOC_NOSUB AS localidade," & _
					" Bai.BAI_NO AS bairro_extenso," & _
					" Bai.BAI_NO_ABREV AS bairro_abreviado," & _
					" Logr.LOG_TIPO_LOGRADOURO AS logradouro_tipo," & _
					" Logr.LOG_NO AS logradouro_nome," & _
					" Logr.LOG_COMPLEMENTO AS logradouro_complemento" & _
				" FROM LOG_LOGRADOURO Logr" & _
					" LEFT JOIN LOG_BAIRRO Bai ON (Logr.BAI_NU_SEQUENCIAL_INI = Bai.BAI_NU_SEQUENCIAL)" & _
					" LEFT JOIN LOG_LOCALIDADE Loc ON (Logr.LOC_NU_SEQUENCIAL = Loc.LOC_NU_SEQUENCIAL)" & _
				" WHERE" & _
					"(Logr.UFE_SG = '" & s_uf & "')" & _
					" AND " & _
					"(Loc.LOC_NOSUB = '" & QuotedStr(s_localidade) & "'" & SQL_COLLATE_CASE_ACCENT & ")"
			
		if s_endereco <> "" then
			strSql = strSql & _
						" AND " & _
						"(Logr.LOG_NO LIKE '" & BD_CURINGA_TODOS & QuotedStr(s_endereco) & BD_CURINGA_TODOS & "'" & SQL_COLLATE_CASE_ACCENT & ")"
			end if
		
		strSql = strSql & _
				" UNION " & _
				"SELECT" & _
					" 'LOCALIDADE' AS tabela_origem," & _
					" CEP_DIG AS cep," & _
					" UFE_SG AS uf," & _
					" LOC_NOSUB AS localidade," & _
					" '' AS bairro_extenso," & _
					" '' AS bairro_abreviado," & _
					" '' AS logradouro_tipo," & _
					" '' AS logradouro_nome," & _
					" '' AS logradouro_complemento" & _
				" FROM LOG_LOCALIDADE " & _
				" WHERE" & _
					" (UFE_SG = '" & s_uf & "')" & _
					" AND (LOC_NOSUB = '" & QuotedStr(s_localidade) & "'" & SQL_COLLATE_CASE_ACCENT & ")" & _
					" AND (LEN(Coalesce(CEP_DIG,'')) > 0)"

	'	CONSULTA DADOS DA TABELA ANTIGA, POIS ELA É MANTIDA P/ MANTER FUNCIONANDO O CADASTRAMENTO MANUAL DE CEP'S
		strSql = strSql & _
				" UNION " & _
				"SELECT TOP 300" & _
					" 'LOGRADOURO' AS tabela_origem," & _
					" cep8_log" & SQL_COLLATE_CASE_ACCENT & " AS cep," & _
					" uf_log" & SQL_COLLATE_CASE_ACCENT & " AS uf," & _
					" nome_local" & SQL_COLLATE_CASE_ACCENT & " AS localidade," & _
					" extenso_bai" & SQL_COLLATE_CASE_ACCENT & " AS bairro_extenso," & _
					" abrev_bai" & SQL_COLLATE_CASE_ACCENT & " AS bairro_abreviado," & _
					" abrev_tipo" & SQL_COLLATE_CASE_ACCENT & " AS logradouro_tipo," & _
					" nome_log" & SQL_COLLATE_CASE_ACCENT & " AS logradouro_nome," & _
					" comple_log" & SQL_COLLATE_CASE_ACCENT & " AS logradouro_complemento" & _
				" FROM t_CEP_LOGRADOURO " & _
				" WHERE" & _
					"(uf_log = '" & s_uf & "')" & _
					" AND " & _
					"(nome_local = '" & QuotedStr(s_localidade) & "'" & SQL_COLLATE_CASE_ACCENT & ")"

		if s_endereco <> "" then
			strSql = strSql & _
						" AND " & _
						"(nome_log LIKE '" & BD_CURINGA_TODOS & QuotedStr(s_endereco) & BD_CURINGA_TODOS & "'" & SQL_COLLATE_CASE_ACCENT & ")"
			end if

		strSql = strSql & _
				" ORDER BY uf, localidade, bairro_extenso, cep, logradouro_nome"
		end if
	
	
'	EXECUTA A CONSULTA
	strResp = ""
	
	if rs.State <> 0 then rs.Close
	rs.open strSql, cn
'	ENCONTROU DADOS NA TABELA DE LOGRADOURO
	do while Not rs.Eof
	'	OBSERVAÇÃO: OS CAMPOS DO RESULTADO DEVEM SER IGUAIS, TANTO NA CONSULTA 
	'	DA TABELA DE LOGRADOURO QUANTO NA DE LOCALIDADE
			
	'	LOGRADOURO
		if (Trim("" & rs("logradouro_tipo")) <> "") And (Trim("" & rs("logradouro_nome")) <> "") then 
			strEspaco = " "
		else
			strEspaco = ""
			end if
		strLogradouro =  Trim("" & rs("logradouro_tipo")) & strEspaco & Trim("" & rs("logradouro_nome"))
		
	'	BAIRRO
		strBairro=Trim("" & rs("bairro_extenso"))
		if strBairro="" then strBairro=Trim("" & rs("bairro_abreviado"))
		strBairro = substitui_caracteres(strBairro, "&", "&amp;")
		
		strResp = strResp & _
				  "<registro>" & _
					"<tabela>" & _
						Trim("" & rs("tabela_origem")) & _
					"</tabela>" & _
					"<cep>" & _
						cep_formata(Trim("" & rs("cep"))) & _
					"</cep>" & _
					"<logradouro_nome>" & _
						substitui_caracteres(iniciais_em_maiusculas(strLogradouro), "&", "&amp;") & _
					"</logradouro_nome>" & _
					"<logradouro_complemento>" & _
						substitui_caracteres(iniciais_em_maiusculas(Trim("" & rs("logradouro_complemento"))), "&", "&amp;") & _
					"</logradouro_complemento>" & _
					"<bairro>" & _
						iniciais_em_maiusculas(strBairro) & _
					"</bairro>" & _
					"<localidade>" & _
						substitui_caracteres(iniciais_em_maiusculas(Trim("" & rs("localidade"))), "&", "&amp;") & _
					"</localidade>" & _
					"<uf>" & _
						Trim("" & rs("uf")) & _
					"</uf>" & _
				  "</registro>"
		rs.MoveNext
		loop


'	ENCONTROU ALGUMA RESPOSTA (LOGRADOURO OU LOCALIDADE)?
	if strResp <> "" then 
		Response.ContentType="text/xml"
		strResp = "<?xml version='1.0' encoding='ISO-8859-1'?>" & _
				  "<TabelaResultadoCep>" & _
				  strResp & _
				  "</TabelaResultadoCep>"
		end if
		
	Response.Write strResp

'	FECHA CONEXAO COM O BANCO DE DADOS
	if rs.State <> 0 then rs.Close
	set rs = nothing

	cn.Close
	set cn = nothing

%>
