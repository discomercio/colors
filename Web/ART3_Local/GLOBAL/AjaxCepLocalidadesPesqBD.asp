<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<% Response.Expires=-1 %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<%
'     ========================================================
'	  A J A X C E P L O C A L I D A D E S P E S Q B D . A S P
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
	If Not bdd_cep_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

	dim s, s_uf, s_retira_acentuacao, s_collate
	s_uf = UCase(Trim(Request("uf")))
	s_retira_acentuacao = Ucase(Trim(Request("retira_acentuacao")))

	
	if (s_uf = "") Or (Not uf_ok(s_uf)) then
		Response.End
		end if
	
	s_collate = ""
	if s_retira_acentuacao = "S" then s_collate = " COLLATE Latin1_General_CI_AI"

'	PESQUISA LOCALIDADES DA UF INFORMADA
	strSql = _
			"SELECT DISTINCT" & _
				" LOC_NOSUB" & s_collate & " AS localidade" & _
			" FROM LOG_LOCALIDADE" & _
			" WHERE" & _
				" (UFE_SG = '" & s_uf & "')" & _
			" ORDER BY" & _
				" LOC_NOSUB" & s_collate

'	EXECUTA A CONSULTA
	strResp = ""
	
	if rs.State <> 0 then rs.Close
	rs.open strSql, cn
'	ENCONTROU DADOS
	do while Not rs.Eof
		s = Trim("" & rs("localidade"))
		if s_retira_acentuacao = "S" then s = retira_acentuacao(s)
		strResp = strResp & _
				  "<registro>" & _
					"<localidade>" & _
						s & _
					"</localidade>" & _
				  "</registro>"
		rs.MoveNext
		loop


'	ENCONTROU ALGUMA RESPOSTA
	if strResp <> "" then 
		Response.ContentType="text/xml"
		strResp = "<?xml version='1.0' encoding='ISO-8859-1'?>" & _
				  "<TabelaResultadoLocalidades>" & _
				  strResp & _
				  "</TabelaResultadoLocalidades>"
		end if
	
	Response.Write strResp

'	FECHA CONEXAO COM O BANCO DE DADOS
	if rs.State <> 0 then rs.Close
	set rs = nothing

	cn.Close
	set cn = nothing

%>
