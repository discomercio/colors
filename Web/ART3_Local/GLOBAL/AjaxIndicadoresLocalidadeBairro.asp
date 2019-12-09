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
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

	dim s, s_cidade, s_retira_acentuacao, s_collate
	dim s_loja, s_vendedor
	
	s_loja = UCase(Trim(Request("loja")))
	s_cidade = Trim(Request("cidade"))
	s_retira_acentuacao = Trim(Request("retira_acentuacao"))
    s_vendedor = UCase(Trim(Request("vendedor")))
	
	s_collate = ""
	if s_retira_acentuacao = "S" then s_collate = " COLLATE Latin1_General_CI_AI"

'   PESQUISA OS INDICADORES DO VENDEDOR INFORMADO
    if (s_vendedor <> "") then
        
        strSql = "SELECT DISTINCT " & _
                    "bairro " & _
                        "FROM t_ORCAMENTISTA_E_INDICADOR " & _
                        "WHERE " & _
                            "(vendedor = '" & s_vendedor & "') " & _
                            "AND (cidade = '" & s_cidade & "' COLLATE Latin1_General_CI_AI)" & _
                            "ORDER BY " & _
                                "bairro"
    else
	
'	PESQUISA LOCALIDADES DA UF INFORMADA
	strSql = _
			"SELECT DISTINCT" & _
				" bairro" & s_collate & " AS bairro" & _
			" FROM t_ORCAMENTISTA_E_INDICADOR" & _
			" WHERE" & _
				" (cidade = '" & s_cidade & "' COLLATE Latin1_General_CI_AI)"
    if s_loja <> "" then
        strSql = strSql & _
				" AND (loja = '" & s_loja & "')"
        end if

    strSql = strSql & _
			" ORDER BY" & _
				" bairro" & s_collate
    end if

'	EXECUTA A CONSULTA
	strResp = ""
	
	if rs.State <> 0 then rs.Close
	rs.open strSql, cn
'	ENCONTROU DADOS
	do while Not rs.Eof
		s = Trim("" & rs("bairro"))
		if s_retira_acentuacao = "S" then s = retira_acentuacao(s)
		strResp = strResp & _
				  "<registro>" & _
					"<bairro>" & _
						s & _
					"</bairro>" & _
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
