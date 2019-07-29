<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<% Response.Expires=-1 %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<%
'     ========================================================
'	  AjaxListaVendedoresLojaPesqBD.asp
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
	dim strSql, strResp, strNome, msg_erro
	
'	CONECTA COM O BANCO DE DADOS
	dim cn, rs
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

	dim s_loja, s_vendedor
	s_loja = UCase(Trim(Request("loja")))
    s_vendedor = UCase(Trim(Request("vendedor")))
	
	if (s_loja = "") Or (converte_numero(s_loja)=0) then
		Response.End
		end if

'   PESQUISA OS INDICADORES DO VENDEDOR INFORMADO
    if (s_vendedor <> "") then
        
        strSql = "SELECT * " & _
                        "FROM t_USUARIO " & _
                        "WHERE " & _
                            "(usuario = '" & s_vendedor & "') " & _
                            "ORDER BY " & _
                                "usuario"
    else

'	PESQUISA OS VENDEDORES DA LOJA INFORMADA
	strSql = _
			"SELECT " & _
				"*" & _
			" FROM t_USUARIO" & _
			" WHERE" & _
				"(" & _
					" (vendedor_loja = 1)" & _
					" AND" & _
						"(usuario IN " & _
							"(" & _
								"SELECT DISTINCT " & _
									"usuario" & _
								" FROM t_USUARIO_X_LOJA" & _
								" WHERE" & _
									" (CONVERT(smallint, loja) = " & s_loja & ")" & _
							")" & _
						")" & _
					")" & _
			" ORDER BY" & _
				" usuario"
    end if
	
'	EXECUTA A CONSULTA
	strResp = ""
	
	if rs.State <> 0 then rs.Close
	rs.open strSql, cn
'	ENCONTROU DADOS
	do while Not rs.Eof
		strNome = Trim("" & rs("nome"))
		if strNome = "" then strNome = "&nbsp;"
	'	Obs: text inside a CDATA section will be ignored by the parser
	'	Isso é importante porque caracteres como '&', '<', '>', etc geram erro no XML parser na leitura dos dados
		strResp = strResp & _
				  "<registro>" & _
					"<usuario>" & _
						"<![CDATA[" & _
						Trim("" & rs("usuario")) & _
						"]]>" & _
					"</usuario>" & _
					"<nome>" & _
						"<![CDATA[" & _
						strNome & _
						"]]>" & _
					"</nome>" & _
				  "</registro>"
		rs.MoveNext
		loop


'	ENCONTROU ALGUMA RESPOSTA
	if strResp <> "" then 
		Response.ContentType="text/xml"
		Response.Charset="ISO-8859-1"
		strResp = "<?xml version='1.0' encoding='ISO-8859-1'?>" & _
				  "<TabelaResultadoListaVendedoresLoja>" & _
				  strResp & _
				  "</TabelaResultadoListaVendedoresLoja>"
		end if
	
	Response.Write strResp

'	FECHA CONEXAO COM O BANCO DE DADOS
	if rs.State <> 0 then rs.Close
	set rs = nothing

	cn.Close
	set cn = nothing

%>
