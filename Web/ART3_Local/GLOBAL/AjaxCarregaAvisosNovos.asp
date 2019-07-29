<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<% Response.Expires=-1 %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<%
'     ========================================================
'	  AjaxCarregaAvisosNovos.asp
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
	dim strSql, strResp, strRazaoSocialNome, msg_erro
	
'	CONECTA COM O BANCO DE DADOS
	dim cn, rs, r, s
    dim id_lista_exibidos_atualizar, id_lista_exibidos_novo, sql_insere
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	If Not cria_recordset_otimista(r, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)


	dim s_loja, s_usuario, id_exibido, id_lido, data_exibido
	s_loja = UCase(Trim(Request("loja")))
    s_usuario = Trim(Request("usuario"))

        strSql = "SELECT t_AVISO.*, t_AVISO_EXIBIDO.id AS id_aviso_exibido, t_AVISO_EXIBIDO.dt_hr_ult_exibicao AS data_exibicao, t_AVISO_LIDO.id AS id_aviso_lido" & _
                    " FROM t_AVISO" & _
                    " LEFT JOIN t_AVISO_EXIBIDO ON (t_AVISO.id = t_AVISO_EXIBIDO.id) AND (t_AVISO_EXIBIDO.usuario = '" & s_usuario & "') " & _
                    " LEFT JOIN t_AVISO_LIDO ON (t_AVISO.id = t_AVISO_LIDO.id) AND (t_AVISO_LIDO.usuario = '" & s_usuario & "') " & _
                    " WHERE("

        if s_loja <> "" then strSql = strSql & "(destinatario = '" & s_loja & "') OR "

        strSql = strSql & "(destinatario = '') OR (destinatario IS NULL))" & _
                    " ORDER BY dt_ult_atualizacao DESC"
  

'	EXECUTA A CONSULTA
	strResp = ""
	
	if rs.State <> 0 then rs.Close
	rs.open strSql, cn
    id_lista_exibidos_atualizar = ""
    id_lista_exibidos_novo = ""
    sql_insere = ""

'	ENCONTROU DADOS
	do while Not rs.Eof
        id_exibido = Trim("" & rs("id_aviso_exibido"))
        id_lido = Trim("" & rs("id_aviso_lido"))
        data_exibido = rs("data_exibicao")

        if (id_exibido = "" And id_lido = "") Or (id_lido = "" And DateDiff("n", data_exibido, Now) > 30) then
            
		    strResp = strResp & _
				  "<registro>" & _
                    "<id>" & _
						Trim("" & rs("id")) & _
					"</id>" & _
					"<usuario>" & _
						Trim("" & rs("usuario")) & _
					"</usuario>" & _
                    "<datahora>" & _
						Trim("" & rs("dt_ult_atualizacao")) & _
					"</datahora>" & _
					"<mensagem>" & _
						"<![CDATA[" & _
						Trim("" & rs("mensagem")) & _
						"]]>" & _
					"</mensagem>" & _
				  "</registro>"
	    
        end if

        if id_exibido <> "" then 
                if id_lista_exibidos_atualizar <> "" then id_lista_exibidos_atualizar = id_lista_exibidos_atualizar & "', '"
                id_lista_exibidos_atualizar = id_lista_exibidos_atualizar & id_exibido
        else
            if id_lista_exibidos_novo <> "" then id_lista_exibidos_novo = id_lista_exibidos_novo & ", "
            id_lista_exibidos_novo = id_lista_exibidos_novo & "('" & rs("id") & "', '" & s_usuario & "', GETDATE())"
        end if

		rs.MoveNext
	loop

        '   MARCAR COMO EXIBIDO
        if strResp <> "" And id_lista_exibidos_atualizar <> "" then 
            s = "UPDATE t_AVISO_EXIBIDO SET dt_hr_ult_exibicao=GETDATE() WHERE id IN ('" & id_lista_exibidos_atualizar & "') AND (usuario='" & s_usuario & "')"
	        set r = cn.Execute(s)
        end if
        if strResp <> "" And id_lista_exibidos_novo <> "" then
            s = "INSERT INTO t_AVISO_EXIBIDO (id, usuario, dt_hr_ult_exibicao) VALUES " & id_lista_exibidos_novo
            set r = cn.Execute(s)

        end if

'	ENCONTROU ALGUMA RESPOSTA
	if strResp <> "" then 
		Response.ContentType="text/xml"
		Response.Charset="ISO-8859-1"
		strResp = "<?xml version='1.0' encoding='ISO-8859-1'?>" & _
				  "<TabelaQuadroDeAvisos>" & _
				  strResp & _
				  "</TabelaQuadroDeAvisos>"
		end if
	
	Response.Write strResp

'	FECHA CONEXAO COM O BANCO DE DADOS
	if rs.State <> 0 then rs.Close
	set rs = nothing

    if r.State <> 0 then r.Close
	set r = nothing

	cn.Close
	set cn = nothing

%>
