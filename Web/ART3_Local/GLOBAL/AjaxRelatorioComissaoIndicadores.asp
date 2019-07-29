<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<% Response.Expires=-1 %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<%
'     =========================================================================
'	   A J A X R E L A T Ó R I O C O M I S S Ã O I N D I C A D O R E S . A S P
'     =========================================================================
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

	dim s, mes, ano, id_a, vendedor

	ano = Trim(Request("ano"))
	mes = Trim(Request("mes"))

	strSql = "SELECT " & _
                "t_COMISSAO_INDICADOR_N1.id, t_COMISSAO_INDICADOR_N1.dt_hr_cadastro, t_COMISSAO_INDICADOR_N2.vendedor, t_COMISSAO_INDICADOR_N2.id_comissao_indicador_n1 " & _
                "FROM t_COMISSAO_INDICADOR_N1 " & _
                "INNER JOIN t_COMISSAO_INDICADOR_N2 ON (t_COMISSAO_INDICADOR_N1.id=t_COMISSAO_INDICADOR_N2.id_comissao_indicador_n1) " & _
	            "WHERE t_COMISSAO_INDICADOR_N2.competencia_ano='" & ano & "' AND " & _
	            "t_COMISSAO_INDICADOR_N2.competencia_mes='" & mes & "' " & _
	            "ORDER BY dt_hr_cadastro"

'	EXECUTA A CONSULTA
	strResp = ""
	
	if rs.State <> 0 then rs.Close
	rs.open strSql, cn
'	ENCONTROU DADOS
	do while Not rs.Eof
        if id_a <> rs("id") then
        vendedor=""
        if strResp <> "" then strResp = strResp & "</option>"
            
		    strResp = strResp & _
				  "<option name='id' value='" & rs("id") & "'>"
             strResp = strResp & formata_data(rs("dt_hr_cadastro")) & " " & formata_hora_hhmm(rs("dt_hr_cadastro")) & "&nbsp;&nbsp;-&nbsp;&nbsp;"
    
        end if
       if vendedor <> "" then strResp = strResp & ", "
        strResp = strResp & rs("vendedor")   
        vendedor = vendedor & rs("vendedor")
        id_a = rs("id")

		rs.MoveNext
		loop
	
	Response.Write strResp & "</select>"

'	FECHA CONEXAO COM O BANCO DE DADOS
	if rs.State <> 0 then rs.Close
	set rs = nothing

	cn.Close
	set cn = nothing

%>
