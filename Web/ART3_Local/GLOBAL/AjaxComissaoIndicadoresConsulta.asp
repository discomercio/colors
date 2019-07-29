<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<% Response.Expires=-1 %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<%
'     =========================================================================
'	  A J A X I N D I C A D O R E S B L O C O N O T A S C O N S U L T A . A S P
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

	dim s, apelido, mes, ano

	apelido = Trim(Request("apelido"))
	ano = Trim(Request("ano"))
	mes = Trim(Request("mes"))

	
'	PESQUISA BLOCOS DE NOTAS DO APELIDO, MÊS E ANO INFORMADOS
	strSql = "SELECT * FROM t_ORCAMENTISTA_E_INDICADOR_BLOCO_NOTAS " & _
	            "WHERE apelido='" & apelido & "' AND " & _
	            "YEAR(dt_cadastro)='" & ano & "' AND " & _
	            "MONTH(dt_cadastro)='" & mes & "' AND " & _
	            "anulado_status=0 " & _
	            "ORDER BY dt_hr_cadastro"

'	EXECUTA A CONSULTA
	strResp = ""
	
	if rs.State <> 0 then rs.Close
	rs.open strSql, cn
'	ENCONTROU DADOS
	do while Not rs.Eof
		strResp = strResp & _
				  "<tr>" & _
				  "<td class='C ME MD MB' style='width:60px;' align='center' valign='top'>" & formata_data_hora(rs("dt_hr_cadastro")) & "</td>" & _
				  "<td class='C MD MB' style='width:80px;' align='center' valign='top'>" & rs("usuario") 
				  if (rs("loja")) <> "" then
				    strResp = strResp & "(Loja&nbsp;" & rs("loja") & ")</td>"
				  else
				    strResp = strResp & "</td>"
				  end if
				  strResp = strResp & "<td class='C MD MB' align='left' style='width:509px' valign='top'>" & substitui_caracteres(rs("mensagem"), chr(13), "<br>") & "</td>" & _
				  "</tr>"
		rs.MoveNext
		loop
		
	if (strResp = "") then
	    strResp = "<tr>" & _
	              "<td class='C ME MB MD' style='width:649px;color:#bbb' align='center' valign='top'>" & _
	              "(NENHUMA ANOTAÇÃO ENCONTRADA)" & _
	              "</td>" & _
	              "</tr>"
	end if
	
	Response.Write strResp

'	FECHA CONEXAO COM O BANCO DE DADOS
	if rs.State <> 0 then rs.Close
	set rs = nothing

	cn.Close
	set cn = nothing

%>
