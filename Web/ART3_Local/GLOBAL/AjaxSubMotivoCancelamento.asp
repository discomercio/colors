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


    dim s, x, r, ha_default,grupo_pai,codigo_pai,id_default
    id_default = ""
    grupo_pai = Trim(Request("grupo_pai"))
    codigo_pai = Trim(Request("codigo_pai"))

	id_default = Trim("" & id_default)
	ha_default=False
	s = "SELECT descricao,codigo" & _
        " FROM t_CODIGO_DESCRICAO" & _
        " WHERE grupo = 'CancelamentoPedido_Motivo_Sub'" & _
	        " AND (" & _
		        " grupo_pai = '" & grupo_pai & "'" & _
		        " AND codigo_pai = " & codigo_pai & "" & _
		        " )"

	set rs = cn.Execute(s)
	strResp = ""
	do while Not rs.Eof
		x = UCase(Trim("" & rs("codigo")))
		if (id_default<>"") And (id_default=x) then
			strResp = strResp & "<option selected"
			ha_default=True
		else
			strResp = strResp & "<option"
			end if
		strResp = strResp & " value='" & x & "'>"
		strResp = strResp & iniciais_em_maiusculas(Trim("" & rs("descricao")))
		strResp = strResp & "</option>" & chr(13)
		rs.MoveNext
		loop
       	
	Response.Write strResp

	if rs.State <> 0 then rs.Close
	set rs = nothing

	cn.Close
	set cn = nothing

%>
