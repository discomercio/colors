<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<% Response.Expires=-1 %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<%
'     ========================================================
'	  JsonPesquisaIndicadores.asp
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

    Const TIPO_CONSULTA_INDICADORES_POR_CPFCNPJ = "CPFCNPJ"
    Const TIPO_CONSULTA_INDICADORES_POR_APELIDO = "APELIDO"
	
'	OBTEM O ID
	dim strSql, strResp, msg_erro
	
'	CONECTA COM O BANCO DE DADOS
	dim cn, r, rs
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
    
    dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

    dim q, loja, usuario, modulo, tipo_consulta
    dim flag_ok
    q = Trim(Request("q"))
    loja = Trim(Request("loja"))
    usuario = Trim(Request("usuario"))
    modulo = Trim(Request("modulo"))
    tipo_consulta = Trim(Request("tipo_consulta"))

    if Trim(q) <> "" then
        q = Replace(q, "*", "%")
        if instr(q, "%") = 0 then
            if Right(q, 1) <> "%" then q = q & "%"
        end if
    end if

    strSql = "SELECT * FROM t_ORCAMENTISTA_E_INDICADOR"

    if Trim(q) <> "" then
        if tipo_consulta = TIPO_CONSULTA_INDICADORES_POR_APELIDO then
            strSql = strSql & " WHERE (apelido LIKE '" & q & "')"
        elseif tipo_consulta = TIPO_CONSULTA_INDICADORES_POR_CPFCNPJ then
            strSql = strSql & " WHERE (cnpj_cpf LIKE '" & q & "')"
        end if
    end if    
    strSql = strSql & " ORDER BY apelido"
     
    strResp = ""
    if rs.State <> 0 then rs.Close
	rs.open strSql, cn

    if modulo = COD_OP_MODULO_LOJA then    
        do while Not rs.Eof
            flag_ok = False
            if converte_numero(Trim("" & rs("loja"))) = converte_numero(loja) then flag_ok = True
            if Not flag_ok then
                if Trim("" & rs("vendedor")) <> "" then
				    strSql = "SELECT " & _
							    "*" & _
						    " FROM t_USUARIO_X_LOJA" & _
						    " WHERE" & _
							    " (CONVERT(smallint, loja) = " & loja & ")" & _
							    " AND (usuario = '" & Trim("" & rs("vendedor")) & "')"
				    set r = cn.Execute(strSql)
				    if Not r.Eof then flag_ok = True
				end if
            end if
            if Not operacao_permitida(OP_LJA_CONSULTA_UNIVERSAL_PEDIDO_ORCAMENTO, s_lista_operacoes_permitidas) then
			    if Not isLojaVrf(loja) then
				    if rs("vendedor") <> usuario then flag_ok = False
				end if
		    end if              

            if flag_ok then
                if strResp <> "" then strResp = strResp & ", " & chr(13)
                if tipo_consulta = TIPO_CONSULTA_INDICADORES_POR_APELIDO then
                    strResp = strResp & """" & rs("apelido") & """"
                elseif tipo_consulta = TIPO_CONSULTA_INDICADORES_POR_CPFCNPJ then
                    strResp = strResp & """" & cnpj_cpf_formata(rs("cnpj_cpf")) & """"
                end if
            end if
            rs.MoveNext
        loop
    elseif modulo = COD_OP_MODULO_CENTRAL then
        do while Not rs.Eof
            if strResp <> "" then strResp = strResp & ", " & chr(13)
                if tipo_consulta = TIPO_CONSULTA_INDICADORES_POR_APELIDO then
                    strResp = strResp & """" & rs("apelido") & """"
                elseif tipo_consulta = TIPO_CONSULTA_INDICADORES_POR_CPFCNPJ then
                    strResp = strResp & """" & cnpj_cpf_formata(rs("cnpj_cpf")) & """"
            end if
            rs.MoveNext
        loop
    end if

	
    strResp = "[" & chr(13) & _
                    strResp & chr(13) & _
                "]"

'	ENCONTROU ALGUMA RESPOSTA
    Response.ContentType="application/json"	
	Response.Write strResp

'	FECHA CONEXAO COM O BANCO DE DADOS
	if rs.State <> 0 then rs.Close
	set rs = nothing

	cn.Close
	set cn = nothing

%>
