<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<% Response.Expires=-1 %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<%
'     ========================================================
'	  AjaxCarregaOrcamentosNovos.asp
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
	dim cn, rs, s

	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

	dim s_loja, s_vendedor, s_where, s_from
	s_loja = UCase(Trim(Request("loja")))
    s_vendedor = Trim(Request("vendedor"))

    '	MONTA CLÁUSULA WHERE
	s_where = ""

'	CRITÉRIO: STATUS DE FECHAMENTO DO ORÇAMENTOS
	s = "(st_fechamento='') OR (st_fechamento IS NULL)"
	if s_where <> "" then s_where = s_where & " AND"
	s_where = s_where & " (" & s & ")"

'	CRITÉRIO: STATUS DO ORÇAMENTO
	s = "(st_orcamento='') OR (st_orcamento IS NULL)"
	if s_where <> "" then s_where = s_where & " AND"
	s_where = s_where & " (" & s & ")"

'	CRITÉRIO: ORÇAMENTOS QUE NÃO VIRARAM PEDIDOS (NOVO CAMPO DE CONTROLE)
	s = "(st_orc_virou_pedido = 0)"
	if s_where <> "" then s_where = s_where & " AND"
	s_where = s_where & " (" & s & ")"

'	CRITÉRIO: LOJA (CADA LOJA SÓ PODE CONSULTAR SEUS PRÓPRIOS ORÇAMENTOS)
	s = "(CONVERT(smallint,loja) = " & s_loja & ")"
	if s_where <> "" then s_where = s_where & " AND"
	s_where = s_where & " (" & s & ")"

	if s_vendedor <> "" then
	'	CRITÉRIO: VENDEDOR (PODE ACESSAR ORÇAMENTOS DE TODOS OS VENDEDORES DA LOJA?)
		s = "(vendedor = '" & s_vendedor & "')"
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (" & s & ")"
		end if
	
'	CLÁUSULA WHERE
	if s_where <> "" then s_where = " WHERE" & s_where
	
'	MONTA CLÁUSULA FROM
	s_from = " FROM t_ORCAMENTO"

	strSql = "SELECT loja," & _
			" data, orcamento," & _
			" orcamentista" & _
			s_from & _
			s_where

	strSql = strSql & " ORDER BY orcamentista, data, orcamento"
  

'	EXECUTA A CONSULTA
	strResp = ""
	
	if rs.State <> 0 then rs.Close
	rs.open strSql, cn

'	ENCONTROU DADOS
	do while Not rs.Eof

		    strResp = strResp & _
				  "<registro>" & _
                    "<orcamento>" & _
						Trim("" & rs("orcamento")) & _
					"</orcamento>" & _
					"<orcamentista><![CDATA[" & _
						Trim("" & rs("orcamentista")) & _
					"]]></orcamentista>" & _
                    "<data>" & _
						Trim("" & rs("data")) & _
					"</data>" & _
					"<loja>" & _
						Trim("" & rs("loja")) & _
					"</loja>" & _
				  "</registro>"

		rs.MoveNext
	loop

       
'	ENCONTROU ALGUMA RESPOSTA
	if strResp <> "" then 
		Response.ContentType="text/xml"
		Response.Charset="ISO-8859-1"
		strResp = "<?xml version='1.0' encoding='ISO-8859-1'?>" & _
				  "<TabelaOrcamentos>" & _
				  strResp & _
				  "</TabelaOrcamentos>"
		end if
	
	Response.Write strResp

'	FECHA CONEXAO COM O BANCO DE DADOS
	if rs.State <> 0 then rs.Close
	set rs = nothing

	cn.Close
	set cn = nothing

%>
