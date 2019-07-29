<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<% Response.Expires=-1 %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<%
'     =============================================================
'	  A J A X T R A N S P O R T A D O R A C E P P E S Q B D . A S P
'     =============================================================
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
	dim strSql, strTransp, msg_erro
	
'	CONECTA COM O BANCO DE DADOS
	dim cn, rs
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

	dim s_cepini, s_cepfim
	s_cepini = retorna_so_digitos(Trim(Request("cepini")))
	s_cepfim = retorna_so_digitos(Trim(Request("cepfim")))
	
	strSql = "SELECT " & _
				" transportadora_id," & _
				" tipo_range," & _
				" cep_unico," & _
				" cep_faixa_inicial," & _
				" cep_faixa_final " & _
			" FROM t_TRANSPORTADORA_CEP " & _
			" WHERE"

	strSql = strSql & " (" & _
		" ((tipo_range = 1) and (cep_unico = '" & s_cepini & "'))" & _
		" OR ((tipo_range = 2) and ('" & s_cepini & "' between cep_faixa_inicial and cep_faixa_final))" & _
		") "

	if s_cepfim <> "" then
		strSql = strSql & " OR (" & _
			" ((tipo_range = 1) and (cep_unico = '" & s_cepfim & "'))" & _
			" OR ((tipo_range = 2) and ('" & s_cepfim & "' between cep_faixa_inicial and cep_faixa_final))"
		'outra situação: cadastrar uma faixa que contenha um CEP já cadastrado
		'(exemplo: cadastrar a faixa 11111-111 a 22222-222, mas já foi cadastrado o CEP 12345-678, portanto, contido nesta faixa)
		strSql = strSql &  _
			" OR ((tipo_range = 1) and (cep_unico between '" & s_cepini & "' and '" & s_cepfim & "'))" & _
			") "
		end if
		
	
	
'	EXECUTA A CONSULTA
	strTransp = ""
	
	if rs.State <> 0 then rs.Close
	rs.open strSql, cn
'	ENCONTROU DADOS NA TABELA CEPs DE ENTREGA
	if Not rs.Eof then
		strTransp = rs("transportadora_id")
		end if

	Response.Write strTransp

'	FECHA CONEXAO COM O BANCO DE DADOS
	if rs.State <> 0 then rs.Close
	set rs = nothing

	cn.Close
	set cn = nothing

%>
