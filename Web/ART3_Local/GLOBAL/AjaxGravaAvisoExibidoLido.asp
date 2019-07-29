<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<% Response.Expires=-1 %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<%
'     ========================================================
'	  AjaxGravaAvisoExibidoLido.asp
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
	dim strSql, strResp, strRazaoSocialNome, msg_erro, alerta
	
'	CONECTA COM O BANCO DE DADOS
	dim cn, r
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(r, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

    Dim s_log
	s_log = ""
	
'	OBTÉM DADOS DO FORMULÁRIO ANTERIOR
	dim aviso_selecionado, usuario
	Dim vAviso, i, s, sql_insere
	aviso_selecionado=trim(request("aviso_selecionado"))
    usuario = trim(request("usuario"))

	if aviso_selecionado<>"" then vAviso=split(aviso_selecionado,"|", -1)

	dim erro_consistencia, erro_fatal
	
	erro_consistencia=false
	erro_fatal=false
	alerta = ""
	
	if Trim(aviso_selecionado) = "" then alerta = "NENHUM AVISO SELECIONADO."

	if alerta <> "" then erro_consistencia=True	
	
	Err.Clear

'	EXECUTA OPERAÇÃO NO BD
	if alerta = "" then 
        sql_insere = ""
		cn.BeginTrans
		for i=Lbound(vAviso) to Ubound(vAviso)
			if sql_insere <> "" then sql_insere = sql_insere & ", "
            sql_insere = sql_insere & "('" & Trim(vAviso(i)) & "', '" & usuario & "', GETDATE())"

			next

		sql_insere = "INSERT INTO t_AVISO_LIDO (id, usuario, data) VALUES " & sql_insere
        set r = cn.Execute(sql_insere)

		if Err = 0 then 
			cn.CommitTrans 
		else 
			cn.RollbackTrans
		end if
    end if

'	FECHA CONEXAO COM O BANCO DE DADOS
	if r.State <> 0 then r.Close
	set r = nothing

	cn.Close
	set cn = nothing

%>
