<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ==========================================================
'	  P E D I D O P R E D E V O L U Ç A O C A N C E L A . A S P
'     ==========================================================
'
'
'	  S E R V E R   S I D E   S C R I P T I N G
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


	On Error GoTo 0
	Err.Clear

	dim s, msg_erro
	dim usuario
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO)

	dim alerta
	alerta=""
	
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

	dim pedido_selecionado
	pedido_selecionado = Trim(request("pedido_selecionado"))
	if (pedido_selecionado = "") then Response.Redirect("aviso.asp?id=" & ERR_PEDIDO_NAO_ESPECIFICADO)

	dim id_devolucao, st_devolucao
	id_devolucao = Trim(request("id_devolucao"))
	if (id_devolucao = "") then Response.Redirect("aviso.asp?id=" & ERR_PARAMETRO_OBRIGATORIO_NAO_ESPECIFICADO)

    st_devolucao = Trim(Request("st_devolucao"))
    if st_devolucao <> COD_ST_PEDIDO_DEVOLUCAO__CADASTRADA then 
        alerta = "Status da pré-devolução não permite cancelamento."
        end if
	

	dim campos_a_omitir
	dim vLog()
	dim s_log
	s_log = ""
	campos_a_omitir = "|dt_cadastro|dt_hr_cadastro|"


'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	if alerta = "" then
	'	INICIA A TRANSAÇÃO
	'	~~~~~~~~~~~~~
		cn.BeginTrans
	'	~~~~~~~~~~~~~
		if Not cria_recordset_pessimista(rs, msg_erro) then
		'	~~~~~~~~~~~~~~~~
			cn.RollbackTrans
		'	~~~~~~~~~~~~~~~~
			Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
			end if
		
		if alerta = "" then
			s = "SELECT * FROM t_PEDIDO_DEVOLUCAO WHERE (id = " & id_devolucao & ")"
			rs.Open s, cn
			rs("usuario_ult_atualizacao") = usuario
            rs("dt_ult_atualizacao") = Date
            rs("dt_hr_ult_atualizacao") = Now
            rs("status") = COD_ST_PEDIDO_DEVOLUCAO__CANCELADA
            rs("status_usuario") = usuario
            rs("status_data") = Date
            rs("status_data_hora") = Now
            rs("st_cancelado") = 1
            rs("usuario_cancelado") = usuario
            rs("dt_cancelado") = Date
            rs("dt_hr_cancelado") = Now
			rs.Update 
			if Err <> 0 then
			'	~~~~~~~~~~~~~~~~
				cn.RollbackTrans
			'	~~~~~~~~~~~~~~~~
				Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
				end if

            if Err=0 then
                s = "SELECT * FROM t_UPLOAD_FILE INNER JOIN t_PEDIDO_DEVOLUCAO_IMAGEM ON (t_UPLOAD_FILE.id=t_PEDIDO_DEVOLUCAO_IMAGEM.id_upload_file)" & _
                        " WHERE (id_pedido_devolucao = " & id_devolucao & ")"
                if rs.State <> 0 then rs.Close
			    rs.Open s, cn
                if Not rs.Eof then
                    do while Not rs.Eof
                        rs("st_delete_file")=1
                        rs("usuario_delete_file")=usuario
                        rs("dt_delete_file")=Date
                        rs("dt_hr_delete_file")=Now
                        rs("dt_delete_file_scheduled_date")=DateAdd("m", 1, Date)
                        rs.Update
                        rs.MoveNext
                        loop
                    end if
                    if Err <> 0 then
			    '	~~~~~~~~~~~~~~~~
				    cn.RollbackTrans
			    '	~~~~~~~~~~~~~~~~
				    Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
				    end if
                end if
			
			log_via_vetor_carrega_do_recordset rs, vLog, campos_a_omitir
			s_log = log_via_vetor_monta_inclusao(vLog)
			
			if rs.State <> 0 then rs.Close
			
			if s_log <> "" then grava_log usuario, "", pedido_selecionado, "", "PED DEVOL CANCELA", s_log
			end if

		if alerta = "" then
		'	~~~~~~~~~~~~~~
			cn.CommitTrans
		'	~~~~~~~~~~~~~~
			if Err=0 then 
				s = "Pré-devolução cancelada com sucesso!!"
				Session(SESSION_CLIPBOARD) = s
				Response.Redirect("mensagem.asp" & "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")))
			else
				alerta=Cstr(Err) & ": " & Err.Description
				end if
		else
		'	~~~~~~~~~~~~~~~~
			cn.RollbackTrans
		'	~~~~~~~~~~~~~~~~
			end if
		end if

%>




<%
'	  C L I E N T   S I D E   S C R I P T I N G
'
'      CCCCCCC   LLL        IIIII  EEEEEEEEE  NNN   NNN  TTTTTTTTT EEEEEEEEE
'     CCC   CCC  LLL         III   EEE        NNNN  NNN     TTT    EEE
'     CCC        LLL         III   EEE        NNNNN NNN     TTT    EEE
'     CCC        LLL         III   EEEEEE     NNN NNNNN     TTT    EEEEEE
'     CCC        LLL         III   EEE        NNN  NNNN     TTT    EEE
'     CCC   CCC  LLL   LLL   III   EEE        NNN   NNN     TTT    EEE
'      CCCCCCC   LLLLLLLLL  IIIII  EEEEEEEEE  NNN   NNN     TTT    EEEEEEEEE
%>



<html>


<head>
	<title>CENTRAL</title>
	</head>



<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>



<!-- C A S C A D I N G   S T Y L E   S H E E T

	 CCCCCCC    SSSSSSS    SSSSSSS
	CCC   CCC  SSS   SSS  SSS   SSS
	CCC        SSS        SSS
	CCC         SSSS       SSSS
	CCC            SSSS       SSSS
	CCC   CCC  SSS   SSS  SSS   SSS
	 CCCCCCC    SSSSSSS    SSSSSSS
-->

<link href="<%=URL_FILE__E_CSS%>" Rel="stylesheet" Type="text/css">

<% if alerta <> "" then %>
<!-- ************************************************************ -->
<!-- **********  PÁGINA PARA EXIBIR MENSAGENS DE ERRO  ********** -->
<!-- ************************************************************ -->
<body onload="bVOLTAR.focus();">
<center>
<br>
<!--  T E L A  -->
<p class="T">A V I S O</p>
<div class="MtAlerta" style="width:600px;font-weight:bold;" align="center"><p style='margin:5px 2px 5px 2px;'><%=alerta%></p></div>
<br><br>
<p class="TracoBottom"></p>
<table cellSpacing="0">
<tr>
	<td align="center"><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()"><img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
</tr>
</table>
</center>
</body>
<% end if %>

</html>


<%
'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>