<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp"        -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ===========================================
'	  G R U P O L O J A S A T U A L I Z A . A S P
'     ===========================================
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
'			I N I C I A L I Z A     P Á G I N A     A S P     N O     S E R V I D O R
' _____________________________________________________________________________________________


	On Error GoTo 0
	Err.Clear
	
	dim s, s_aux, usuario, alerta
	
	usuario = trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, r
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	Dim criou_novo_reg
	Dim s_log, s_log_id_operacao, s_log_loja_inclui, s_log_loja_exclui
	Dim campos_a_omitir
	Dim vLog1()
	Dim vLog2()
	s_log = ""
	s_log_loja_inclui = ""
	s_log_loja_exclui = ""
	campos_a_omitir = ""
	
'	OBTÉM DADOS DO FORMULÁRIO ANTERIOR
	dim operacao_selecionada, s_grupo_lojas, s_descricao
	dim ckb_inclui, ckb_exclui, v_inclui, v_exclui, i, v
	operacao_selecionada=request("operacao_selecionada")
	s_grupo_lojas=retorna_so_digitos(trim(request("grupo_lojas_selecionado")))
	s_descricao=Trim(request("c_descricao"))

'	RETORNA OS VALORES DOS CHECKBOXES MARCADOS SEPARADOS POR VÍRGULA
'	LEMBRE-SE: 1) OS CHECKBOXES NÃO MARCADOS RETORNAM UMA STRING VAZIA.
'	========== 2) EM UM ARRAY DE CHECKBOXES, OCORRE O SEGUINTE: SE O ARRAY 
'				  NO FORMULÁRIO TEM 10 ITENS E APENAS 3 ESTÃO MARCADOS, AO
'				  LER ESSE ARRAY NA PÁGINA ASP, O ARRAY IRÁ TER APENAS OS 3
'				  ITENS QUE ESTAVAM MARCADOS. NESTE CASO, A CHAMADA
'				  Request.Form("nome_do_campo").Count IRÁ RETORNAR O VALOR 3.
	ckb_inclui=Trim(Request.Form("ckb_inclui"))
	ckb_exclui=Trim(Request.Form("ckb_exclui"))
	
	if s_grupo_lojas = "" then Response.Redirect("aviso.asp?id=" & ERR_ID_INVALIDO)
	s_grupo_lojas=normaliza_codigo(s_grupo_lojas, TAM_MIN_GRUPO_LOJAS)

	redim v_inclui(0)
	v_inclui(Ubound(v_inclui)) = ""
	
	redim v_exclui(0)
	v_exclui(Ubound(v_exclui)) = ""

	if ckb_inclui <> "" then
		v = split(ckb_inclui,",",-1)
		for i = Lbound(v) to Ubound(v)
			if Trim(v(i)) <> "" then
				if v_inclui(Ubound(v_inclui)) <> "" then redim preserve v_inclui(Ubound(v_inclui)+1)
				v_inclui(Ubound(v_inclui)) = normaliza_codigo(Trim(v(i)), TAM_MIN_LOJA)
				end if
			next
		end if

	if ckb_exclui <> "" then
		v = split(ckb_exclui,",",-1)
		for i = Lbound(v) to Ubound(v)
			if Trim(v(i)) <> "" then
				if v_exclui(Ubound(v_exclui)) <> "" then redim preserve v_exclui(Ubound(v_exclui)+1)
				v_exclui(Ubound(v_exclui)) = normaliza_codigo(Trim(v(i)), TAM_MIN_LOJA)
				end if
			next
		end if

	dim erro_consistencia, erro_fatal
	erro_consistencia=false
	erro_fatal=false
	
	alerta = ""
	if s_grupo_lojas = "" then
		alerta="NÚMERO DO GRUPO DE LOJAS É INVÁLIDO."
	elseif s_descricao = "" then
		alerta="PREENCHA A DESCRIÇÃO PARA O GRUPO DE LOJAS."
		end if
	
	if alerta <> "" then erro_consistencia=True
		
	Err.Clear
	
	dim msg_erro
	if Not cria_recordset_otimista(r, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

'	EXECUTA OPERAÇÃO NO BD
	select case operacao_selecionada
		case OP_EXCLUI
		'	 =========
			if Not erro_fatal then
			'	INFO P/ LOG
				s="SELECT * FROM t_LOJA_GRUPO WHERE grupo = '" & s_grupo_lojas & "'"
				r.Open s, cn
				if Not r.EOF then
					log_via_vetor_carrega_do_recordset r, vLog1, campos_a_omitir
					s_log = log_via_vetor_monta_exclusao(vLog1)
					end if
				r.Close
				
				s="SELECT * FROM t_LOJA_GRUPO_ITEM WHERE grupo = '" & s_grupo_lojas & "' ORDER BY CONVERT(smallint,loja)"
				r.Open s, cn
				s = ""
				do while Not r.EOF
					if s <> "" then s = s & ", "
					s = s & Trim("" & r("loja"))
					r.movenext
					loop
				r.Close
				if s <> "" then 
					s = "lojas pertencentes ao grupo: " & s
					if s_log <> "" then s_log = s_log & ", "
					s_log = s_log & s
					end if
				
			'	APAGA!!
			'	~~~~~~~~~~~~~
				cn.BeginTrans
			'	~~~~~~~~~~~~~
				s="DELETE FROM t_LOJA_GRUPO_ITEM WHERE grupo = '" & s_grupo_lojas & "'"
				cn.Execute(s)
				if Err <> 0 then
				'	~~~~~~~~~~~~~~~~
					cn.RollbackTrans
				'	~~~~~~~~~~~~~~~~
					Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
					end if
				
				s="DELETE FROM t_LOJA_GRUPO WHERE grupo = '" & s_grupo_lojas & "'"
				cn.Execute(s)
				if Err <> 0 then
				'	~~~~~~~~~~~~~~~~
					cn.RollbackTrans
				'	~~~~~~~~~~~~~~~~
					Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
					end if
				
				if s_log <> "" then grava_log usuario, "", "", "", OP_LOG_GRUPO_LOJAS_EXCLUSAO, s_log

			'	~~~~~~~~~~~~~~
				cn.CommitTrans
			'	~~~~~~~~~~~~~~
				If Err <> 0 then 
					erro_fatal=True
					alerta = "FALHA AO REMOVER O GRUPO DE LOJAS (" & Cstr(Err) & ": " & Err.Description & ")."
					end if
				end if



		case OP_INCLUI, OP_CONSULTA
		'	 ======================
			if alerta = "" then 
			'	~~~~~~~~~~~~~
				cn.BeginTrans
			'	~~~~~~~~~~~~~
				s = "SELECT * FROM t_LOJA_GRUPO WHERE grupo = '" & s_grupo_lojas & "'"
				r.Open s, cn
				if r.EOF then 
					r.AddNew 
					criou_novo_reg = True
					r("grupo")=s_grupo_lojas
					r("dt_cadastro") = Date
				else
					criou_novo_reg = False
					log_via_vetor_carrega_do_recordset r, vLog1, campos_a_omitir
					end if
					
				r("descricao")=s_descricao
				r("dt_ult_atualizacao")=Now

				log_via_vetor_carrega_do_recordset r, vLog2, campos_a_omitir
				r.Update
				if Err <> 0 then
				'	~~~~~~~~~~~~~~~~
					cn.RollbackTrans
				'	~~~~~~~~~~~~~~~~
					Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
					end if
				r.Close
				
			'	EXCLUI LOJAS DO GRUPO
				s = ""
				for i=Lbound(v_exclui) to Ubound(v_exclui)
					if Trim(v_exclui(i))<>"" then
						if s <> "" then s = s & " OR"
						s = s & " (loja='" & Trim(v_exclui(i)) & "')"
						if s_log_loja_exclui <> "" then s_log_loja_exclui = s_log_loja_exclui & ", "
						s_log_loja_exclui = s_log_loja_exclui & Trim(v_exclui(i))
						end if
					next
				if s_log_loja_exclui <> "" then s_log_loja_exclui = "loja(s) excluída(s) do grupo: " & s_log_loja_exclui
				
				if s <> "" then
					s = "DELETE FROM t_LOJA_GRUPO_ITEM WHERE" & _
						" (grupo = '" & s_grupo_lojas & "')" & _
						" AND (" & s & ")"
					cn.Execute(s)
					if Err <> 0 then
					'	~~~~~~~~~~~~~~~~
						cn.RollbackTrans
					'	~~~~~~~~~~~~~~~~
						Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
						end if
					end if
				
				
			'	INCLUI LOJAS AO GRUPO
				for i=Lbound(v_inclui) to Ubound(v_inclui)
					if Trim(v_inclui(i))<>"" then
						s = "INSERT INTO t_LOJA_GRUPO_ITEM (grupo, loja, dt_cadastro)" & _
							" VALUES (" & _
							"'" & s_grupo_lojas & "'" & _
							",'" & Trim(v_inclui(i)) & "'" & _
							"," & bd_monta_data(Date) & _
							")"
						cn.Execute(s)
						if Err <> 0 then
						'	~~~~~~~~~~~~~~~~
							cn.RollbackTrans
						'	~~~~~~~~~~~~~~~~
							Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
							end if
						if s_log_loja_inclui <> "" then s_log_loja_inclui = s_log_loja_inclui & ", "
						s_log_loja_inclui = s_log_loja_inclui & Trim(v_inclui(i))
						end if
					next
				if s_log_loja_inclui <> "" then s_log_loja_inclui = "loja(s) adicionada(s) ao grupo: " & s_log_loja_inclui
				
				
				if alerta = "" then 
					if criou_novo_reg then
						s_log = log_via_vetor_monta_inclusao(vLog2)
						s_log_id_operacao = OP_LOG_GRUPO_LOJAS_INCLUSAO
					else
						s_log = log_via_vetor_monta_alteracao(vLog1, vLog2)
						if s_log <> "" then 
							s_log="grupo de lojas=" & s_grupo_lojas & "; " & s_log
							s_log_id_operacao = OP_LOG_GRUPO_LOJAS_ALTERACAO
							end if
						end if
					if s_log_loja_inclui <> "" then
						if s_log <> "" then s_log = s_log & "; "
						s_log = s_log & s_log_loja_inclui
						end if
					if s_log_loja_exclui <> "" then
						if s_log <> "" then s_log = s_log & "; "
						s_log = s_log & s_log_loja_exclui
						end if
					if s_log <> "" then grava_log usuario, "", "", "", s_log_id_operacao, s_log
					end if
				
				if alerta = "" then 
				'	~~~~~~~~~~~~~~
					cn.CommitTrans
				'	~~~~~~~~~~~~~~
				else
				'	~~~~~~~~~~~~~~~~
					cn.RollbackTrans
				'	~~~~~~~~~~~~~~~~
					end if
				
				set r = nothing
				end if
		
		
		case else
		'	 ====
			alerta="OPERAÇÃO INVÁLIDA."
			
		end select
		
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
	<title>CENTRAL ADMINISTRATIVA</title>
	</head>



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


<body onload="bVOLTAR.focus();">
<center>
<br>

<!--  T E L A  -->

<p class="T">A V I S O</p>

<% 
	s = ""
	s_aux="'MtAviso'"
	if alerta <> "" then
		s = "<P style='margin:5px 2px 5px 2px;'>" & alerta & "</P>"
		s_aux="'MtAlerta'"
	else
		select case operacao_selecionada
			case OP_INCLUI
				s = "GRUPO DE LOJAS " & chr(34) & s_grupo_lojas & chr(34) & " CADASTRADO COM SUCESSO."
			case OP_CONSULTA, OP_ALTERA
				s = "GRUPO DE LOJAS " & chr(34) & s_grupo_lojas & chr(34) & " ALTERADO COM SUCESSO."
			case OP_EXCLUI
				s = "GRUPO DE LOJAS " & chr(34) & s_grupo_lojas & chr(34) & " EXCLUÍDO COM SUCESSO."
			end select
		if s <> "" then s="<P style='margin:5px 2px 5px 2px;'>" & s & "</P>"
		end if
%>
<div class=<%=s_aux%> style="width:400px;FONT-WEIGHT:bold;" align="CENTER"><%=s%></div>
<BR><BR>

<p class="TracoBottom"></p>

<table cellSpacing="0">
<tr>
<%
	s="GrupoLojas.asp" & "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))
	if erro_consistencia And (Not erro_fatal) then s="javascript:history.back()"
%>
	<td align="CENTER"><a name="bVOLTAR" id="bVOLTAR" href="<%=s%>"><img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
</tr>
</table>

</center>
</body>
</html>


<%
'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>