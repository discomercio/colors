<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp"        -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ===========================================
'	  M I D I A A T U A L I Z A . A S P
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
	Dim s_log
	Dim campos_a_omitir
	Dim vLog1()
	Dim vLog2()
	s_log = ""
	campos_a_omitir = ""
	
'	OBTÉM DADOS DO FORMULÁRIO ANTERIOR
	dim operacao_selecionada, s_midia, s_nome, s_indisponivel
	operacao_selecionada=request("operacao_selecionada")
	s_midia=retorna_so_digitos(trim(request("midia_selecionada")))
	s_nome=Trim(request("nome"))
	s_indisponivel=trim(request("rb_indisponivel"))

	if s_midia = "" then Response.Redirect("aviso.asp?id=" & ERR_ID_INVALIDO)
	
	s_midia=normaliza_codigo(s_midia, TAM_MIN_MIDIA)

	dim erro_consistencia, erro_fatal
	
	erro_consistencia=false
	erro_fatal=false
	
	alerta = ""
	if s_midia = "" then
		alerta="NÚMERO DE VEÍCULO DE MÍDIA INVÁLIDO."
	elseif s_nome = "" then
		alerta="PREENCHA O NOME (APELIDO) DO VEÍCULO DE MÍDIA."
	elseif s_indisponivel = "" then
		alerta="INFORME SE O VEÍCULO DE MÍDIA ESTÁ DISPONÍVEL OU NÃO."
		end if
	
	if alerta <> "" then erro_consistencia=True	
		
	Err.Clear
	
	dim msg_erro
	if Not cria_recordset_otimista(r, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

'	EXECUTA OPERAÇÃO NO BD
	select case operacao_selecionada
		case OP_EXCLUI
		'	 =========
			s="SELECT COUNT(*) AS qtde FROM t_CLIENTE WHERE (midia = '" & s_midia & "')"
			r.Open s, cn
		'	ASSEGURA QUE A COMPARAÇÃO SERÁ FEITA ENTRE MESMO TIPO DE DADOS
			if Cstr(r("qtde")) > Cstr(0) then
				erro_fatal=True
				alerta = "VEÍCULO DE MÍDIA NÃO PODE SER REMOVIDO PORQUE ESTÁ SENDO REFERENCIADO NA TABELA DE USUÁRIOS."
				end if
			r.Close 
			
			if Not erro_fatal then
				s="SELECT COUNT(*) AS qtde FROM t_PEDIDO WHERE (midia = '" & s_midia & "')"
				r.Open s, cn
			'	ASSEGURA QUE A COMPARAÇÃO SERÁ FEITA ENTRE MESMO TIPO DE DADOS
				if Cstr(r("qtde")) > Cstr(0) then
					erro_fatal=True
					alerta = "VEÍCULO DE MÍDIA NÃO PODE SER REMOVIDO PORQUE ESTÁ SENDO REFERENCIADO NA TABELA DE PEDIDOS."
					end if
				r.Close 
				end if
			
			if Not erro_fatal then
				s="SELECT COUNT(*) AS qtde FROM t_ORCAMENTO WHERE (midia = '" & s_midia & "')"
				r.Open s, cn
			'	ASSEGURA QUE A COMPARAÇÃO SERÁ FEITA ENTRE MESMO TIPO DE DADOS
				if Cstr(r("qtde")) > Cstr(0) then
					erro_fatal=True
					alerta = "VEÍCULO DE MÍDIA NÃO PODE SER REMOVIDO PORQUE ESTÁ SENDO REFERENCIADO NA TABELA DE ORÇAMENTOS."
					end if
				r.Close 
				end if
				
			if Not erro_fatal then
			'	INFO P/ LOG
				s="SELECT * FROM t_MIDIA WHERE id = '" & s_midia & "'"
				r.Open s, cn
				if Not r.EOF then
					log_via_vetor_carrega_do_recordset r, vLog1, campos_a_omitir
					s_log = log_via_vetor_monta_exclusao(vLog1)
					end if
				r.Close
			
			'	APAGA!!
				s="DELETE FROM t_MIDIA WHERE id = '" & s_midia & "'"
				cn.Execute(s)
				If Err = 0 then 
					if s_log <> "" then grava_log usuario, "", "", "", OP_LOG_MIDIA_EXCLUSAO, s_log
				else
					erro_fatal=True
					alerta = "FALHA AO REMOVER O VEÍCULO DE MÍDIA (" & Cstr(Err) & ": " & Err.Description & ")."
					end if
				end if


		case OP_INCLUI, OP_CONSULTA
		'	 ======================
			if alerta = "" then 
				s = "SELECT * FROM t_MIDIA WHERE id = '" & s_midia & "'"
				r.Open s, cn
				if r.EOF then 
					r.AddNew 
					criou_novo_reg = True
					r("id")=s_midia
					r("dt_cadastro") = Date
				else
					criou_novo_reg = False
					log_via_vetor_carrega_do_recordset r, vLog1, campos_a_omitir
					end if
					
				r("apelido")=s_nome
				r("indisponivel")=CLng(s_indisponivel)
				r("dt_ult_atualizacao")=Now
				
				r.Update

				If Err = 0 then 
					log_via_vetor_carrega_do_recordset r, vLog2, campos_a_omitir
					if criou_novo_reg then
						s_log = log_via_vetor_monta_inclusao(vLog2)
						if s_log <> "" then grava_log usuario, "", "", "", OP_LOG_MIDIA_INCLUSAO, s_log
					else
						s_log = log_via_vetor_monta_alteracao(vLog1, vLog2)
						if s_log <> "" then 
							s_log="midia=" & Trim("" & r("id")) & "; " & s_log
							grava_log usuario, "", "", "", OP_LOG_MIDIA_ALTERACAO, s_log
							end if
						end if
				else
					erro_fatal=True
					alerta = "FALHA AO GRAVAR OS DADOS (" & Cstr(Err) & ": " & Err.Description & ")."
					end if
				
				r.Close
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
				s = "VEÍCULO DE MÍDIA " & chr(34) & s_midia & chr(34) & " CADASTRADO COM SUCESSO."
			case OP_CONSULTA, OP_ALTERA
				s = "VEÍCULO DE MÍDIA " & chr(34) & s_midia & chr(34) & " ALTERADO COM SUCESSO."
			case OP_EXCLUI
				s = "VEÍCULO DE MÍDIA " & chr(34) & s_midia & chr(34) & " EXCLUÍDO COM SUCESSO."
			end select
		if s <> "" then s="<P style='margin:5px 2px 5px 2px;'>" & s & "</P>"
		end if
%>
<div class=<%=s_aux%> style="width:400px;font-weight:bold;" align="center"><%=s%></div>
<br><br>

<p class="TracoBottom"></p>

<table cellSpacing="0">
<tr>
<%
	s="midia.asp" & "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))
	if erro_consistencia And (Not erro_fatal) then s="javascript:history.back()"
%>
	<td align="center"><a name="bVOLTAR" id="bVOLTAR" href="<%=s%>"><img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
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