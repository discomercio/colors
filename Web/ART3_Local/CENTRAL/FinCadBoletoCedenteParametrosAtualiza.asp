<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp"        -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ===========================================
'	  FinCadBoletoCedenteParametrosAtualiza.asp
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
	dim cn, r, msg_erro
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	if Not cria_recordset_otimista(r, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

	Dim criou_novo_reg
	Dim s_log
	Dim campos_a_omitir
	Dim vLog1()
	Dim vLog2()
	s_log = ""
	campos_a_omitir = ""
	
'	OBTÉM DADOS DO FORMULÁRIO ANTERIOR
	dim intNsuNovo
	dim operacao_selecionada
	dim s_id_conta_corrente_selecionado, s_id_boleto_cedente_selecionado
	dim s_num_banco, s_nome_banco, s_agencia, s_digito_agencia, s_conta, s_digito_conta
	dim s_carteira, s_codigo_empresa, s_nome_empresa
	dim s_juros_mora, s_perc_multa, s_nsu_arq_remessa, s_st_ativo, s_qtde_dias_protestar_apos_padrao
	dim s_segunda_mensagem_padrao, s_mensagem_1_padrao, s_mensagem_2_padrao
	dim s_mensagem_3_padrao, s_mensagem_4_padrao
	dim s_endereco, s_endereco_numero, s_endereco_complemento, s_bairro, s_cidade, s_uf, s_cep
	dim s_apelido, s_loja_default_boleto_plano_contas
	operacao_selecionada=Request.Form("operacao_selecionada")
	s_id_conta_corrente_selecionado=""
	s_id_boleto_cedente_selecionado=""
	if operacao_selecionada=OP_INCLUI then
		s_id_conta_corrente_selecionado=retorna_so_digitos(Trim(Request.Form("id_conta_corrente_selecionado")))
	else
		s_id_boleto_cedente_selecionado=retorna_so_digitos(Trim(Request.Form("id_boleto_cedente_selecionado")))
		end if
	
	s_num_banco=retorna_so_digitos(Trim(Request.Form("c_num_banco")))
	s_nome_banco=Trim(Request.Form("c_nome_banco"))
	s_agencia=retorna_so_digitos(Trim(Request.Form("c_agencia")))
	s_digito_agencia=retorna_so_digitos(Trim(Request.Form("c_digito_agencia")))
	s_conta=retorna_so_digitos(Trim(Request.Form("c_conta")))
	s_digito_conta=retorna_so_digitos(Trim(Request.Form("c_digito_conta")))
	s_carteira=retorna_so_digitos(Trim(Request.Form("c_carteira")))
	s_codigo_empresa=retorna_so_digitos(Trim(Request.Form("c_codigo_empresa")))
	s_nome_empresa=Trim(Request.Form("c_nome_empresa"))
	s_juros_mora=Trim(Request.Form("c_juros_mora"))
	s_perc_multa=Trim(Request.Form("c_perc_multa"))
	s_nsu_arq_remessa=retorna_so_digitos(Trim(Request.Form("c_nsu_arq_remessa")))
	s_st_ativo=Trim(Request.Form("rb_st_ativo"))
	s_qtde_dias_protestar_apos_padrao=retorna_so_digitos(Trim(Request.Form("c_qtde_dias_protestar_apos_padrao")))
	if s_qtde_dias_protestar_apos_padrao="" then s_qtde_dias_protestar_apos_padrao="0"
	s_segunda_mensagem_padrao=Trim(Request.Form("c_segunda_mensagem_padrao"))
	s_mensagem_1_padrao=Trim(Request.Form("c_mensagem_1_padrao"))
	s_mensagem_2_padrao=Trim(Request.Form("c_mensagem_2_padrao"))
	s_mensagem_3_padrao=Trim(Request.Form("c_mensagem_3_padrao"))
	s_mensagem_4_padrao=Trim(Request.Form("c_mensagem_4_padrao"))
	s_endereco=Trim(request("endereco"))
	s_endereco_numero=Trim(request("endereco_numero"))
	s_endereco_complemento=Trim(request("endereco_complemento"))
	s_bairro=Trim(request("bairro"))
	s_cidade=Trim(request("cidade"))
	s_uf=Ucase(Trim(request("uf")))
	s_cep=retorna_so_digitos(Trim(request("cep")))
	s_apelido=Trim(request("c_apelido"))
	s_loja_default_boleto_plano_contas=Trim(request("c_loja_default_boleto_plano_contas"))
	if s_loja_default_boleto_plano_contas<>"" then s_loja_default_boleto_plano_contas=normaliza_codigo(s_loja_default_boleto_plano_contas, TAM_MIN_LOJA)

	if operacao_selecionada=OP_INCLUI then
		if converte_numero(s_id_conta_corrente_selecionado) <= 0 then Response.Redirect("aviso.asp?id=" & ERR_ID_INVALIDO)
	else
		if converte_numero(s_id_boleto_cedente_selecionado) <= 0 then Response.Redirect("aviso.asp?id=" & ERR_ID_INVALIDO)
		end if
	
	dim erro_consistencia, erro_fatal
	
	erro_consistencia=false
	erro_fatal=false
	
	alerta = ""
	
	if alerta = "" then
		if operacao_selecionada=OP_INCLUI then
			if s_id_conta_corrente_selecionado = "" then alerta="NÚMERO DE IDENTIFICAÇÃO DA CONTA CORRENTE NÃO FOI FORNECIDO."
		else
			if s_id_boleto_cedente_selecionado = "" then alerta="NÚMERO DE IDENTIFICAÇÃO DA CONTA DO CEDENTE NÃO FOI FORNECIDO."
			end if
		end if
	
	if alerta = "" then
		if s_num_banco = "" then
			alerta="INFORME O NÚMERO DO BANCO."
		elseif s_nome_banco = "" then
			alerta="INFORME O NOME DO BANCO."
		elseif s_agencia = "" then
			alerta="INFORME O NÚMERO DA AGÊNCIA."
		elseif s_digito_agencia = "" then
			alerta="INFORME O DÍGITO DA AGÊNCIA."
		elseif s_conta = "" then
			alerta="INFORME O NÚMERO DA CONTA."
		elseif s_digito_conta = "" then
			alerta="INFORME O DÍGITO DA CONTA."
		elseif s_carteira = "" then
			alerta="INFORME O NÚMERO DA CARTEIRA."
		elseif s_codigo_empresa = "" then
			alerta="INFORME O CÓDIGO DA EMPRESA."
		elseif s_nome_empresa = "" then
			alerta="INFORME O NOME DA EMPRESA."
		elseif s_juros_mora = "" then
			alerta="INFORME O PERCENTUAL DO JUROS DE MORA AO MÊS."
		elseif converte_numero(s_juros_mora) < 0 then
			alerta="PERCENTUAL DO JUROS DE MORA NÃO PODE SER NEGATIVO."
		elseif s_perc_multa = "" then
			alerta="INFORME O PERCENTUAL DA MULTA."
		elseif converte_numero(s_perc_multa) < 0 then
			alerta="PERCENTUAL DA MULTA NÃO PODE SER NEGATIVO."
		elseif converte_numero(s_perc_multa) > 100 then
			alerta="PERCENTUAL DA MULTA NÃO PODE EXCEDER 100%"
		elseif s_nsu_arq_remessa = "" then
			alerta="INFORME O NÚMERO SEQUENCIAL DE REMESSA."
		elseif s_st_ativo = "" then
			alerta="INFORME O STATUS DA CONTA (ATIVO/INATIVO)."
		elseif s_endereco = "" then
			alerta="PREENCHA O ENDEREÇO."
		elseif Len(s_endereco) > CLng(MAX_TAMANHO_CAMPO_ENDERECO) then
			alerta="ENDEREÇO EXCEDE O TAMANHO MÁXIMO PERMITIDO:<br>TAMANHO ATUAL: " & Cstr(Len(s_endereco)) & " CARACTERES<br>TAMANHO MÁXIMO: " & Cstr(MAX_TAMANHO_CAMPO_ENDERECO) & " CARACTERES"
		elseif s_endereco_numero = "" then
			alerta="PREENCHA O NÚMERO DO ENDEREÇO."
		elseif s_bairro = "" then
			alerta="PREENCHA O BAIRRO."
		elseif s_cidade = "" then
			alerta="PREENCHA A CIDADE."
		elseif (s_uf="") Or (Not uf_ok(s_uf)) then
			alerta="UF INVÁLIDA."
		elseif s_cep = "" then
			alerta="INFORME O CEP."
		elseif Not cep_ok(s_cep) then
			alerta="CEP INVÁLIDO."
		elseif s_apelido="" then
			alerta="INFORME UM NOME CURTO (APELIDO) PARA O CEDENTE."
		elseif s_loja_default_boleto_plano_contas="" then
			alerta="INFORME O Nº DA LOJA PADRÃO PARA SE OBTER O PLANO DE CONTAS PARA O QUAL OS LANÇAMENTOS SERÃO VINCULADOS NA SITUAÇÃO EM QUE NÃO FOR POSSÍVEL DETERMINAR O Nº DA LOJA ASSOCIADA AO BOLETO."
			end if
		end if

	if alerta = "" then
		if operacao_selecionada <> OP_EXCLUI then
			s = "SELECT loja FROM t_LOJA WHERE (loja = '" & s_loja_default_boleto_plano_contas & "')"
			r.Open s, cn
			if r.Eof then
				alerta = "LOJA '" & s_loja_default_boleto_plano_contas & "' NÃO ESTÁ CADASTRADA NO SISTEMA."
				end if
			r.Close
			end if
		end if
	
	if alerta = "" then
		if operacao_selecionada=OP_INCLUI then
		'	GERA O NSU PARA O NOVO REGISTRO
			if Not fin_gera_nsu(T_FIN_BOLETO_CEDENTE, intNsuNovo, msg_erro) then
				alerta = "FALHA AO GERAR NSU PARA O NOVO REGISTRO (" & msg_erro & ")"
			else
				if intNsuNovo <= 0 then
					alerta = "NSU GERADO É INVÁLIDO (" & intNsuNovo & ")"
					end if
				end if
			
			if alerta = "" then
				s_id_boleto_cedente_selecionado = Cstr(intNsuNovo)
				end if
			end if
		end if

	if alerta = "" then
		if operacao_selecionada = OP_CONSULTA then
			s = "SELECT" & _
					" Coalesce(Max(nsu_arq_remessa),0) AS max_nsu_arq_remessa" & _
				" FROM t_FIN_BOLETO_ARQ_REMESSA" & _
				" WHERE" & _
					" id_boleto_cedente = " & s_id_boleto_cedente_selecionado
			r.Open s, cn
			if Not r.Eof then
				if Clng(r("max_nsu_arq_remessa")) > Clng(s_nsu_arq_remessa) then
					alerta = "O valor informado para o nº sequencial de remessa é inválido, pois o último arquivo de remessa gerado utilizou o número " & r("max_nsu_arq_remessa") & "!!"
					end if
				end if
			r.Close
			end if
		end if
	
	if alerta <> "" then erro_consistencia=True	
		
	Err.Clear
	
'	EXECUTA OPERAÇÃO NO BD
	select case operacao_selecionada
		case OP_EXCLUI
		'	 =========
			s = "SELECT" & _
					" TOP 1 *" & _
				" FROM t_FIN_BOLETO_ARQ_REMESSA" & _
				" WHERE" & _
					" (id_boleto_cedente = " & s_id_boleto_cedente_selecionado & ")"
			r.Open s, cn
			if Not r.Eof then
				erro_fatal=True
				alerta = "REGISTRO NÃO PODE SER REMOVIDO PORQUE ESTÁ SENDO REFERENCIADO NA TABELA DE ARQUIVO DE REMESSA DE BOLETOS."
				end if
			r.Close 

			if Not erro_fatal then
				s = "SELECT" & _
						" TOP 1 *" & _
					" FROM t_FIN_BOLETO" & _
					" WHERE" & _
						" (id_boleto_cedente = " & s_id_boleto_cedente_selecionado & ")"
				r.Open s, cn
				if Not r.Eof then
					erro_fatal=True
					alerta = "REGISTRO NÃO PODE SER REMOVIDO PORQUE ESTÁ SENDO REFERENCIADO NA TABELA DE BOLETOS."
					end if
				r.Close 
				end if
			
			if Not erro_fatal then
			'	INFO P/ LOG
				s = "SELECT " & _
						"*" & _
					" FROM t_FIN_BOLETO_CEDENTE" & _
					" WHERE" & _
						" (id = " & s_id_boleto_cedente_selecionado & ")"
				r.Open s, cn
				if Not r.EOF then
					log_via_vetor_carrega_do_recordset r, vLog1, campos_a_omitir
					s_log = log_via_vetor_monta_exclusao(vLog1)
					end if
				r.Close
			
			'	APAGA!!
				s = "DELETE" & _
					" FROM t_FIN_BOLETO_CEDENTE" & _
					" WHERE" &  _
						" (id = " & s_id_boleto_cedente_selecionado & ")"
				cn.Execute(s)
				If Err = 0 then 
					if s_log <> "" then grava_log usuario, "", "", "", OP_LOG_BOLETO_CEDENTE_PARAMETROS_EXCLUSAO, s_log
				else
					erro_fatal=True
					alerta = "FALHA AO EXCLUIR O REGISTRO (" & Cstr(Err) & ": " & Err.Description & ")."
					end if
				end if


		case OP_INCLUI, OP_CONSULTA
		'	 ======================
			if alerta = "" then 
				s = "SELECT " & _
						"*" & _
					" FROM t_FIN_BOLETO_CEDENTE" & _
					" WHERE" & _
						 " (id = " & s_id_boleto_cedente_selecionado & ")"
				r.Open s, cn
				if r.EOF then 
					r.AddNew 
					criou_novo_reg = True
					r("id") = CLng(s_id_boleto_cedente_selecionado)
					r("id_conta_corrente") = Cint(s_id_conta_corrente_selecionado)
					r("dt_cadastro") = Now
					r("usuario_cadastro") = usuario
					r("dt_indice_arq_remessa_no_dia") = Date
				else
					criou_novo_reg = False
					log_via_vetor_carrega_do_recordset r, vLog1, campos_a_omitir
					end if
				
				r("apelido")=s_apelido
				r("loja_default_boleto_plano_contas")=s_loja_default_boleto_plano_contas
				r("st_ativo")=CLng(s_st_ativo)
				r("nsu_arq_remessa")=CLng(s_nsu_arq_remessa)
				r("codigo_empresa")=s_codigo_empresa
				r("nome_empresa")=s_nome_empresa
				r("num_banco")=s_num_banco
				r("nome_banco")=s_nome_banco
				r("agencia")=s_agencia
				r("digito_agencia")=s_digito_agencia
				r("conta")=s_conta
				r("digito_conta")=s_digito_conta
				r("carteira")=s_carteira
				r("juros_mora")=converte_numero(s_juros_mora)
				r("perc_multa")=converte_numero(s_perc_multa)
				r("qtde_dias_protestar_apos_padrao")=CLng(s_qtde_dias_protestar_apos_padrao)
				r("segunda_mensagem_padrao")=s_segunda_mensagem_padrao
				r("mensagem_1_padrao")=s_mensagem_1_padrao
				r("mensagem_2_padrao")=s_mensagem_2_padrao
				r("mensagem_3_padrao")=s_mensagem_3_padrao
				r("mensagem_4_padrao")=s_mensagem_4_padrao
				r("endereco")=s_endereco
				r("endereco_numero")=s_endereco_numero
				r("endereco_complemento")=s_endereco_complemento
				r("bairro")=s_bairro
				r("cidade")=s_cidade
				r("cep")=s_cep
				r("uf")=s_uf
				r("dt_ult_atualizacao")=Now
				r("usuario_ult_atualizacao")=usuario
				
				r.Update

				If Err = 0 then 
					log_via_vetor_carrega_do_recordset r, vLog2, campos_a_omitir
					if criou_novo_reg then
						s_log = log_via_vetor_monta_inclusao(vLog2)
						if s_log <> "" then grava_log usuario, "", "", "", OP_LOG_BOLETO_CEDENTE_PARAMETROS_INCLUSAO, s_log
					else
						s_log = log_via_vetor_monta_alteracao(vLog1, vLog2)
						if s_log <> "" then 
							s_log="Id=" & Trim("" & r("id")) & "; " & s_log
							grava_log usuario, "", "", "", OP_LOG_BOLETO_CEDENTE_PARAMETROS_ALTERACAO, s_log
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
      CCCCCCC    SSSSSSS    SSSSSSS-->
<link href="<%=URL_FILE__E_CSS%>" rel="stylesheet" type="text/css">


<body onload="bVOLTAR.focus();">
<center>
<br>

<!--  T E L A  -->

<p class="T">A V I S O</p>

<% 
	s = ""
	s_aux="'MtAviso'"
	if alerta <> "" then
		s = "<p style='margin:5px 2px 5px 2px;'>" & alerta & "</p>"
		s_aux="'MtAlerta'"
	else
		select case operacao_selecionada
			case OP_INCLUI
				s = "REGISTRO ID=" & chr(34) & s_id_boleto_cedente_selecionado & chr(34) & " CADASTRADO COM SUCESSO."
			case OP_CONSULTA, OP_ALTERA
				s = "REGISTRO ID=" & chr(34) & s_id_boleto_cedente_selecionado & chr(34) & " ALTERADO COM SUCESSO."
			case OP_EXCLUI
				s = "REGISTRO ID=" & chr(34) & s_id_boleto_cedente_selecionado & chr(34) & " EXCLUÍDO COM SUCESSO."
			end select
		if s <> "" then s="<P style='margin:5px 2px 5px 2px;'>" & s & "</P>"
		end if
%>
<div class=<%=s_aux%> style="width:400px;font-weight:bold;" align="center"><%=s%></div>
<br><br>

<p class="TracoBottom"></p>

<table cellspacing="0">
<tr>
<%
	s="FinCadBoletoCedenteParametrosMenu.asp" & "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))
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