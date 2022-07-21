<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp"        -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =====================================
'	  L O J A A T U A L I Z A . A S P
'     =====================================
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
	dim operacao_selecionada, s_loja, s_nome, s_razao_social, s_cnpj, s_ie
	dim s_endereco, s_endereco_numero, s_endereco_complemento, s_bairro, s_cidade, s_cep, s_uf
	dim s_ddd, s_telefone, s_fax, s_comissao_indicacao
	dim s_plano_contas_empresa, s_plano_contas_grupo, s_plano_contas_conta
	dim s_plano_contas_empresa_comissao_indicador, s_unidade_negocio
	operacao_selecionada=request("operacao_selecionada")
	s_loja=retorna_so_digitos(trim(request("loja_selecionada")))
	s_nome=Trim(request("nome"))
	s_razao_social=Trim(request("razao_social"))
	s_cnpj=retorna_so_digitos(request("cnpj"))
	s_ie=Trim(request("ie"))
	s_endereco=Trim(request("endereco"))
	s_endereco_numero=Trim(request("endereco_numero"))
	s_endereco_complemento=Trim(request("endereco_complemento"))
	s_bairro=Trim(request("bairro"))
	s_cidade=Trim(request("cidade"))
	s_cep=retorna_so_digitos(Trim(request("cep")))
	s_uf=Ucase(Trim(request("uf")))
	s_ddd=retorna_so_digitos(Trim(request("ddd")))
	s_telefone=retorna_so_digitos(Trim(request("telefone")))
	s_fax=retorna_so_digitos(Trim(request("fax")))
	s_comissao_indicacao=Trim(request("comissao_indicacao"))

	s_plano_contas_empresa=Trim(Request.Form("c_plano_contas_empresa"))
	s_plano_contas_grupo=""
	s_plano_contas_conta=""
	s=Trim(Request.Form("c_plano_contas_conta"))
	dim vAux
	if (s<>"") And (Instr(s,"|") > 0) then
		vAux=Split(s, "|")
		s_plano_contas_conta=vAux(Lbound(vAux))
		s_plano_contas_grupo=vAux(Ubound(vAux))
		end if

	s_plano_contas_empresa_comissao_indicador = Trim(Request.Form("c_plano_contas_empresa_comissao_indicador"))
	s_unidade_negocio = Trim(Request.Form("c_unidade_negocio"))

	if s_loja = "" then Response.Redirect("aviso.asp?id=" & ERR_ID_INVALIDO)
	
	s_loja=normaliza_codigo(s_loja, TAM_MIN_LOJA)

	dim erro_consistencia, erro_fatal
	
	erro_consistencia=false
	erro_fatal=false
	
	alerta = ""
	if s_loja = "" then
		alerta="NÚMERO DE LOJA INVÁLIDO."
	elseif s_nome = "" then
		alerta="PREENCHA O NOME (APELIDO) DA LOJA."
	elseif Not cnpj_ok(s_cnpj) then
		alerta="CNPJ INVÁLIDO."
	elseif Not cep_ok(s_cep) then
		alerta="CEP INVÁLIDO."
	elseif Not uf_ok(s_uf) then
		alerta="UF INVÁLIDA."
	elseif Not ddd_ok(s_ddd) then
		alerta="DDD INVÁLIDO."
	elseif Not telefone_ok(s_telefone) then
		alerta="TELEFONE INVÁLIDO."
	elseif Not telefone_ok(s_fax) then
		alerta="FAX INVÁLIDO."
	elseif (s_ddd <> "") And ((s_telefone = "") And (s_fax = "")) then
		alerta="PREENCHA O TELEFONE OU O Nº DO FAX."
	elseif (s_ddd = "") And ((s_telefone <> "") Or (s_fax <> "")) then
		alerta="PREENCHA O DDD."
	elseif (converte_numero(s_comissao_indicacao)<0) Or (converte_numero(s_comissao_indicacao)>100) then
		alerta="PERCENTUAL DE COMISSÃO POR INDICAÇÕES É INVÁLIDO."
	elseif s_plano_contas_empresa = "" then
		alerta="SELECIONE A EMPRESA DO PLANO DE CONTAS"
	elseif converte_numero(s_plano_contas_empresa)=0 then
		alerta="EMPRESA DO PLANO DE CONTAS É INVÁLIDA"
	elseif s_plano_contas_conta = "" then
		alerta="SELECIONE A CONTA DO PLANO DE CONTAS"
	elseif converte_numero(s_plano_contas_conta)=0 then
		alerta="CONTA DO PLANO DE CONTAS É INVÁLIDA"
	elseif s_plano_contas_grupo = "" then
		alerta="GRUPO DE CONTAS NÃO INFORMADO"
	elseif converte_numero(s_plano_contas_grupo)=0 then
		alerta="GRUPO DE CONTAS É INVÁLIDO"
		end if
	
	if alerta = "" then
		if (s_endereco<>"") Or (s_bairro<>"") Or (s_cidade<>"") Or (s_uf<>"") Or (s_cep<>"") then
			if s_endereco="" then
				alerta="PREENCHA O ENDEREÇO."
			elseif Len(s_endereco) > CLng(MAX_TAMANHO_CAMPO_ENDERECO) then
				alerta="ENDEREÇO EXCEDE O TAMANHO MÁXIMO PERMITIDO:<br>TAMANHO ATUAL: " & Cstr(Len(s_endereco)) & " CARACTERES<br>TAMANHO MÁXIMO: " & Cstr(MAX_TAMANHO_CAMPO_ENDERECO) & " CARACTERES"
			elseif s_endereco_numero="" then
				alerta="PREENCHA O NÚMERO DO ENDEREÇO."
			elseif s_cidade="" then
				alerta="PREENCHA A CIDADE DO ENDEREÇO."
			elseif s_uf="" then
				alerta="PREENCHA A UF DO ENDEREÇO."
			elseif s_cep="" then
				alerta="PREENCHA O CEP DO ENDEREÇO."
				end if
			end if
		end if
		
	if alerta <> "" then erro_consistencia=True	
		
	Err.Clear
	
	dim msg_erro
	if Not cria_recordset_otimista(r, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

'	EXECUTA OPERAÇÃO NO BD
	select case operacao_selecionada
		case OP_EXCLUI
		'	 =========
			s="SELECT COUNT(*) AS qtde FROM t_PEDIDO WHERE (loja = '" & s_loja & "')"
			r.Open s, cn
		'	ASSEGURA QUE A COMPARAÇÃO SERÁ FEITA ENTRE MESMO TIPO DE DADOS
			if Cstr(r("qtde")) > Cstr(0) then
				erro_fatal=True
				alerta = "LOJA NÃO PODE SER REMOVIDA PORQUE ESTÁ SENDO REFERENCIADA NA TABELA DE PEDIDOS."
				end if
			r.Close 
			
			if Not erro_fatal then
				s="SELECT COUNT(*) AS qtde FROM t_ORCAMENTO WHERE (loja = '" & s_loja & "')"
				r.Open s, cn
			'	ASSEGURA QUE A COMPARAÇÃO SERÁ FEITA ENTRE MESMO TIPO DE DADOS
				if Cstr(r("qtde")) > Cstr(0) then
					erro_fatal=True
					alerta = "LOJA NÃO PODE SER REMOVIDA PORQUE ESTÁ SENDO REFERENCIADA NA TABELA DE ORÇAMENTOS."
					end if
				r.Close 
				end if
				
			if Not erro_fatal then
				s="SELECT COUNT(*) AS qtde FROM t_USUARIO_X_LOJA WHERE (loja = '" & s_loja & "')"
				r.Open s, cn
			'	ASSEGURA QUE A COMPARAÇÃO SERÁ FEITA ENTRE MESMO TIPO DE DADOS
				if Cstr(r("qtde")) > Cstr(0) then
					erro_fatal=True
					alerta = "LOJA NÃO PODE SER REMOVIDA PORQUE ESTÁ SENDO REFERENCIADA NA TABELA DE USUÁRIOS."
					end if
				r.Close 
				end if

			if Not erro_fatal then
				s="SELECT COUNT(*) AS qtde FROM t_PEDIDO WHERE (loja_indicou = '" & s_loja & "')"
				r.Open s, cn
			'	ASSEGURA QUE A COMPARAÇÃO SERÁ FEITA ENTRE MESMO TIPO DE DADOS
				if Cstr(r("qtde")) > Cstr(0) then
					erro_fatal=True
					alerta = "LOJA NÃO PODE SER REMOVIDA PORQUE ESTÁ SENDO REFERENCIADA COMO " & _
							 chr(34) & "LOJA QUE INDICOU" & chr(34) & " NA TABELA DE PEDIDOS."
					end if
				r.Close 
				end if

			if Not erro_fatal then
				s="SELECT COUNT(*) AS qtde FROM t_ORCAMENTO WHERE (loja_indicou = '" & s_loja & "')"
				r.Open s, cn
			'	ASSEGURA QUE A COMPARAÇÃO SERÁ FEITA ENTRE MESMO TIPO DE DADOS
				if Cstr(r("qtde")) > Cstr(0) then
					erro_fatal=True
					alerta = "LOJA NÃO PODE SER REMOVIDA PORQUE ESTÁ SENDO REFERENCIADA COMO " & _
							 chr(34) & "LOJA QUE INDICOU" & chr(34) & " NA TABELA DE ORÇAMENTOS."
					end if
				r.Close 
				end if

			if Not erro_fatal then
				s="SELECT * FROM t_LOJA_GRUPO_ITEM WHERE (loja = '" & s_loja & "')"
				r.Open s, cn
			'	ASSEGURA QUE A COMPARAÇÃO SERÁ FEITA ENTRE MESMO TIPO DE DADOS
				if Not r.Eof then
					erro_fatal=True
					alerta = "LOJA NÃO PODE SER REMOVIDA PORQUE AINDA PERTENCE AO GRUPO DE LOJAS " & chr(34) & Trim("" & r("grupo")) & chr(34)
					end if
				r.Close 
				end if
			
			if Not erro_fatal then
			'	INFO P/ LOG
				s="SELECT * FROM t_LOJA WHERE loja = '" & s_loja & "'"
				r.Open s, cn
				if Not r.EOF then
					log_via_vetor_carrega_do_recordset r, vLog1, campos_a_omitir
					s_log = log_via_vetor_monta_exclusao(vLog1)
					end if
				r.Close
				
			'	APAGA!!
				s="DELETE FROM t_LOJA WHERE loja = '" & s_loja & "'"
				cn.Execute(s)
				If Err = 0 then 
					if s_log <> "" then grava_log usuario, "", "", "", OP_LOG_LOJA_EXCLUSAO, s_log
				else
					erro_fatal=True
					alerta = "FALHA AO REMOVER A LOJA (" & Cstr(Err) & ": " & Err.Description & ")."
					end if
				
				if Not erro_fatal then
					s="DELETE FROM t_PRODUTO_LOJA WHERE loja = '" & s_loja & "'"
					cn.Execute(s)
					If Err <> 0 then 
						erro_fatal=True
						alerta = "FALHA AO REMOVER A LISTA DE PRODUTOS DA LOJA (" & Cstr(Err) & ": " & Err.Description & ")."
						end if
					end if
				end if


		case OP_INCLUI, OP_CONSULTA
		'	 ======================
			if alerta = "" then 
				s = "SELECT * FROM t_LOJA WHERE loja = '" & s_loja & "'"
				r.Open s, cn
				if r.EOF then 
					r.AddNew 
					criou_novo_reg = True
					r("loja")=s_loja
					r("dt_cadastro") = Date
				else
					criou_novo_reg = False
					log_via_vetor_carrega_do_recordset r, vLog1, campos_a_omitir
					end if
					
				r("nome")=s_nome
				r("razao_social")=s_razao_social
				r("cnpj")=s_cnpj
				r("ie")=s_ie
				r("endereco")=s_endereco
				r("endereco_numero")=s_endereco_numero
				r("endereco_complemento")=s_endereco_complemento
				r("bairro")=s_bairro
				r("cidade")=s_cidade
				r("cep")=s_cep
				r("uf")=s_uf
				r("ddd")=s_ddd
				r("telefone")=s_telefone
				r("fax")=s_fax
				r("dt_ult_atualizacao")=Now
				r("comissao_indicacao")=converte_numero(s_comissao_indicacao)
				r("id_plano_contas_empresa")=CInt(s_plano_contas_empresa)
				r("id_plano_contas_grupo")=CInt(s_plano_contas_grupo)
				r("id_plano_contas_conta")=CLng(s_plano_contas_conta)
				r("natureza")=COD_FIN_NATUREZA__CREDITO
				
				if s_plano_contas_empresa_comissao_indicador <> "" then
					r("id_plano_contas_empresa_comissao_indicador") = CInt(s_plano_contas_empresa_comissao_indicador)
				else
					r("id_plano_contas_empresa_comissao_indicador") = Null
					end if

				if s_unidade_negocio <> "" then
					r("unidade_negocio") = s_unidade_negocio
				else
					r("unidade_negocio") = Null
					end if

				r.Update

				If Err = 0 then 
					log_via_vetor_carrega_do_recordset r, vLog2, campos_a_omitir
					if criou_novo_reg then
						s_log = log_via_vetor_monta_inclusao(vLog2)
						if s_log <> "" then grava_log usuario, "", "", "", OP_LOG_LOJA_INCLUSAO, s_log
					else
						s_log = log_via_vetor_monta_alteracao(vLog1, vLog2)
						if s_log <> "" then 
							s_log="loja=" & Trim("" & r("loja")) & "; " & s_log
							grava_log usuario, "", "", "", OP_LOG_LOJA_ALTERACAO, s_log
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
				s = "LOJA " & chr(34) & s_loja & chr(34) & " CADASTRADA COM SUCESSO."
			case OP_CONSULTA, OP_ALTERA
				s = "LOJA " & chr(34) & s_loja & chr(34) & " ALTERADA COM SUCESSO."
			case OP_EXCLUI
				s = "LOJA " & chr(34) & s_loja & chr(34) & " EXCLUÍDA COM SUCESSO."
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
	s="loja.asp" & "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))
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