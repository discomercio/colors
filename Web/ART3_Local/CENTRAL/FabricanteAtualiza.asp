<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp"        -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ===========================================
'	  F A B R I C A N T E A T U A L I Z A . A S P
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
	dim operacao_selecionada, s_fabricante, s_nome, s_razao_social, s_cnpj, s_ie
	dim s_endereco, s_endereco_numero, s_endereco_complemento, s_bairro, s_cidade, s_cep, s_uf, s_ddd, s_telefone, s_fax, s_contato, s_markup
	operacao_selecionada=request("operacao_selecionada")
	s_fabricante=retorna_so_digitos(trim(request("fabricante_selecionado")))
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
	s_contato=Trim(request("contato"))
	s_markup=Trim(request("markup"))

	if s_fabricante = "" then Response.Redirect("aviso.asp?id=" & ERR_ID_INVALIDO)
	
	s_fabricante=normaliza_codigo(s_fabricante, TAM_MIN_FABRICANTE)

	dim erro_consistencia, erro_fatal
	
	erro_consistencia=false
	erro_fatal=false
	
	alerta = ""
	if s_fabricante = "" then
		alerta="NÚMERO DE FABRICANTE INVÁLIDO."
	elseif s_nome = "" then
		alerta="PREENCHA O NOME (APELIDO) DO FABRICANTE."
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
	elseif (converte_numero(s_markup)<0) Or (converte_numero(s_markup)>100) then
		alerta="PERCENTUAL DE MARK UP É INVÁLIDO."
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
			s="SELECT COUNT(*) AS qtde FROM t_PEDIDO_ITEM WHERE (fabricante = '" & s_fabricante & "')"
			r.Open s, cn
		'	ASSEGURA QUE A COMPARAÇÃO SERÁ FEITA ENTRE MESMO TIPO DE DADOS
			if Cstr(r("qtde")) > Cstr(0) then
				erro_fatal=True
				alerta = "FABRICANTE NÃO PODE SER REMOVIDO PORQUE ESTÁ SENDO REFERENCIADO NA TABELA DE PEDIDOS."
				end if
			r.Close 
			
			if Not erro_fatal then
				s="SELECT COUNT(*) AS qtde FROM t_ORCAMENTO_ITEM WHERE (fabricante = '" & s_fabricante & "')"
				r.Open s, cn
			'	ASSEGURA QUE A COMPARAÇÃO SERÁ FEITA ENTRE MESMO TIPO DE DADOS
				if Cstr(r("qtde")) > Cstr(0) then
					erro_fatal=True
					alerta = "FABRICANTE NÃO PODE SER REMOVIDO PORQUE ESTÁ SENDO REFERENCIADO NA TABELA DE ORÇAMENTOS."
					end if
				r.Close
				end if
			
			if Not erro_fatal then
				s="SELECT COUNT(*) AS qtde FROM t_ESTOQUE WHERE (fabricante = '" & s_fabricante & "')"
				r.Open s, cn
			'	ASSEGURA QUE A COMPARAÇÃO SERÁ FEITA ENTRE MESMO TIPO DE DADOS
				if Cstr(r("qtde")) > Cstr(0) then
					erro_fatal=True
					alerta = "FABRICANTE NÃO PODE SER REMOVIDO PORQUE ESTÁ SENDO REFERENCIADO NA TABELA DE ESTOQUE."
					end if
				r.Close 
				end if
			
			if Not erro_fatal then
				s="SELECT COUNT(*) AS qtde FROM t_PRODUTO WHERE (fabricante = '" & s_fabricante & "')"
				r.Open s, cn
			'	ASSEGURA QUE A COMPARAÇÃO SERÁ FEITA ENTRE MESMO TIPO DE DADOS
				if Cstr(r("qtde")) > Cstr(0) then
					erro_fatal=True
					alerta = "FABRICANTE NÃO PODE SER REMOVIDO PORQUE ESTÁ SENDO REFERENCIADO NA TABELA DE PRODUTOS."
					end if
				r.Close 
				end if

			if Not erro_fatal then
				s="SELECT COUNT(*) AS qtde FROM t_PRODUTO_LOJA WHERE (fabricante = '" & s_fabricante & "')"
				r.Open s, cn
			'	ASSEGURA QUE A COMPARAÇÃO SERÁ FEITA ENTRE MESMO TIPO DE DADOS
				if Cstr(r("qtde")) > Cstr(0) then
					erro_fatal=True
					alerta = "FABRICANTE NÃO PODE SER REMOVIDO PORQUE ESTÁ SENDO REFERENCIADO NA TABELA DE PRODUTOS (POR LOJA)."
					end if
				r.Close 
				end if
			
			if Not erro_fatal then
			'	INFO P/ LOG
				s="SELECT * FROM t_FABRICANTE WHERE fabricante = '" & s_fabricante & "'"
				r.Open s, cn
				if Not r.EOF then
					log_via_vetor_carrega_do_recordset r, vLog1, campos_a_omitir
					s_log = log_via_vetor_monta_exclusao(vLog1)
					end if
				r.Close
			
			'	APAGA!!
				s="DELETE FROM t_FABRICANTE WHERE fabricante = '" & s_fabricante & "'"
				cn.Execute(s)
				If Err = 0 then 
					if s_log <> "" then grava_log usuario, "", "", "", OP_LOG_FABRICANTE_EXCLUSAO, s_log
				else
					erro_fatal=True
					alerta = "FALHA AO REMOVER O FABRICANTE (" & Cstr(Err) & ": " & Err.Description & ")."
					end if
				end if


		case OP_INCLUI, OP_CONSULTA
		'	 ======================
			if alerta = "" then 
				s = "SELECT * FROM t_FABRICANTE WHERE fabricante = '" & s_fabricante & "'"
				r.Open s, cn
				if r.EOF then 
					r.AddNew 
					criou_novo_reg = True
					r("fabricante")=s_fabricante
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
				r("contato")=s_contato
				r("dt_ult_atualizacao")=Now
				r("markup")=converte_numero(s_markup)
				
				r.Update

				If Err = 0 then 
					log_via_vetor_carrega_do_recordset r, vLog2, campos_a_omitir
					if criou_novo_reg then
						s_log = log_via_vetor_monta_inclusao(vLog2)
						if s_log <> "" then grava_log usuario, "", "", "", OP_LOG_FABRICANTE_INCLUSAO, s_log
					else
						s_log = log_via_vetor_monta_alteracao(vLog1, vLog2)
						if s_log <> "" then 
							s_log="fabricante=" & Trim("" & r("fabricante")) & "; " & s_log
							grava_log usuario, "", "", "", OP_LOG_FABRICANTE_ALTERACAO, s_log
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
				s = "FABRICANTE " & chr(34) & s_fabricante & chr(34) & " CADASTRADO COM SUCESSO."
			case OP_CONSULTA, OP_ALTERA
				s = "FABRICANTE " & chr(34) & s_fabricante & chr(34) & " ALTERADO COM SUCESSO."
			case OP_EXCLUI
				s = "FABRICANTE " & chr(34) & s_fabricante & chr(34) & " EXCLUÍDO COM SUCESSO."
			end select
		if s <> "" then s="<p style='margin:5px 2px 5px 2px;'>" & s & "</p>"
		end if
%>
<div class=<%=s_aux%> style="width:400px;font-weight:bold;" align="center"><%=s%></div>
<br><br>

<p class="TracoBottom"></p>

<table cellspacing="0">
<tr>
<%
	s="fabricante.asp" & "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))
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