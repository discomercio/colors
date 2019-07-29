<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =================================================================
'	  R E L A N A L I S E C R E D I T O C O N F I R M A . A S P
'     =================================================================
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

	dim blnEdicaoCampoObs1Bloqueado
	blnEdicaoCampoObs1Bloqueado = True
	
	dim s, usuario, msg_erro, i, n, v_dados(), opcao_filtro_pedido
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

	dim vLog1()
	dim vLog2()
	dim s_log
	dim campos_a_omitir
	s_log = ""
	campos_a_omitir = "|analise_credito_data|analise_credito_usuario|"

'	INICIALIZA VETOR
	redim v_dados(0)
	set v_dados(0) = New cl_SEIS_COLUNAS
	with v_dados(0)
		.c1 = ""
		.c2 = ""
		.c3 = ""
		.c4 = ""
		.c5 = ""
        .c6 = ""
		end with
	
'	OBTÉM DADOS DO FORMULÁRIO
	opcao_filtro_pedido = Ucase(Trim(request("opcao_filtro_pedido")))
	n = Request.Form("c_pedido").Count
	for i = 1 to n
		s=Trim(Request.Form("c_pedido")(i))
		if s <> "" then
			if Trim(v_dados(ubound(v_dados)).c1) <> "" then
				redim preserve v_dados(ubound(v_dados)+1)
				set v_dados(ubound(v_dados)) = New cl_SEIS_COLUNAS
				end if
			with v_dados(ubound(v_dados))
			'	IMPORTANTE: O ÍNDICE NO REQUEST.FORM VARIA DE 1 A COUNT(),
			'	MAS NO FORMULÁRIO VARIA DE ZERO A COUNT()-1!!!
			'	NO FORMULÁRIO, HÁ CAMPO(S) DO TIPO HIDDEN P/ FORÇAR
			'	A CRIAÇÃO DE ARRAY DE CAMPOS MESMO QUANDO HÁ APENAS 1 PEDIDO!!
				.c1=Trim(Request.Form("c_pedido")(i))
				.c2=Trim(Request.Form("c_obs1")(i))
				.c3=Trim(Request.Form("rb_credito_ped_" & Cstr(i-1)))
				.c4=Trim(Request.Form("c_descr_forma_pagto")(i))
				.c5=Trim(Request.Form("ckb_analise_endereco_" & Cstr(i-1)))
                .c6=Trim(Request.Form("c_pendente_vendas_motivo_" & Cstr(i-1)))
				end with
			end if
		next

	dim c_valor_inferior, c_valor_superior, c_lista_loja
	dim c_vendedor, c_indicador
	c_lista_loja = Trim(Request.Form("c_lista_loja"))
	c_valor_inferior = Trim(Request.Form("c_valor_inferior"))
	c_valor_superior = Trim(Request.Form("c_valor_superior"))
	c_vendedor = Trim(Request.Form("c_vendedor"))
	c_indicador = Trim(Request.Form("c_indicador"))
	
	dim alerta
	alerta=""

'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	
	if alerta = "" then
	'	~~~~~~~~~~~~~
		cn.BeginTrans
	'	~~~~~~~~~~~~~
		If Not cria_recordset_pessimista(rs, msg_erro) then 
		'	~~~~~~~~~~~~~~~~
			cn.RollbackTrans
		'	~~~~~~~~~~~~~~~~
			Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
			end if
        
    '   EM CASO DE 'PENDENTE VENDAS' VERIFICA SE O USUÁRIO SELECIONOU O MOTIVO
        for i=Lbound(v_dados) to Ubound(v_dados)
            if v_dados(i).c3 = COD_AN_CREDITO_PENDENTE_VENDAS then
                if v_dados(i).c6 = "" then
                    alerta=texto_add_br(alerta)
					alerta=alerta & "Não foi informado o motivo do status 'Pendente Vendas' do pedido " & v_dados(i).c1 & "."
                end if
            end if
        next

		for i=Lbound(v_dados) to Ubound(v_dados)
			if v_dados(i).c1 <> "" then
				s = "SELECT * FROM t_PEDIDO WHERE (pedido = '" & v_dados(i).c1 & "')"
				if rs.State <> 0 then rs.Close
				rs.Open s, cn
				if rs.Eof then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Pedido " & v_pedido(i) & " não foi encontrado."
				else
					log_via_vetor_carrega_do_recordset rs, vLog1, campos_a_omitir
				'	EDIÇÃO DO CAMPO "OBS_1" BLOQUEADO?
					if Not blnEdicaoCampoObs1Bloqueado then
						rs("obs_1")=v_dados(i).c2
						end if
						
					if IsNumeric(v_dados(i).c3) then 
						rs("analise_credito")=CLng(v_dados(i).c3)
						rs("analise_credito_data")=Now
						rs("analise_credito_usuario")=usuario
                        if v_dados(i).c3 = COD_AN_CREDITO_PENDENTE_VENDAS then rs("analise_credito_pendente_vendas_motivo") = v_dados(i).c6
						end if
					
					rs("forma_pagto")=v_dados(i).c4
					
					if Trim("" & v_dados(i).c5) = "AN_END_MARCAR_JA_TRATADO_OK" then
						if rs("analise_endereco_tratar_status") <> 0 then
							rs("analise_endereco_tratado_status") = CLng(COD_ANALISE_ENDERECO_TRATADO_STATUS_OK)
							rs("analise_endereco_tratado_data") = Date
							rs("analise_endereco_tratado_data_hora") = Now
							rs("analise_endereco_tratado_usuario") = usuario
							end if
						end if
					
					rs.Update
					if Err <> 0 then 
						alerta=texto_add_br(alerta)
						alerta=alerta & Cstr(Err) & ": " & Err.Description
					else
					'	INFORMAÇÕES PARA O LOG
						log_via_vetor_carrega_do_recordset rs, vLog2, campos_a_omitir
						s_log = log_via_vetor_monta_alteracao(vLog1, vLog2)
						if s_log <> "" then grava_log usuario, "", v_dados(i).c1, "", OP_LOG_ANALISE_CREDITO, s_log
						end if
					end if
				end if
			
		'	SE HOUVE ERRO, CANCELA O LAÇO
			if alerta <> "" then exit for
			next

		if alerta = "" then
		'	~~~~~~~~~~~~~~
			cn.CommitTrans
		'	~~~~~~~~~~~~~~
			if Err=0 then 
			'	EXIBE OS PRÓXIMOS "N" PEDIDOS P/ ANALISAR (SE HOUVER)
				if opcao_filtro_pedido = "S" then
					Response.Redirect("resumo.asp" & "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")))
				else
					Session("c_lista_loja") = c_lista_loja
					Session("c_valor_inferior") = c_valor_inferior
					Session("c_valor_superior") = c_valor_superior
					Session("c_vendedor") = c_vendedor
					Session("c_indicador") = c_indicador
					Response.Redirect("RelAnaliseCredito.asp?origem=A" & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")))
					end if
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
	if rs.State <> 0 then rs.Close
	set rs = nothing

'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>