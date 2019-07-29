<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/estoque.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =====================================================
'	  P E D I D O D E V O L U C A O C O N F I R M A . A S P
'     =====================================================
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

	dim s, usuario, pedido_selecionado, id_pedido_base
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	if Not operacao_permitida(OP_CEN_CADASTRA_DEVOLUCAO_PRODUTO, s_lista_operacoes_permitidas) then 
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	pedido_selecionado = ucase(Trim(request("pedido_selecionado")))
	if (pedido_selecionado = "") then Response.Redirect("aviso.asp?id=" & ERR_PEDIDO_NAO_ESPECIFICADO)
	s = normaliza_num_pedido(pedido_selecionado)
	if s <> "" then pedido_selecionado = s
	
'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs, sx
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	dim vl_TotalFamiliaPrecoVenda, vl_TotalFamiliaPrecoNF, vl_TotalFamiliaPago, vl_TotalFamiliaDevolucaoPrecoVenda, vl_TotalFamiliaDevolucaoPrecoNF, vl_saldo_a_pagar, st_pagto, st_pagto_a
	dim iv, j, k, n, v_devol, alerta, deve_devolver, s_log, msg_erro, id_item_devolvido
	redim v_devol(0)
	set v_devol(Ubound(v_devol)) = New cl_ITEM_DEVOLUCAO_MERCADORIAS
	v_devol(Ubound(v_devol)).produto = ""
	v_devol(Ubound(v_devol)).qtde_a_devolver = 0
	v_devol(Ubound(v_devol)).motivo = ""
	
	s_log = ""
	alerta = ""
	deve_devolver = False

	n = Request.Form("c_qtde_devolucao").Count
	for iv = 1 to n
		s=Trim(Request.Form("c_produto")(iv))
		if s <> "" then
			if Trim(v_devol(Ubound(v_devol)).produto) <> "" then
				redim preserve v_devol(ubound(v_devol)+1)
				set v_devol(ubound(v_devol)) = New cl_ITEM_DEVOLUCAO_MERCADORIAS
				end if
			with v_devol(ubound(v_devol))
				.pedido=pedido_selecionado
				.produto=Ucase(Trim(Request.Form("c_produto")(iv)))
				
				s=retorna_so_digitos(Request.Form("c_fabricante")(iv))
				.fabricante=normaliza_codigo(s, TAM_MIN_FABRICANTE)
				
				s = Trim(Request.Form("c_qtde")(iv))
				if IsNumeric(s) then .qtde = CLng(s) else .qtde = 0
				
				s = Trim(Request.Form("c_devolucao_anterior")(iv))
				.qtde_devolvida_anteriormente = converte_numero(s) 
				
				s = Trim(Request.Form("c_qtde_devolucao")(iv))
				if IsNumeric(s) then .qtde_a_devolver = CLng(s) else .qtde_a_devolver = 0
				
				s = filtra_nome_identificador(Trim(Request.Form("c_motivo")(iv)))
				.motivo = s
				
				if .qtde_a_devolver > 0 then deve_devolver = True
				end with
			end if
		next

	if Not deve_devolver then
		alerta = "Não foi especificado nenhum produto para a operação de devolução."
	else
		for iv = Lbound(v_devol) to Ubound(v_devol)
			with v_devol(iv)
				if .produto <> "" then
					if .qtde_a_devolver > (.qtde - .qtde_devolvida_anteriormente) then
						alerta = texto_add_br(alerta)
						alerta = alerta & "Produto " & .produto & " do fabricante " & .fabricante & " especifica quantidade inválida para devolução."
						end if
					
					if (.qtde_a_devolver > 0) then
						if .motivo = "" then
							alerta = texto_add_br(alerta)
							alerta = alerta & "Não foi informado o motivo da devolução do produto " & .produto & " do fabricante " & .fabricante
							end if
						end if
					end if
				end with
			next
		end if
		
	if alerta = "" then
	'	~~~~~~~~~~~~~
		cn.BeginTrans
	'	~~~~~~~~~~~~~
		if Not cria_recordset_otimista(sx, msg_erro) then
		'	~~~~~~~~~~~~~~~~
			cn.RollbackTrans
		'	~~~~~~~~~~~~~~~~
			Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
			end if

		if Not cria_recordset_pessimista(rs, msg_erro) then
		'	~~~~~~~~~~~~~~~~
			cn.RollbackTrans
		'	~~~~~~~~~~~~~~~~
			Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
			end if

		for iv = Lbound(v_devol) to Ubound(v_devol)
			with v_devol(iv)
				if (Trim(.produto)<>"") And (.qtde_a_devolver > 0) then
				'	REGISTRA O PRODUTO DEVOLVIDO
					s = "SELECT * FROM t_PEDIDO_ITEM" & _
						" WHERE (pedido='" & .pedido & "')" & _
						" AND (fabricante='" & .fabricante & "')" & _
						" AND (produto='" & .produto & "')"
					if sx.State <> 0 then sx.Close
					sx.open s, cn
					if Err <> 0 then
					'	~~~~~~~~~~~~~~~~
						cn.RollbackTrans
					'	~~~~~~~~~~~~~~~~
						Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
						end if
		
					if sx.Eof then
						alerta = texto_add_br(alerta)
						alerta = alerta & "Item de pedido referente ao produto " & .produto & " do fabricante " & _
										  .fabricante & " não foi encontrado."
					else
						if Not gera_nsu(NSU_PEDIDO_ITEM_DEVOLVIDO, id_item_devolvido, msg_erro) then 
						'	~~~~~~~~~~~~~~~~
							cn.RollbackTrans
						'	~~~~~~~~~~~~~~~~
							Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_GERAR_NSU)
							end if

						s = "SELECT * FROM t_PEDIDO_ITEM_DEVOLVIDO WHERE (pedido='X')"
						if rs.State <> 0 then rs.Close
						rs.Open s, cn
						rs.AddNew
						for j = 0 to rs.Fields.Count-1
							for k = 0 to sx.Fields.Count-1
								if Ucase(rs.Fields(j).Name)=Ucase(sx.Fields(k).Name) then
									rs.Fields(j) = sx.Fields(k)
									exit for
									end if
								next
							next
						
						rs("id") = id_item_devolvido
						rs("qtde") = .qtde_a_devolver
						rs("motivo") = .motivo
						rs("devolucao_data") = Date
						rs("devolucao_hora") = retorna_so_digitos(formata_hora(Now))
						rs("devolucao_usuario") = usuario
						rs.Update 
						s_log = s_log & log_produto_monta(.qtde_a_devolver, .fabricante, .produto)
						if Err <> 0 then
						'	~~~~~~~~~~~~~~~~
							cn.RollbackTrans
						'	~~~~~~~~~~~~~~~~
							Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
							end if
						
						if Not estoque_processa_devolucao_mercadorias_v2(usuario, .pedido, .fabricante, .produto, id_item_devolvido, .qtde_a_devolver, msg_erro) then
						'	~~~~~~~~~~~~~~~~
							cn.RollbackTrans
						'	~~~~~~~~~~~~~~~~
							Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_MOVIMENTO_ESTOQUE)
							end if
						end if
					end if
				end with
			next
		
		
	'	ATUALIZA O STATUS DE PAGAMENTO (SE NECESSÁRIO)
	'	==============================================
		if alerta = "" then
		'	OBTÉM OS VALORES A PAGAR, JÁ PAGO E O STATUS DE PAGAMENTO (PARA TODA A FAMÍLIA DE PEDIDOS)
			if Not calcula_pagamentos(pedido_selecionado, vl_TotalFamiliaPrecoVenda, vl_TotalFamiliaPrecoNF, vl_TotalFamiliaPago, vl_TotalFamiliaDevolucaoPrecoVenda, vl_TotalFamiliaDevolucaoPrecoNF, st_pagto, msg_erro) then 
			'	~~~~~~~~~~~~~~~~
				cn.RollbackTrans
			'	~~~~~~~~~~~~~~~~
				Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
				end if
			
			vl_saldo_a_pagar = vl_TotalFamiliaPrecoNF-vl_TotalFamiliaPago-vl_TotalFamiliaDevolucaoPrecoNF
			
			id_pedido_base = retorna_num_pedido_base(pedido_selecionado)
			s = "SELECT * FROM t_PEDIDO WHERE (pedido='" & id_pedido_base & "')"
			if rs.State <> 0 then rs.Close
			rs.Open s, cn
			if Err <> 0 then
			'	~~~~~~~~~~~~~~~~
				cn.RollbackTrans
			'	~~~~~~~~~~~~~~~~
				Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
				end if
			
			if rs.Eof then
				alerta = texto_add_br(alerta)
				alerta = alerta & "Pedido-base " & id_pedido_base & " não foi encontrado."
			else
				st_pagto_a = Trim("" & rs("st_pagto"))
				if (vl_TotalFamiliaDevolucaoPrecoNF + vl_TotalFamiliaPago) >= (vl_TotalFamiliaPrecoNF - MAX_VALOR_MARGEM_ERRO_PAGAMENTO) then
					if Trim("" & rs("st_pagto")) <> ST_PAGTO_PAGO then
						rs("dt_st_pagto") = Date
						rs("dt_hr_st_pagto") = Now
						rs("usuario_st_pagto") = usuario
						end if
					rs("st_pagto") = ST_PAGTO_PAGO
				elseif vl_TotalFamiliaPago > 0 then
					if Trim("" & rs("st_pagto")) <> ST_PAGTO_PARCIAL then
						rs("dt_st_pagto") = Date
						rs("dt_hr_st_pagto") = Now
						rs("usuario_st_pagto") = usuario
						end if
					rs("st_pagto") = ST_PAGTO_PARCIAL
				else
					if Trim("" & rs("st_pagto")) <> ST_PAGTO_NAO_PAGO then
						rs("dt_st_pagto") = Date
						rs("dt_hr_st_pagto") = Now
						rs("usuario_st_pagto") = usuario
						end if
					rs("st_pagto") = ST_PAGTO_NAO_PAGO
					end if

				if st_pagto_a <> Trim("" & rs("st_pagto")) then
					s = formata_texto_log(Lcase(x_status_pagto(st_pagto_a))) & " => " & _
						formata_texto_log(Lcase(x_status_pagto(Trim("" & rs("st_pagto")))))
				else
					s = formata_texto_log(Lcase(x_status_pagto(Trim("" & rs("st_pagto")))))
					end if
				
				if s_log <> "" then s_log = s_log & ", "
				s_log = s_log & "st_pagto: " & s & ", " & _
						"valor do pedido: " & SIMBOLO_MONETARIO & " " & _
						formata_moeda(vl_TotalFamiliaPrecoNF) & ", " & _
						"valor pago: " & SIMBOLO_MONETARIO & " " & _
						formata_moeda(vl_TotalFamiliaPago) & ", " & _
						"valor das devoluções: " & SIMBOLO_MONETARIO & " " & _
						formata_moeda(vl_TotalFamiliaDevolucaoPrecoNF)

				rs.Update
				if Err <> 0 then
				'	~~~~~~~~~~~~~~~~
					cn.RollbackTrans
				'	~~~~~~~~~~~~~~~~
					Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
					end if
				end if
			end if
			
			
	'	GRAVA O LOG E CONCLUI A TRANSAÇÃO
	'	=================================
		if alerta = "" then
			s_log = "Devolução de mercadorias:" & s_log
			grava_log usuario, "", pedido_selecionado, "", OP_LOG_PEDIDO_DEVOLUCAO, s_log
		'	~~~~~~~~~~~~~~
			cn.CommitTrans
		'	~~~~~~~~~~~~~~
			if Err=0 then 
				s = "Devolução concluída com sucesso!!"
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
		
		if sx.State <> 0 then sx.Close
		set sx = nothing
		if rs.State <> 0 then rs.Close
		set rs = nothing
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
