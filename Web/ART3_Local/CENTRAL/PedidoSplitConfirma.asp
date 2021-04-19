<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/estoque.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =============================================
'	  P E D I D O S P L I T C O N F I R M A . A S P
'     =============================================
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

	dim s, usuario, pedido_selecionado
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

	pedido_selecionado = ucase(Trim(request("pedido_selecionado")))
	if (pedido_selecionado = "") then Response.Redirect("aviso.asp?id=" & ERR_PEDIDO_NAO_ESPECIFICADO)
	s = normaliza_num_pedido(pedido_selecionado)
	if s <> "" then pedido_selecionado = s
	
'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs, sx
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	dim blnUsarMemorizacaoCompletaEnderecos
	blnUsarMemorizacaoCompletaEnderecos = isActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos

	dim sBlocoNotasMsg
	dim total_estoque_sem_presenca, total_estoque_vendido
	dim i, n, v_split, alerta, deve_splitar, id_pedido_filhote, s_log, msg_erro
	redim v_split(0)
	set v_split(Ubound(v_split)) = New cl_ITEM_SPLIT
	v_split(Ubound(v_split)).produto = ""
	v_split(Ubound(v_split)).qtde_split = 0

	dim v_item, total_a_splitar, total_produtos
	total_a_splitar = 0
	total_produtos = 0
	
	s_log = ""
	alerta = ""
	deve_splitar = False
	n = Request.Form("c_qtde_split").Count
	for i = 1 to n
		s=Trim(Request.Form("c_produto")(i))
		if s <> "" then
			if Trim(v_split(Ubound(v_split)).produto) <> "" then
				redim preserve v_split(ubound(v_split)+1)
				set v_split(ubound(v_split)) = New cl_ITEM_SPLIT
				end if
			with v_split(ubound(v_split))
				.pedido=pedido_selecionado
				.produto=Ucase(Trim(Request.Form("c_produto")(i)))
				s=retorna_so_digitos(Request.Form("c_fabricante")(i))
				.fabricante=normaliza_codigo(s, TAM_MIN_FABRICANTE)
				s = Trim(Request.Form("c_qtde")(i))
				if IsNumeric(s) then .qtde = CLng(s) else .qtde = 0
				s = Trim(Request.Form("c_qtde_falta")(i))
				.qtde_estoque_vendido = .qtde - converte_numero(s) 
				s = Trim(Request.Form("c_qtde_split")(i))
				if IsNumeric(s) then .qtde_split = CLng(s) else .qtde_split = 0
				if .qtde_split > 0 then 
					deve_splitar = True
					total_a_splitar = total_a_splitar + .qtde_split
					end if
				end with
			end if
		next

	if Not deve_splitar then
		alerta = "Não foi especificado nenhum produto para a operação de split."
	else
		for i = Lbound(v_split) to Ubound(v_split)
			with v_split(i)
				if .produto <> "" then
					if .qtde_split > .qtde_estoque_vendido then
						alerta = texto_add_br(alerta)
						alerta = alerta & "Produto " & .produto & " do fabricante " & .fabricante & " especifica quantidade inválida para split."
						end if
					end if
				end with
			next
		end if
	
	if alerta = "" then
		if Not le_pedido_item(pedido_selecionado, v_item, msg_erro) then 
			alerta = msg_erro
		else
			for i=Lbound(v_item) to Ubound(v_item)
				with v_item(i)
					if Trim("" & .produto)<>"" then
						if .qtde > 0 then total_produtos = total_produtos + .qtde
						end if
					end with
				next
			
			if (total_produtos - total_a_splitar) <= 0 then
				alerta = "É necessário que sobre pelo menos um produto no pedido original."
				end if
			end if
		end if
		
	if alerta = "" then
	'	~~~~~~~~~~~~~
		cn.BeginTrans
	'	~~~~~~~~~~~~~
		if Not gera_num_pedido_filhote(pedido_selecionado, id_pedido_filhote, msg_erro) then 
		'	~~~~~~~~~~~~~~~~
			cn.RollbackTrans
		'	~~~~~~~~~~~~~~~~
			Response.Redirect("aviso.asp?id=" & ERR_FALHA_GERAR_ID_PEDIDO_FILHOTE)
			end if
		
		if Not cria_recordset_pessimista(rs, msg_erro) then
		'	~~~~~~~~~~~~~~~~
			cn.RollbackTrans
		'	~~~~~~~~~~~~~~~~
			Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
			end if

	'	CRIA O PEDIDO FILHOTE
		s = "SELECT * FROM t_PEDIDO WHERE (pedido='" & pedido_selecionado & "')"
		set sx = cn.execute(s)
		if Err <> 0 then
		'	~~~~~~~~~~~~~~~~
			cn.RollbackTrans
		'	~~~~~~~~~~~~~~~~
			Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
			end if
		
		if sx.Eof then
			alerta = "Pedido " & pedido_selecionado & " não foi encontrado."
		else
			s = "SELECT * FROM t_PEDIDO WHERE (pedido='X')"
			rs.Open s, cn
			rs.AddNew
			rs("pedido") = id_pedido_filhote
			rs("loja") = Trim("" & sx("loja"))
			rs("data") = sx("data")
			rs("hora") = Trim("" & sx("hora"))
			rs("split_status") = 1
			rs("split_data") = Date
			rs("split_hora") = retorna_so_digitos(formata_hora(Now))
			rs("split_usuario") = usuario
			rs("id_cliente") = Trim("" & sx("id_cliente"))
			rs("servicos") = ""
			rs("vendedor") = Trim("" & sx("vendedor"))
			rs("st_orc_virou_pedido") = sx("st_orc_virou_pedido")
			rs("orcamento") = Trim("" & sx("orcamento"))
			rs("orcamentista") = Trim("" & sx("orcamentista"))
			rs("indicador") = Trim("" & sx("indicador"))
			rs("st_entrega") = ST_ENTREGA_SEPARAR
			rs("st_pagto") = ""
			rs("usuario_st_pagto") = ""
			rs("st_recebido") = ""
			rs("obs_1") = ""
			rs("obs_2") = ""
			rs("qtde_parcelas") = 0
			rs("forma_pagto") = ""
			rs("midia") = Trim("" & sx("midia"))
			rs("loja_indicou") = Trim("" & sx("loja_indicou"))
			rs("comissao_loja_indicou") = sx("comissao_loja_indicou")
			rs("venda_externa") = sx("venda_externa")
			
			rs("st_end_entrega") = sx("st_end_entrega")
			rs("EndEtg_endereco") = sx("EndEtg_endereco")
			rs("EndEtg_endereco_numero") = sx("EndEtg_endereco_numero")
			rs("EndEtg_endereco_complemento") = sx("EndEtg_endereco_complemento")
			rs("EndEtg_bairro") = sx("EndEtg_bairro")
			rs("EndEtg_cidade") = sx("EndEtg_cidade")
			rs("EndEtg_uf") = sx("EndEtg_uf")
			rs("EndEtg_cep") = sx("EndEtg_cep")
			rs("EndEtg_cod_justificativa") = sx("EndEtg_cod_justificativa")
			if blnUsarMemorizacaoCompletaEnderecos then
				rs("EndEtg_email") = sx("EndEtg_email")
				rs("EndEtg_email_xml") = sx("EndEtg_email_xml")
				rs("EndEtg_nome") = sx("EndEtg_nome")
				rs("EndEtg_ddd_res") = sx("EndEtg_ddd_res")
				rs("EndEtg_tel_res") = sx("EndEtg_tel_res")
				rs("EndEtg_ddd_com") = sx("EndEtg_ddd_com")
				rs("EndEtg_tel_com") = sx("EndEtg_tel_com")
				rs("EndEtg_ramal_com") = sx("EndEtg_ramal_com")
				rs("EndEtg_ddd_cel") = sx("EndEtg_ddd_cel")
				rs("EndEtg_tel_cel") = sx("EndEtg_tel_cel")
				rs("EndEtg_ddd_com_2") = sx("EndEtg_ddd_com_2")
				rs("EndEtg_tel_com_2") = sx("EndEtg_tel_com_2")
				rs("EndEtg_ramal_com_2") = sx("EndEtg_ramal_com_2")
				rs("EndEtg_tipo_pessoa") = sx("EndEtg_tipo_pessoa")
				rs("EndEtg_cnpj_cpf") = sx("EndEtg_cnpj_cpf")
				rs("EndEtg_contribuinte_icms_status") = sx("EndEtg_contribuinte_icms_status")
				rs("EndEtg_produtor_rural_status") = sx("EndEtg_produtor_rural_status")
				rs("EndEtg_ie") = sx("EndEtg_ie")
				rs("EndEtg_rg") = sx("EndEtg_rg")
				end if

			rs("st_etg_imediata") = sx("st_etg_imediata")
			rs("etg_imediata_data") = sx("etg_imediata_data")
			rs("etg_imediata_usuario") = sx("etg_imediata_usuario")
			
			rs("PagtoAntecipadoQuitadoStatus") = sx("PagtoAntecipadoQuitadoStatus")
			rs("PagtoAntecipadoQuitadoDataHora") = sx("PagtoAntecipadoQuitadoDataHora")
			rs("PagtoAntecipadoQuitadoUsuario") = sx("PagtoAntecipadoQuitadoUsuario")

			rs("pedido_bs_x_ac") = sx("pedido_bs_x_ac")
			rs("pedido_bs_x_marketplace") = sx("pedido_bs_x_marketplace")
			rs("marketplace_codigo_origem") = sx("marketplace_codigo_origem")

			rs("GarantiaIndicadorStatus") = sx("GarantiaIndicadorStatus")
			rs("GarantiaIndicadorUsuarioUltAtualiz") = sx("GarantiaIndicadorUsuarioUltAtualiz")
			rs("GarantiaIndicadorDtHrUltAtualiz") = sx("GarantiaIndicadorDtHrUltAtualiz")
			
			rs("perc_desagio_RA_liquida") = sx("perc_desagio_RA_liquida")
			
			rs("permite_RA_status") = sx("permite_RA_status")
			rs("st_violado_permite_RA_status") = sx("st_violado_permite_RA_status")
			rs("dt_hr_violado_permite_RA_status") = sx("dt_hr_violado_permite_RA_status")
			rs("usuario_violado_permite_RA_status") = sx("usuario_violado_permite_RA_status")
			rs("opcao_possui_RA") = sx("opcao_possui_RA")
			
			if CInt(sx("transportadora_selecao_auto_status")) <> 0 then
				rs("transportadora_id") = sx("transportadora_id")
				rs("transportadora_selecao_auto_status") = sx("transportadora_selecao_auto_status")
				rs("transportadora_selecao_auto_cep") = sx("transportadora_selecao_auto_cep")
				rs("transportadora_selecao_auto_tipo_endereco") = sx("transportadora_selecao_auto_tipo_endereco")
				rs("transportadora_selecao_auto_transportadora") = sx("transportadora_selecao_auto_transportadora")
				rs("transportadora_selecao_auto_data_hora") = sx("transportadora_selecao_auto_data_hora")
				end if
			
			rs("id_nfe_emitente") = sx("id_nfe_emitente")
			rs("usuario_cadastro") = sx("usuario_cadastro")
			
			rs("plataforma_origem_pedido") = sx("plataforma_origem_pedido")

			rs("endereco_memorizado_status") = sx("endereco_memorizado_status")
			rs("endereco_logradouro") = sx("endereco_logradouro")
			rs("endereco_numero") = sx("endereco_numero")
			rs("endereco_complemento") = sx("endereco_complemento")
			rs("endereco_bairro") = sx("endereco_bairro")
			rs("endereco_cidade") = sx("endereco_cidade")
			rs("endereco_uf") = sx("endereco_uf")
			rs("endereco_cep") = sx("endereco_cep")

			if blnUsarMemorizacaoCompletaEnderecos then
				rs("st_memorizacao_completa_enderecos") = sx("st_memorizacao_completa_enderecos")
				rs("endereco_email") = sx("endereco_email")
				rs("endereco_email_xml") = sx("endereco_email_xml")
				rs("endereco_nome") = sx("endereco_nome")
				rs("endereco_ddd_res") = sx("endereco_ddd_res")
				rs("endereco_tel_res") = sx("endereco_tel_res")
				rs("endereco_ddd_com") = sx("endereco_ddd_com")
				rs("endereco_tel_com") = sx("endereco_tel_com")
				rs("endereco_ramal_com") = sx("endereco_ramal_com")
				rs("endereco_ddd_cel") = sx("endereco_ddd_cel")
				rs("endereco_tel_cel") = sx("endereco_tel_cel")
				rs("endereco_ddd_com_2") = sx("endereco_ddd_com_2")
				rs("endereco_tel_com_2") = sx("endereco_tel_com_2")
				rs("endereco_ramal_com_2") = sx("endereco_ramal_com_2")
				rs("endereco_tipo_pessoa") = sx("endereco_tipo_pessoa")
				rs("endereco_cnpj_cpf") = sx("endereco_cnpj_cpf")
				rs("endereco_contribuinte_icms_status") = sx("endereco_contribuinte_icms_status")
				rs("endereco_produtor_rural_status") = sx("endereco_produtor_rural_status")
				rs("endereco_ie") = sx("endereco_ie")
				rs("endereco_rg") = sx("endereco_rg")
				rs("endereco_contato") = sx("endereco_contato")
				end if

			rs("StBemUsoConsumo") = sx("StBemUsoConsumo")
			rs("InstaladorInstalaStatus") = sx("InstaladorInstalaStatus")
			rs("InstaladorInstalaUsuarioUltAtualiz") = sx("InstaladorInstalaUsuarioUltAtualiz")
			rs("InstaladorInstalaDtHrUltAtualiz") = sx("InstaladorInstalaDtHrUltAtualiz")

			rs("sistema_responsavel_cadastro") = sx("sistema_responsavel_cadastro")
			rs("sistema_responsavel_atualizacao") = COD_SISTEMA_RESPONSAVEL_CADASTRO__ERP
			
			rs.Update
			if Err <> 0 then
			'	~~~~~~~~~~~~~~~~
				cn.RollbackTrans
			'	~~~~~~~~~~~~~~~~
				Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
				end if
				

			for i=Lbound(v_split) to Ubound(v_split)
				with v_split(i)
					if (.produto <> "") And (.qtde_split > 0) then

					'	TRANSFERE OS PRODUTOS DO ESTOQUE "VENDIDO"
						if Not estoque_produto_split_v2(usuario, pedido_selecionado, id_pedido_filhote, .fabricante, .produto, .qtde_split, msg_erro) then
						'	~~~~~~~~~~~~~~~~
							cn.RollbackTrans
						'	~~~~~~~~~~~~~~~~
							Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_MOVIMENTO_ESTOQUE)
							end if
						
						s_log = s_log & log_produto_monta(.qtde_split, .fabricante, .produto)

					'	CRIA O ITEM DO PEDIDO FILHOTE
						s = "SELECT * FROM t_PEDIDO_ITEM WHERE" & _
							" (pedido='" & pedido_selecionado & "')" & _
							" AND (fabricante='" & .fabricante & "')" & _
							" AND (produto='" & .produto & "')"
						if sx.State <> 0 then sx.Close
						set sx=cn.execute(s)
						if Err <> 0 then
						'	~~~~~~~~~~~~~~~~
							cn.RollbackTrans
						'	~~~~~~~~~~~~~~~~
							Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
							end if
						if sx.Eof then
							alerta = texto_add_br(alerta)
							alerta = alerta & "Item do pedido " & pedido_selecionado & " não foi encontrado: produto " & .produto & " do fabricante " & .fabricante & "."
							exit for
						else
							if rs.State <> 0 then rs.Close
							s = "SELECT * FROM t_PEDIDO_ITEM WHERE (pedido='X') AND (fabricante='X') AND (produto='X')"
							rs.Open s, cn
							rs.AddNew
							for n=0 to rs.Fields.Count-1
								rs.Fields(n).Value = sx.Fields(n).Value
								next
							
							rs("pedido") = id_pedido_filhote
							rs("qtde") = .qtde_split
							rs.Update 
							if Err <> 0 then
							'	~~~~~~~~~~~~~~~~
								cn.RollbackTrans
							'	~~~~~~~~~~~~~~~~
								Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
								end if
							
						'	ACERTA A QUANTIDADE NO PEDIDO ORIGINAL
							if rs.State <> 0 then rs.Close
							s = "SELECT * FROM t_PEDIDO_ITEM WHERE" & _
								" (pedido='" & pedido_selecionado & "')" & _
								" AND (fabricante='" & .fabricante & "')" & _
								" AND (produto='" & .produto & "')"
							rs.Open s, cn
							if Err <> 0 then
							'	~~~~~~~~~~~~~~~~
								cn.RollbackTrans
							'	~~~~~~~~~~~~~~~~
								Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
								end if
							
							if rs.EOF then
								alerta = texto_add_br(alerta)
								alerta = alerta & "Item do pedido " & pedido_selecionado & " não foi encontrado: produto " & .produto & " do fabricante " & .fabricante & "."
								exit for
							else
								rs("qtde") = rs("qtde") - .qtde_split
								if rs("qtde") > 0 then
									rs.Update 
								else
									rs.Delete
									end if
								end if
							end if
						end if
					end with
				next
			end if
		
		if alerta = "" then
			sBlocoNotasMsg = "Pedido gerado através de split manual do pedido " & pedido_selecionado & " por '" & usuario & "'"
			if Not grava_bloco_notas_pedido(id_pedido_filhote, ID_USUARIO_SISTEMA, "", COD_NIVEL_ACESSO_BLOCO_NOTAS_PEDIDO__RESTRITO, sBlocoNotasMsg, COD_TIPO_MSG_BLOCO_NOTAS_PEDIDO__AUTOMATICA_SPLIT_MANUAL, msg_erro) then
				alerta = "Falha ao gravar bloco de notas com mensagem automática no pedido (" & id_pedido_filhote & ")"
				end if
			end if

		if alerta = "" then
		'	PEDIDO-BASE: ATUALIZA STATUS DE ENTREGA
			total_estoque_sem_presenca = 0
			s = "SELECT Sum(qtde) AS total FROM t_ESTOQUE_MOVIMENTO" & _
				" WHERE (anulado_status=0)" & _
				" AND (estoque = '" & ID_ESTOQUE_SEM_PRESENCA & "')" & _
				" AND (pedido = '" & pedido_selecionado & "')"
			if sx.State <> 0 then sx.Close
			set sx=cn.execute(s)
			if Err <> 0 then
			'	~~~~~~~~~~~~~~~~
				cn.RollbackTrans
			'	~~~~~~~~~~~~~~~~
				Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
				end if
			if Not sx.Eof then
				if IsNumeric(sx("total")) then total_estoque_sem_presenca = CLng(sx("total"))
				end if
			
			total_estoque_vendido = 0
			s = "SELECT Sum(qtde) AS total FROM t_ESTOQUE_MOVIMENTO" & _
				" WHERE (anulado_status=0)" & _
				" AND (estoque = '" & ID_ESTOQUE_VENDIDO & "')" & _
				" AND (pedido = '" & pedido_selecionado & "')"
			if sx.State <> 0 then sx.Close
			set sx=cn.execute(s)
			if Err <> 0 then
			'	~~~~~~~~~~~~~~~~
				cn.RollbackTrans
			'	~~~~~~~~~~~~~~~~
				Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
				end if
			if Not sx.Eof then
				if IsNumeric(sx("total")) then total_estoque_vendido = CLng(sx("total"))
				end if
			
			s = "SELECT * FROM t_PEDIDO WHERE (pedido = '" & pedido_selecionado & "')"
			if rs.State <> 0 then rs.Close
			rs.open s, cn
			if rs.Eof then
			'	~~~~~~~~~~~~~~~~
				cn.RollbackTrans
			'	~~~~~~~~~~~~~~~~
				Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
			else
			'	STATUS DE ENTREGA
				if total_estoque_vendido = 0 then
					s = ST_ENTREGA_ESPERAR
				elseif total_estoque_sem_presenca = 0 then
					s = ST_ENTREGA_SEPARAR
				else
					s = ST_ENTREGA_SPLIT_POSSIVEL
					end if
						
				if Trim("" & rs("st_entrega")) <> s then
					rs("st_entrega") = s
					rs.Update
					if Err <> 0 then
					'	~~~~~~~~~~~~~~~~
						cn.RollbackTrans
					'	~~~~~~~~~~~~~~~~
						Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
						end if
					end if
				end if
			
			s_log = "Filhote de pedido nº " & id_pedido_filhote & " criado com:" & s_log
			grava_log usuario, "", pedido_selecionado, "", OP_LOG_PEDIDO_SPLIT, s_log
		'	~~~~~~~~~~~~~~
			cn.CommitTrans
		'	~~~~~~~~~~~~~~
			if Err=0 then 
				Response.Redirect("pedido.asp?pedido_selecionado=" & id_pedido_filhote & "&url_back=X" & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")))
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
