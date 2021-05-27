<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/estoque.asp" -->
<!-- #include file = "../global/Global.asp"    -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ===========================================
'	  P E D I D O C A N C E L A . A S P
'     ===========================================
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

	pedido_selecionado = Trim(request.Form("pedido_selecionado"))
	if (pedido_selecionado = "") then Response.Redirect("aviso.asp?id=" & ERR_PEDIDO_NAO_ESPECIFICADO)
	
'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

	dim id_nfe_emitente
	dim vl_TotalFamiliaPrecoVenda, vl_TotalFamiliaPrecoNF, vl_TotalFamiliaPago, vl_TotalFamiliaDevolucaoPrecoVenda, vl_TotalFamiliaDevolucaoPrecoNF, st_pagto, id_pedido_base
	dim vl_total_RA, vl_total_RA_liquido,motivo_cancelamento,c_cod_motivo,c_cod_sub_motivo
	dim alerta, s_log, msg_erro
	alerta=""
	motivo_cancelamento = Request("motivo_cancelamento")
    c_cod_motivo = Request("c_cod_motivo")
    c_cod_sub_motivo = Request("c_cod_sub_motivo")
'	VERIFICA SE A NFe JÁ FOI EMITIDA
	dim strSerieNFe, strNumeroNFe
	if alerta = "" then
		if Not operacao_permitida(OP_CEN_CANCELAR_PEDIDO_COM_NFE_EMITIDA, s_lista_operacoes_permitidas) then
			if Not IsPedidoCancelavelNFeEmitida(pedido_selecionado, strSerieNFe, strNumeroNFe, msg_erro) then
				alerta = "Não é possível cancelar o pedido " & pedido_selecionado & " porque a NFe nº " & strNumeroNFe & " não está cancelada!!"
				end if
			end if
		end if

	if alerta = "" then
	'	~~~~~~~~~~~~~
		cn.BeginTrans
	'	~~~~~~~~~~~~~
		if TRATAMENTO_ACESSO_CONCORRENTE_LOCK_EXCLUSIVO_MANUAL_HABILITADO then
		'	BLOQUEIA REGISTRO PARA EVITAR ACESSO CONCORRENTE (REALIZA O FLIP EM UM CAMPO BIT APENAS P/ ADQUIRIR O LOCK EXCLUSIVO)
		'	OBS: TODOS OS MÓDULOS DO SISTEMA QUE REALIZEM ESTA OPERAÇÃO DE CADASTRAMENTO DEVEM SINCRONIZAR O ACESSO OBTENDO O LOCK EXCLUSIVO DO REGISTRO DE CONTROLE DESIGNADO
			s = "UPDATE t_CONTROLE SET" & _
					" dummy = ~dummy" & _
				" WHERE" & _
					" id_nsu = '" & ID_XLOCK_SYNC_PEDIDO & "'"
			cn.Execute(s)
			end if

		if Not cria_recordset_pessimista(rs, msg_erro) then
		'	~~~~~~~~~~~~~~~~
			cn.RollbackTrans
		'	~~~~~~~~~~~~~~~~
			Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
			end if

		s = "SELECT * FROM t_PEDIDO WHERE (pedido='" & pedido_selecionado & "')"
		rs.Open s, cn
		if rs.EOF then
			alerta="Pedido " & pedido_selecionado & " não está cadastrado."
		else
			id_nfe_emitente = rs("id_nfe_emitente")
			s = Trim("" & rs("st_entrega"))
			if s = ST_ENTREGA_CANCELADO then
				alerta="Pedido " & pedido_selecionado & " já foi cancelado em " & formata_data(rs("cancelado_data")) & "."
			elseif s = ST_ENTREGA_ENTREGUE then
				alerta="Pedido " & pedido_selecionado & " já foi entregue em " & formata_data(rs("entregue_data")) & " e não pode mais ser cancelado."
			elseif Trim("" & rs("obs_2")) <> "" then
				alerta="Pedido " & pedido_selecionado & " não pode ser cancelado enquanto o campo 'Observações II' estiver preenchido."
			elseif (rs("transportadora_selecao_auto_status") = 0) And (Trim("" & rs("transportadora_id")) <> "") then
				alerta="Pedido " & pedido_selecionado & " não pode ser cancelado enquanto o campo 'Transportadora' estiver preenchido."
			else
			'	CANCELA O PEDIDO
			'	================
				rs("st_entrega") = ST_ENTREGA_CANCELADO
				rs("cancelado_data") = Date
				rs("cancelado_usuario") = usuario
                rs("cancelado_data_hora") = Now
                rs("cancelado_motivo") = motivo_cancelamento
                rs("cancelado_codigo_motivo") = c_cod_motivo
                if c_cod_sub_motivo <> "" then rs("cancelado_codigo_sub_motivo") = c_cod_sub_motivo
				rs("sistema_responsavel_atualizacao") = COD_SISTEMA_RESPONSAVEL_CADASTRO__ERP
				rs.Update
				if Err <> 0 then
				'	~~~~~~~~~~~~~~~~
					cn.RollbackTrans
				'	~~~~~~~~~~~~~~~~
					Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
					end if
				rs.Close 
				
			'	PRODUTOS: ESTORNA DO ESTOQUE VENDIDO DE VOLTA P/ O ESTOQUE DE VENDA E
			'	========= CANCELA LISTA DE PRODUTOS VENDIDOS SEM PRESENÇA NO ESTOQUE.
				if Not estoque_pedido_cancela(usuario, pedido_selecionado, s_log, msg_erro) then
				'	~~~~~~~~~~~~~~~~
					cn.RollbackTrans
				'	~~~~~~~~~~~~~~~~
					Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_MOVIMENTO_ESTOQUE)
					end if
				
				grava_log usuario, "", pedido_selecionado, "", OP_LOG_PEDIDO_CANCELA, s_log
				
			'	PROCESSA OS PRODUTOS VENDIDOS SEM PRESENÇA NO ESTOQUE
				if Not estoque_processa_produtos_vendidos_sem_presenca_v2(id_nfe_emitente, usuario, msg_erro) then
				'	~~~~~~~~~~~~~~~~
					cn.RollbackTrans
				'	~~~~~~~~~~~~~~~~
					Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_MOVIMENTO_ESTOQUE)
					end if
				
			'	ATUALIZA O VALOR TOTAL DA FAMÍLIA DE PEDIDOS
			'	OBTÉM OS VALORES A PAGAR, JÁ PAGO E O STATUS DE PAGAMENTO (PARA TODA A FAMÍLIA DE PEDIDOS)
				if Not calcula_pagamentos(pedido_selecionado, vl_TotalFamiliaPrecoVenda, vl_TotalFamiliaPrecoNF, vl_TotalFamiliaPago, vl_TotalFamiliaDevolucaoPrecoVenda, vl_TotalFamiliaDevolucaoPrecoNF, st_pagto, msg_erro) then 
				'	~~~~~~~~~~~~~~~~
					cn.RollbackTrans
				'	~~~~~~~~~~~~~~~~
					Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
					end if
				
				vl_total_RA = vl_TotalFamiliaPrecoNF - vl_TotalFamiliaPrecoVenda
				
				if Not calcula_total_RA_liquido_BD(pedido_selecionado, vl_total_RA_liquido, msg_erro) then
				'	~~~~~~~~~~~~~~~~
					cn.RollbackTrans
				'	~~~~~~~~~~~~~~~~
					Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
					end if
				
				id_pedido_base = retorna_num_pedido_base(pedido_selecionado)
				s = "SELECT * FROM t_PEDIDO WHERE (pedido='" & id_pedido_base & "')"
				if rs.State <> 0 then rs.Close
				rs.Open s, cn
				if Not rs.Eof then
					rs("vl_total_familia") = vl_TotalFamiliaPrecoVenda
					rs("vl_total_NF") = vl_TotalFamiliaPrecoNF
					rs("vl_total_RA") = vl_total_RA
					rs("vl_total_RA_liquido") = vl_total_RA_liquido
					rs("qtde_parcelas_desagio_RA") = 0
					if vl_total_RA <> 0 then
						rs("st_tem_desagio_RA") = 1
					else
						rs("st_tem_desagio_RA") = 0
						end if
					rs.Update
					if Err <> 0 then
					'	~~~~~~~~~~~~~~~~
						cn.RollbackTrans
					'	~~~~~~~~~~~~~~~~
						Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
						end if
					rs.Close 
					end if
				end if
			end if
	'	~~~~~~~~~~~~~~
		cn.CommitTrans
	'	~~~~~~~~~~~~~~
		
		if rs.State <> 0 then rs.Close
		set rs = nothing
		
		if Err=0 then 
			if alerta = "" then Response.Redirect("pedido.asp?pedido_selecionado=" & pedido_selecionado & "&url_back=X" & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")))
		else
			alerta=Cstr(Err) & ": " & Err.Description
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