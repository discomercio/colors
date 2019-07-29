<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =================================================================
'	  RelSolicitacaoColetasGravaDados.asp
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

	Const COD_TIPO_RELATORIO_SOLICITACAO_COLETA = "SOLICITACAO_COLETA"
	Const COD_TIPO_RELATORIO_PRONTO_PARA_ROMANEIO = "PRONTO_PARA_ROMANEIO"

	dim blnExecutaUpdate
	dim s, usuario, msg_erro, s_log, s_log_aux, s_log_transp
	Dim campos_a_omitir
	Dim vLog1()
	Dim vLog2()
	campos_a_omitir = "|timestamp|a_entregar_data|a_entregar_hora|a_entregar_usuario|transportadora_data|transportadora_usuario|danfe_impressa_data|danfe_impressa_data_hora|danfe_impressa_usuario|"

	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

	dim alerta
	alerta=""

'	OBT�M FILTROS
	dim rb_loja, c_loja, c_loja_de, c_loja_ate, c_filtro_transportadora, c_filtro_dt_entrega, c_nfe_emitente
	dim rb_tipo_relatorio

	rb_loja = Ucase(Trim(Request.Form("rb_loja")))
	c_loja = Trim(Request.Form("c_loja"))
	c_loja_de = Trim(Request.Form("c_loja_de"))
	c_loja_ate = Trim(Request.Form("c_loja_ate"))
	rb_tipo_relatorio = Trim(Request.Form("rb_tipo_relatorio"))
	c_filtro_transportadora = Trim(Request.Form("c_filtro_transportadora"))
	c_filtro_dt_entrega = Trim(Request.Form("c_filtro_dt_entrega"))
	c_nfe_emitente = Trim(Request.Form("c_nfe_emitente"))

'	OBT�M DADOS DO FORMUL�RIO
	dim i, n, s_pedido, vAux, strFlagSolicEmissaoNFeJaCadastrada
	dim intNsu

'	TIPO DE RELAT�RIO: SOLICITA��O DE COLETA
'	========================================
'	CHECK BOX P/ INDICAR A GRAVA��O DA DATA DE COLETA + TRANSPORTADORA
'	CHECK BOX P/ SOLICITAR A EMISS�O DA NFe
	dim v_pedido_solicitacao, qtde_pedido_solicitacao
	redim v_pedido_solicitacao(0)
	v_pedido_solicitacao(Ubound(v_pedido_solicitacao))=""
	qtde_pedido_solicitacao=0
	
	if rb_tipo_relatorio = COD_TIPO_RELATORIO_SOLICITACAO_COLETA then
		n = Request.Form("ckb_emitir_nfe").Count
		for i = 1 to n
			s_pedido = Trim(Request.Form("ckb_emitir_nfe")(i))
			if s_pedido <> "" then
				vAux=Split(s_pedido,"|")
				s_pedido = Trim(vAux(LBound(vAux)))
				strFlagSolicEmissaoNFeJaCadastrada = UCase(Trim(vAux(UBound(vAux))))
				if strFlagSolicEmissaoNFeJaCadastrada = "N" then
					if Trim(v_pedido_solicitacao(Ubound(v_pedido_solicitacao)))<>"" then
						redim preserve v_pedido_solicitacao(Ubound(v_pedido_solicitacao)+1)
						end if
					v_pedido_solicitacao(Ubound(v_pedido_solicitacao)) = s_pedido
					qtde_pedido_solicitacao = qtde_pedido_solicitacao + 1
					end if
				end if
			next
		end if

	dim c_transportadora, c_dt_entrega
	dim v_pedido_transp, qtde_pedido_transp
	redim v_pedido_transp(0)
	v_pedido_transp(UBound(v_pedido_transp))=""
	qtde_pedido_transp=0
	
	c_transportadora = Trim(Request.Form("c_transportadora"))
	c_dt_entrega = Trim(Request.Form("c_dt_entrega"))
	
	if rb_tipo_relatorio = COD_TIPO_RELATORIO_SOLICITACAO_COLETA then
		n = Request.Form("ckb_gravar_transp_e_dt_entrega").Count
		for i = 1 to n
			s_pedido = Trim(Request.Form("ckb_gravar_transp_e_dt_entrega")(i))
			if s_pedido <> "" then
				if Trim(v_pedido_transp(Ubound(v_pedido_transp)))<>"" then
					redim preserve v_pedido_transp(UBound(v_pedido_transp)+1)
					end if
				v_pedido_transp(Ubound(v_pedido_transp)) = s_pedido
				qtde_pedido_transp = qtde_pedido_transp + 1
				end if
			next
		end if
	
	if alerta = "" then
		if rb_tipo_relatorio = COD_TIPO_RELATORIO_SOLICITACAO_COLETA then
			if (qtde_pedido_solicitacao = 0) And (qtde_pedido_transp = 0) then
				alerta = "N�o foi especificado nenhum pedido para solicitar a emiss�o da NFe e nenhum pedido para gravar transportadora/data de coleta."
				end if
			end if
		end if

	if alerta = "" then
		if qtde_pedido_transp > 0 then
			if c_dt_entrega = "" then
				alerta = "Data de coleta n�o foi informada."
			elseif Not isDate(c_dt_entrega) then
				alerta = "Data de coleta � inv�lida."
			elseif StrToDate(c_dt_entrega) < Date then
				alerta = "Data de coleta n�o pode ser uma data passada."
				end if
			end if
		end if

'	TIPO DE RELAT�RIO: PEDIDOS PRONTOS P/ ROMANEIO
'	==============================================
'	CHECK BOX P/ ASSINALAR SE A DANFE J� FOI IMPRESSA, LEMBRANDO QUE OS PEDIDOS MARCADOS COM "DANFE J� IMPRESSA" S�O RETIRADOS DA LISTAGEM
	dim v_pedido_danfe_impressa, qtde_pedido_danfe_impressa
'	CHECK BOX P/ ASSINALAR SE A DANFE SER� EXIBIDA PARA IMPRESS�O NO PROGRAMA PRNDANFE, LEMBRANDO QUE OS PEDIDOS MARCADOS COM "DANFE J� IMPRESSA" S�O RETIRADOS DA LISTAGEM
	dim v_pedido_danfe_a_imprimir, qtde_pedido_danfe_a_imprimir
	redim v_pedido_danfe_impressa(0)
	v_pedido_danfe_impressa(UBound(v_pedido_danfe_impressa))=""
	qtde_pedido_danfe_impressa=0
	redim v_pedido_danfe_a_imprimir(0)
	v_pedido_danfe_a_imprimir(UBound(v_pedido_danfe_a_imprimir))=""
	qtde_pedido_danfe_a_imprimir=0
	
	if rb_tipo_relatorio = COD_TIPO_RELATORIO_PRONTO_PARA_ROMANEIO then
		n = Request.Form("ckb_danfe_impressa").Count
		for i = 1 to n
			s_pedido = Trim(Request.Form("ckb_danfe_impressa")(i))
			if s_pedido <> "" then
				if Trim(v_pedido_danfe_impressa(Ubound(v_pedido_danfe_impressa)))<>"" then
					redim preserve v_pedido_danfe_impressa(UBound(v_pedido_danfe_impressa)+1)
					end if
				v_pedido_danfe_impressa(Ubound(v_pedido_danfe_impressa)) = s_pedido
				qtde_pedido_danfe_impressa = qtde_pedido_danfe_impressa + 1
				end if
			next

		n = Request.Form("ckb_danfe_a_imprimir").Count
		for i = 1 to n
			s_pedido = Trim(Request.Form("ckb_danfe_a_imprimir")(i))
			if s_pedido <> "" then
				if Trim(v_pedido_danfe_a_imprimir(Ubound(v_pedido_danfe_a_imprimir)))<>"" then
					redim preserve v_pedido_danfe_a_imprimir(UBound(v_pedido_danfe_a_imprimir)+1)
					end if
				v_pedido_danfe_a_imprimir(Ubound(v_pedido_danfe_a_imprimir)) = s_pedido
				qtde_pedido_danfe_a_imprimir = qtde_pedido_danfe_a_imprimir + 1
				end if
			next
		end if

	if alerta = "" then
		if rb_tipo_relatorio = COD_TIPO_RELATORIO_PRONTO_PARA_ROMANEIO then
			if (qtde_pedido_danfe_impressa = 0) and (qtde_pedido_danfe_a_imprimir = 0) then
				alerta = "N�o foi selecionado nenhum pedido para gravar a sinaliza��o de que a DANFE foi impressa."
				alerta = alerta & chr(13) & "Tamb�m n�o foi marcada nenhuma DANFE para impress�o."
				end if
			end if
		end if
				

'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	
	s_log = ""
	s_log_transp = ""
	
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

		if rb_tipo_relatorio = COD_TIPO_RELATORIO_SOLICITACAO_COLETA then
		'	GRAVA SOLICITA��O DE EMISS�O DE NFe
		'	===================================
			for i=Lbound(v_pedido_solicitacao) to Ubound(v_pedido_solicitacao)
				if v_pedido_solicitacao(i) <> "" then
					s = "SELECT * FROM t_PEDIDO WHERE (pedido = '" & v_pedido_solicitacao(i) & "')"
					if rs.State <> 0 then rs.Close
					rs.Open s, cn
					if rs.Eof then
						alerta=texto_add_br(alerta)
						alerta=alerta & "Pedido " & v_pedido_solicitacao(i) & " n�o foi encontrado."
					else
						blnExecutaUpdate = False
					'	CORRIGE STATUS DE PEDIDOS CADASTRADOS ANTES DA CRIA��O DO CAMPO
						if Trim("" & rs("romaneio_status")) = Cstr(COD_ROMANEIO_STATUS__NAO_DEFINIDO) then 
							rs("romaneio_status")=CInt(COD_ROMANEIO_STATUS__INICIAL)
							blnExecutaUpdate = True
							end if
							
					'	CORRIGE STATUS DE PEDIDOS CADASTRADOS ANTES DA CRIA��O DO CAMPO
						if Trim("" & rs("danfe_impressa_status")) = Cstr(COD_DANFE_IMPRESSA_STATUS__NAO_DEFINIDO) then 
							rs("danfe_impressa_status")=CInt(COD_DANFE_IMPRESSA_STATUS__INICIAL)
							blnExecutaUpdate = True
							end if
							
						if blnExecutaUpdate then rs.Update
						
						s = "SELECT" & _
								" Count(*) AS qtde" & _
							" FROM t_PEDIDO_NFe_EMISSAO_SOLICITADA" & _
							" WHERE" & _
								" (pedido = '" & v_pedido_solicitacao(i) & "')" & _
								" AND (" & _
									"(nfe_emitida_status = " & COD_NFE_EMISSAO_SOLICITADA__PENDENTE & ")" & _
									" OR (nfe_emitida_status = " & COD_NFE_EMISSAO_SOLICITADA__ATENDIDA & ")" & _
									")"
						if rs.State <> 0 then rs.Close
						rs.Open s, cn
						if CLng(rs("qtde")) > 0 then alerta = "Pedido " & v_pedido_solicitacao(i) & " j� teve a emiss�o de NFe solicitada."
						
						if alerta = "" then
						'	INFORMA��ES PARA O LOG
							if s_log <> "" then s_log = s_log & ", "
							s_log = s_log & v_pedido_solicitacao(i)
							
						'	GERA O NSU PARA O NOVO REGISTRO
							if Not fin_gera_nsu(T_PEDIDO_NFE_EMISSAO_SOLICITADA, intNsu, msg_erro) then 
								alerta = "FALHA AO GERAR NSU PARA O NOVO REGISTRO (" & msg_erro & ")"
							else
								if intNsu <= 0 then
									alerta = "NSU GERADO � INV�LIDO (" & intNsu & ")"
									end if
								end if
							
							s = "SELECT * FROM t_PEDIDO_NFe_EMISSAO_SOLICITADA WHERE (id = -1)"
							if rs.State <> 0 then rs.Close
							rs.Open s, cn
							rs.AddNew
							rs("id") = intNsu
							rs("pedido") = v_pedido_solicitacao(i)
							rs("usuario") = usuario
							rs.Update
							if Err <> 0 then 
								alerta=texto_add_br(alerta)
								alerta=alerta & Cstr(Err) & ": " & Err.Description
								end if
							end if
						end if
					end if
					
			'	SE HOUVE ERRO, CANCELA O LA�O
				if alerta <> "" then exit for
				next

			if alerta = "" then
				if s_log <> "" then
					s_log = "Emiss�o de NFe solicitada para o(s) pedido(s): " & s_log
					grava_log usuario, "", "", "", OP_LOG_PEDIDO_NFE_EMISSAO_SOLICITADA, s_log
					end if
				end if
			
			
		'	GRAVA TRANSPORTADORA + DATA DE COLETA
		'	=====================================
			s_log = ""
			s_log_aux = ""
			for i=Lbound(v_pedido_transp) to Ubound(v_pedido_transp)
				if v_pedido_transp(i) <> "" then
					s = "SELECT * FROM t_PEDIDO WHERE (pedido = '" & v_pedido_transp(i) & "')"
					if rs.State <> 0 then rs.Close
					rs.Open s, cn
					if rs.Eof then
						alerta=texto_add_br(alerta)
						alerta=alerta & "Pedido " & v_pedido_transp(i) & " n�o foi encontrado."
					else
					'	INFORMA��ES PARA O LOG
						if s_log <> "" then s_log = s_log & ", "
						s_log = s_log & v_pedido_transp(i)
						
						log_via_vetor_carrega_do_recordset rs, vLog1, campos_a_omitir
						
					'	CORRIGE STATUS DE PEDIDOS CADASTRADOS ANTES DA CRIA��O DO CAMPO
						if Trim("" & rs("romaneio_status")) = Cstr(COD_ROMANEIO_STATUS__NAO_DEFINIDO) then 
							rs("romaneio_status")=CInt(COD_ROMANEIO_STATUS__INICIAL)
							end if
							
					'	CORRIGE STATUS DE PEDIDOS CADASTRADOS ANTES DA CRIA��O DO CAMPO
						if Trim("" & rs("danfe_impressa_status")) = Cstr(COD_DANFE_IMPRESSA_STATUS__NAO_DEFINIDO) then 
							rs("danfe_impressa_status")=CInt(COD_DANFE_IMPRESSA_STATUS__INICIAL)
							end if
						
						if c_transportadora <> "" then
							if Ucase(Trim("" & rs("transportadora_id"))) <> Ucase(c_transportadora) then
								if s_log_transp <> "" then s_log_transp = s_log_transp & "; "
								s_log_transp = s_log_transp & v_pedido_transp(i) & ": '" & Ucase(Trim("" & rs("transportadora_id"))) & "' -> '" & Ucase(c_transportadora) & "'"
							'	LIMPA DADOS DA SELE��O AUTOM�TICA DE TRANSPORTADORA BASEADO NO CEP
							'	MANT�M OS DADOS ANTERIORES (SE HOUVER) P/ FINS DE HIST�RICO/LOG DOS SEGUINTES CAMPOS:
							'	transportadora_selecao_auto_cep, transportadora_selecao_auto_tipo_endereco e transportadora_selecao_auto_transportadora
								rs("transportadora_selecao_auto_status") = TRANSPORTADORA_SELECAO_AUTO_STATUS_FLAG_N
								rs("transportadora_selecao_auto_data_hora") = Now
							
							'	GRAVA DADOS DA TRANSPORTADORA
								rs("transportadora_id")=c_transportadora
								rs("transportadora_data")=Now
								rs("transportadora_usuario")=usuario
								rs("transportadora_num_coleta")=""
								rs("transportadora_contato")=""
								end if
							end if
						
						if Trim("" & rs("a_entregar_data_marcada")) <> Trim("" & StrToDate(c_dt_entrega)) then
							rs("a_entregar_status")=1
							rs("a_entregar_data_marcada")=StrToDate(c_dt_entrega)
							rs("a_entregar_data")=Date
							rs("a_entregar_hora")=retorna_so_digitos(formata_hora(Now))
							rs("a_entregar_usuario")=usuario
							end if

						rs.Update
						if Err = 0 then
							log_via_vetor_carrega_do_recordset rs, vLog2, campos_a_omitir
							s = log_via_vetor_monta_alteracao(vLog1, vLog2)
							if s <> "" then
								if s_log_aux <> "" then s_log_aux = s_log_aux & " " & chr(13)
								s_log_aux = s_log_aux & "Pedido " & v_pedido_transp(i) & ": " & s
								end if
						else
							alerta=texto_add_br(alerta)
							alerta=alerta & Cstr(Err) & ": " & Err.Description
							end if
						end if
					end if
				
			'	SE HOUVE ERRO, CANCELA O LA�O
				if alerta <> "" then exit for
				next
			
			if alerta = "" then
				if s_log <> "" then
					if c_transportadora <> "" then
						s_log = "Solicita��o de Coletas: anota��o da transportadora '" & c_transportadora & "' e data de coleta '" & c_dt_entrega & "' para o(s) pedido(s): " & s_log & "; Log de altera��o da transportadora: " & s_log_transp
					else
						s_log = "Solicita��o de Coletas: anota��o da data de coleta '" & c_dt_entrega & "' para o(s) pedido(s): " & s_log
						end if
					if s_log_aux <> "" then s_log = s_log & " " & chr(13) & s_log_aux
					grava_log usuario, "", "", "", OP_LOG_PEDIDO_ALTERACAO, s_log
					end if
				end if
			
			end if	'if rb_tipo_relatorio = COD_TIPO_RELATORIO_SOLICITACAO_COLETA


		if rb_tipo_relatorio = COD_TIPO_RELATORIO_PRONTO_PARA_ROMANEIO then
		'	GRAVA INDICA��O SE A DANFE J� FOI IMPRESSA
		'	==========================================
			s_log = ""
			s_log_aux = ""
			for i=Lbound(v_pedido_danfe_impressa) to Ubound(v_pedido_danfe_impressa)
				if v_pedido_danfe_impressa(i) <> "" then
					s = "SELECT * FROM t_PEDIDO WHERE (pedido = '" & v_pedido_danfe_impressa(i) & "')"
					if rs.State <> 0 then rs.Close
					rs.Open s, cn
					if rs.Eof then
						alerta=texto_add_br(alerta)
						alerta=alerta & "Pedido " & v_pedido_danfe_impressa(i) & " n�o foi encontrado."
					else
					'	INFORMA��ES PARA O LOG
						if s_log <> "" then s_log = s_log & ", "
						s_log = s_log & v_pedido_danfe_impressa(i)
						
						log_via_vetor_carrega_do_recordset rs, vLog1, campos_a_omitir

						rs("danfe_impressa_status")=CInt(COD_DANFE_IMPRESSA_STATUS__OK)
						rs("danfe_impressa_data")=Date
						rs("danfe_impressa_data_hora")=Now
						rs("danfe_impressa_usuario")=usuario

						rs.Update
						if Err = 0 then
							log_via_vetor_carrega_do_recordset rs, vLog2, campos_a_omitir
							s = log_via_vetor_monta_alteracao(vLog1, vLog2)
							if s <> "" then
								if s_log_aux <> "" then s_log_aux = s_log_aux & " " & chr(13)
								s_log_aux = s_log_aux & "Pedido " & v_pedido_danfe_impressa(i) & ": " & s
								end if
						else
							alerta=texto_add_br(alerta)
							alerta=alerta & Cstr(Err) & ": " & Err.Description
							end if
						end if
					end if
					
			'	SE HOUVE ERRO, CANCELA O LA�O
				if alerta <> "" then exit for
				next
			
			if alerta = "" then
				if s_log <> "" then
					s_log = "Anota��o de DANFE impressa para o(s) pedido(s): " & s_log
					if s_log_aux <> "" then s_log = s_log & " " & chr(13) & s_log_aux
					grava_log usuario, "", "", "", OP_LOG_PEDIDO_ALTERACAO, s_log
					end if
				end if

		'	GRAVA MARCA��O SE A DANFE SER� IMPRESSA
		'	=======================================
			s_log = ""
			s_log_aux = ""
			for i=Lbound(v_pedido_danfe_a_imprimir) to Ubound(v_pedido_danfe_a_imprimir)
				if v_pedido_danfe_a_imprimir(i) <> "" then
					s = "SELECT * FROM t_PEDIDO WHERE (pedido = '" & v_pedido_danfe_a_imprimir(i) & "')"
					if rs.State <> 0 then rs.Close
					rs.Open s, cn
					if rs.Eof then
						alerta=texto_add_br(alerta)
						alerta=alerta & "Pedido " & v_pedido_danfe_a_imprimir(i) & " n�o foi encontrado."
					else
					'	INFORMA��ES PARA O LOG
						if s_log <> "" then s_log = s_log & ", "
						s_log = s_log & v_pedido_danfe_a_imprimir(i)
						
						log_via_vetor_carrega_do_recordset rs, vLog1, campos_a_omitir

						rs("danfe_a_imprimir_status")=CInt(COD_DANFE_A_IMPRIMIR_STATUS__MARCADA)
						rs("danfe_a_imprimir_data_hora")=Now
						rs("danfe_a_imprimir_usuario")=usuario

						rs.Update
						if Err = 0 then
							log_via_vetor_carrega_do_recordset rs, vLog2, campos_a_omitir
							s = log_via_vetor_monta_alteracao(vLog1, vLog2)
							if s <> "" then
								if s_log_aux <> "" then s_log_aux = s_log_aux & " " & chr(13)
								s_log_aux = s_log_aux & "Pedido " & v_pedido_danfe_a_imprimir(i) & ": " & s
								end if
						else
							alerta=texto_add_br(alerta)
							alerta=alerta & Cstr(Err) & ": " & Err.Description
							end if
						end if
					end if
					
			'	SE HOUVE ERRO, CANCELA O LA�O
				if alerta <> "" then exit for
				next
			if alerta = "" then
				if s_log <> "" then
					s_log = "Marca��o de DANFE a imprimir para o(s) pedido(s): " & s_log
					if s_log_aux <> "" then s_log = s_log & " " & chr(13) & s_log_aux
					grava_log usuario, "", "", "", OP_LOG_PEDIDO_ALTERACAO, s_log
					end if
				end if

			end if	'if rb_tipo_relatorio = COD_TIPO_RELATORIO_PRONTO_PARA_ROMANEIO


	'	FINALIZA TRANSA��O
	'	==================
		if alerta = "" then
		'	~~~~~~~~~~~~~~
			cn.CommitTrans
		'	~~~~~~~~~~~~~~
			if Err<>0 then 
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

<script language="JavaScript" type="text/javascript">
function fRetornar(f) {
	f.action = "RelSolicitacaoColetasExec.asp?url_back=X";
	f.submit();
}
</script>




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
<!-- **********  P�GINA PARA EXIBIR MENSAGENS DE ERRO  ********** -->
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
<% else %>

<!-- ***************************************************** -->
<!-- **********  P�GINA PARA EXIBIR RESULTADO   ********** -->
<!-- ***************************************************** -->
<body onload="window.status='Conclu�do';bVOLTAR.focus();" link=#000000 alink=#000000 vlink=#000000>

<center>

<form id="f" name="f" method="post">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<!-- FILTROS -->
<input type="hidden" name="rb_loja" id="rb_loja" value="<%=rb_loja%>" />
<input type="hidden" name="c_loja" id="c_loja" value="<%=c_loja%>" />
<input type="hidden" name="c_loja_de" id="c_loja_de" value="<%=c_loja_de%>" />
<input type="hidden" name="c_loja_ate" id="c_loja_ate" value="<%=c_loja_ate%>" />
<input type="hidden" name="rb_tipo_relatorio" id="rb_tipo_relatorio" value="<%=rb_tipo_relatorio%>" />
<input type="hidden" name="c_filtro_transportadora" id="c_filtro_transportadora" value="<%=c_filtro_transportadora%>" />
<input type="hidden" name="c_filtro_dt_entrega" id="c_filtro_dt_entrega" value="<%=c_filtro_dt_entrega%>" />
<input type="hidden" name="c_nfe_emitente" id="c_nfe_emitente" value="<%=c_nfe_emitente%>" />


<!--  I D E N T I F I C A � � O   D A   T E L A  -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">

<tr>
	<td align="right" valign="bottom"><p class="PEDIDO">Solicita��o de Coletas<span class="C">&nbsp;</span></p></td>
</tr>
</table>
<br>
<br>

<% if rb_tipo_relatorio = COD_TIPO_RELATORIO_SOLICITACAO_COLETA then %>

<% if qtde_pedido_transp > 0 then %>
<!-- ************   MENSAGEM  ************ -->
<% 
	s = ""
	for i=Lbound(v_pedido_transp) to Ubound(v_pedido_transp)
		if v_pedido_transp(i) <> "" then
			if s <> "" then s = s & ", "
			s = s & v_pedido_transp(i)
			end if
		next
	
	if s = "" then s = "nenhum pedido"
%>
<% if c_transportadora <> "" then %>
<div class='MtAviso' style="width:649px;font-weight:bold;border:1pt solid black;" align="center"><p style='margin:5px 2px 5px 2px;'>Anota��o da transportadora '<%=c_transportadora%>' e data de coleta <%=c_dt_entrega%><br />Pedido(s): <%=s%></p></div>
<% else %>
<div class='MtAviso' style="width:649px;font-weight:bold;border:1pt solid black;" align="center"><p style='margin:5px 2px 5px 2px;'>Anota��o da data de coleta <%=c_dt_entrega%><br />Pedido(s): <%=s%></p></div>
<% end if %>
<br>
<% end if %>

<%if qtde_pedido_solicitacao > 0 then %>
<!-- ************   MENSAGEM  ************ -->
<% 
	s = ""
	for i=Lbound(v_pedido_solicitacao) to Ubound(v_pedido_solicitacao)
		if v_pedido_solicitacao(i) <> "" then
			if s <> "" then s = s & ", "
			s = s & v_pedido_solicitacao(i)
			end if
		next
	
	if s = "" then s = "nenhum pedido"
%>
<div class='MtAviso' style="width:649px;font-weight:bold;border:1pt solid black;" align="center"><p style='margin:5px 2px 5px 2px;'>Solicita��o de Emiss�o da NFe<br />Pedido(s): <%=s%></p></div>
<br>
<% end if %>

<% elseif rb_tipo_relatorio = COD_TIPO_RELATORIO_PRONTO_PARA_ROMANEIO then %>
<% 
	s = ""
	for i=Lbound(v_pedido_danfe_impressa) to Ubound(v_pedido_danfe_impressa)
		if v_pedido_danfe_impressa(i) <> "" then
			if s <> "" then s = s & ", "
			s = s & v_pedido_danfe_impressa(i)
			end if
		next

	for i=Lbound(v_pedido_danfe_a_imprimir) to Ubound(v_pedido_danfe_a_imprimir)
		if v_pedido_danfe_a_imprimir(i) <> "" then
			if InStr(s, v_pedido_danfe_a_imprimir(i)) <= 0 then
				if s <> "" then s = s & ", "
				s = s & v_pedido_danfe_a_imprimir(i)
				end if
			end if
		next
	
	if s = "" then s = "nenhum pedido"
%>
<div class='MtAviso' style="width:649px;font-weight:bold;border:1pt solid black;" align="center"><p style='margin:5px 2px 5px 2px;'>Anota��o de DANFE impressa / Marca��o para impress�o<br />Pedido(s): <%=s%></p></div>
<br>
<% end if %>

<!-- ************   SEPARADOR   ************ -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>

<!-- ************   LINKS: P�GINA INICIAL / ENCERRA SESS�O   ************ -->
<table width="649" cellPadding="0" CellSpacing="0">
<tr><td align="right"><span class="Rc">
	<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para p�gina inicial" class="LPagInicial">p�gina inicial</a>&nbsp;&nbsp;&nbsp;
	<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sess�o do usu�rio" class="LSessaoEncerra">encerra</a>
	</span>
</td></tr>
</table>

<!-- ************   BOT�ES   ************ -->
<table class="notPrint" width="649" cellSpacing="0">
<tr>
	<td align="center"><a name="bVOLTAR" id="bVOLTAR" href="javascript:fRetornar(f)" title="Retornar para a p�gina anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
</tr>
</table>
</form>

</center>
</body>

<% end if %>

</html>


<%
'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>