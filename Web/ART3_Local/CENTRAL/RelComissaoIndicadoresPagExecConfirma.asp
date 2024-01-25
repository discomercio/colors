<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =================================================================================
'	  R E L C O M I S S A O I N D I C A D O R E S P A G E X E C C O N F I R M A . A S P
'     =================================================================================
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

	dim s, msg_erro
	dim usuario, loja
	usuario = Trim(Session("usuario_atual"))
	loja = ""
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO)

	dim alerta
	alerta=""
	
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

	if Not operacao_permitida(OP_CEN_REL_PEDIDOS_INDICADORES_PAGAMENTO, s_lista_operacoes_permitidas) then
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs, rs2,rs3
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	if Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	if Not cria_recordset_otimista(rs2, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	if Not cria_recordset_otimista(rs3, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

	dim c_lista_info_indicador, c_lista_info_pedido
	dim vIndicador, vPedido
	c_lista_info_indicador = Trim(Request.Form("c_lista_info_indicador"))
	c_lista_info_pedido = Trim(Request.Form("c_lista_info_pedido"))
	
	if Not deserializa_cl_REL_PEDIDOS_INDICADORES_INFO_INDICADOR(c_lista_info_indicador, vIndicador, msg_erro) then
		alerta = texto_add_br(alerta)
		alerta = alerta & "Falha ao tentar decodificar os dados de indicadores"
		if msg_erro <> "" then alerta = alerta & " (" & msg_erro & ")"
		end if

	if Not deserializa_cl_REL_PEDIDOS_INDICADORES_INFO_PEDIDO(c_lista_info_pedido, vPedido , msg_erro) then
		alerta = texto_add_br(alerta)
		alerta = alerta & "Falha ao tentar decodificar os dados de pedidos"
		if msg_erro <> "" then alerta = alerta & " (" & msg_erro & ")"
		end if

	if alerta = "" then
		for i=LBound(vIndicador) to UBound(vIndicador)
			with vIndicador(i)
				if .IdIndicador <> 0 then
					s = "IdIndicador=" & .IdIndicador & _
						", IdVendedor=" & .IdVendedor & _
						", meio_pagto=" & .meio_pagto & _
						", indicador_com_desconto=" & .indicador_com_desconto & _
						", indicador_negativo=" & .indicador_negativo & _
						", vl_total_comissao=" & formata_moeda(.vl_total_comissao) & _
						", vl_total_comissao_arredondado=" & formata_moeda(.vl_total_comissao_arredondado) & _
						", vl_total_RA=" & formata_moeda(.vl_total_RA) & _
						", vl_total_RA_arredondado=" & formata_moeda(.vl_total_RA_arredondado) & _
						", vl_total_pagto=" & formata_moeda(.vl_total_pagto) & _
						", vl_total_desc_planilha=" & formata_moeda(.vl_total_desc_planilha) & _
						", qtde_reg_descontos=" & Cstr(.qtde_reg_descontos)
					Response.Write s & "<br />"
					end if
				end with
			next
		
		Response.Write "<br /><br /><br />"
		for i=LBound(vPedido) to UBound(vPedido)
			with vPedido(i)
				if .pedido <> "" then
					s = "pedido=" & .pedido & _
						", IdIndicador=" & .IdIndicador & _
						", IdVendedor=" & .IdVendedor & _
						", operacao=" & .operacao & _
						", id_registro_operacao=" & .id_registro_operacao & _
						", vl_pedido=" & formata_moeda(.vl_pedido) & _
						", vl_comissao=" & formata_moeda(.vl_comissao) & _
						", vl_RA_bruto=" & formata_moeda(.vl_RA_bruto) & _
						", vl_RA_liquido=" & formata_moeda(.vl_RA_liquido)
					Response.Write s & "<br />"
					end if
				end with
			next
		end if

	Response.End

	dim sql_n1, sql_n2, sql_n3, sql_n4, s_pedido_indicador, indicador_a, vendedor_a
	dim ckb_bloco_indicador, v_pedido, i, mes_competencia, ano_competencia, c_vendedor, v_vendedor, j, pos_v_planilha_desconto, n, c_lista_completa_pedidos
	dim c_lista_vl_comissao, c_lista_vl_RA_bruto, c_lista_vl_RA_liquido, c_lista_vl_pedido, c_lista_vl_total_comissao, c_lista_vl_total_comissao_arredondado, c_lista_meio_pagto, c_lista_vl_total_RA, c_lista_vl_total_RA_arredondado
	dim v_lista_completa_pedidos, v_lista_vl_comissao, v_lista_vl_RA_bruto, v_lista_vl_RA_liquido, v_lista_vl_pedido, v_lista_vl_total_comissao, v_lista_vl_total_comissao_arredondado, v_lista_meio_pagto
	dim v_lista_vl_total_RA, v_lista_vl_total_RA_arredondado
	dim intNsuNovoComissaoN1, intNsuNovoComissaoN2, intNsuNovoComissaoN3, intNsuNovoComissaoN4, marcado, ind_completo
	dim lista_indicador_com_desconto, lista_indicador_negativo, total_desconto_planilha, v_total_desconto_planilha, c_lista_qtde_reg_descontos, v_lista_qtde_reg_descontos
	dim tem_desconto, c_lista_operacao, v_lista_operacao,vendedor,mensagem, aviso, c_lista_vl_total_pagto, v_lista_vl_total_pagto, rb_visao

	mes_competencia = Trim(Request.Form("mes"))
	ano_competencia = Trim(Request.Form("ano"))
	c_lista_completa_pedidos = Trim(Request.Form("c_lista_completa_pedidos"))
	c_lista_vl_comissao = Trim(Request.Form("c_lista_vl_comissao"))
	c_lista_vl_RA_bruto = Trim(Request.Form("c_lista_vl_RA_bruto"))
	c_lista_vl_RA_liquido = Trim(Request.Form("c_lista_vl_RA_liquido"))
	c_lista_vl_pedido = Trim(Request.Form("c_lista_vl_pedido"))
	c_lista_vl_total_comissao = Trim(Request.Form("c_lista_vl_total_comissao"))
	c_lista_vl_total_comissao_arredondado = Trim(Request.Form("c_lista_vl_total_comissao_arredondado"))
	c_lista_vl_total_RA = Trim(Request.Form("c_lista_vl_total_RA"))
	c_lista_vl_total_RA_arredondado = Trim(Request.Form("c_lista_vl_total_RA_arredondado"))
	c_lista_meio_pagto =  Trim(Request.Form("c_lista_meio_pagto"))
	'O valor do check box 'ckb_bloco_indicador' � a rela��o de pedidos
	ckb_bloco_indicador = Trim(Request.Form("ckb_bloco_indicador"))
	c_vendedor = Trim(Request.Form("c_vendedor"))
	v_pedido = split(ckb_bloco_indicador, ", ")
	v_lista_completa_pedidos = split(c_lista_completa_pedidos, ";")
	v_lista_vl_comissao = split(c_lista_vl_comissao, ";")
	v_lista_vl_RA_bruto = split(c_lista_vl_RA_bruto, ";")
	v_lista_vl_RA_liquido = split(c_lista_vl_RA_liquido, ";")
	v_lista_vl_pedido = split(c_lista_vl_pedido, ";")
	v_lista_vl_total_comissao = split(c_lista_vl_total_comissao, ";")
	v_lista_vl_total_comissao_arredondado = split(c_lista_vl_total_comissao_arredondado, ";")
	v_lista_vl_total_RA = split(c_lista_vl_total_RA, ";")
	v_lista_vl_total_RA_arredondado = split(c_lista_vl_total_RA_arredondado, ";")
	v_lista_meio_pagto = split(c_lista_meio_pagto, ";")
	lista_indicador_com_desconto = Trim(Request.Form("c_lista_indicador_com_desconto"))
	lista_indicador_negativo = Trim(Request.Form("c_lista_indicador_negativo"))
	total_desconto_planilha = Trim(Request.Form("c_lista_vl_total_desc_planilha"))
	v_total_desconto_planilha = split(total_desconto_planilha, ";")
	c_lista_qtde_reg_descontos = Trim(Request.Form("c_lista_qtde_reg_descontos"))
	v_lista_qtde_reg_descontos = split(c_lista_qtde_reg_descontos, ";")
	c_lista_operacao = Trim(Request.Form("c_lista_operacao"))
	v_lista_operacao = split(c_lista_operacao, ", ")
	c_lista_vl_total_pagto = Request.Form("c_lista_vl_total_pagto")
	v_lista_vl_total_pagto = split(c_lista_vl_total_pagto, ";")
	rb_visao = Trim(Request.Form("rb_visao"))
 
	pos_v_planilha_desconto = 0
	n = 0
	vendedor_a = "XXXXXXXXX"
	indicador_a = "XXXXXXXXX"

	if alerta = "" then
		cn.BeginTrans

		sql_n1 = "SELECT * FROM t_COMISSAO_INDICADOR_N1 WHERE (id = -1)"
		if Not fin_gera_nsu(T_COMISSAO_INDICADOR_N1, intNsuNovoComissaoN1, msg_erro) then 
				alerta = "FALHA AO GERAR NSU PARA O NOVO REGISTRO (" & msg_erro & ")"
			else
				if intNsuNovoComissaoN1 <= 0 then
					alerta = "NSU GERADO � INV�LIDO (" & intNsuNovoComissaoN1 & ")"
					end if
				end if
		rs.Open sql_n1, cn
		rs.AddNew
		rs("id") = intNsuNovoComissaoN1
		rs("usuario") = usuario
		rs.Update
		if rs.State <> 0 then rs.Close

		sql_n2 = "SELECT * FROM t_COMISSAO_INDICADOR_N2 WHERE (id = -1)"
		sql_n3 = "SELECT * FROM t_COMISSAO_INDICADOR_N3 WHERE (id = -1)"
		sql_n4 = "SELECT * FROM t_COMISSAO_INDICADOR_N4 WHERE (id = -1)"
		v_vendedor = split(c_vendedor, ", ")
	
		for i=LBound(v_lista_completa_pedidos) to UBound(v_lista_completa_pedidos)
			s_pedido_indicador = "SELECT t_PEDIDO__BASE.vendedor vendedor, " & _
										"t_PEDIDO__BASE.indicador indicador, " & _
										"t_PEDIDO.perc_RT perc_RT, " & _
										"t_ORCAMENTISTA_E_INDICADOR.banco banco, " & _
										"t_ORCAMENTISTA_E_INDICADOR.agencia agencia, " & _
										"t_ORCAMENTISTA_E_INDICADOR.conta conta, " & _
										"t_ORCAMENTISTA_E_INDICADOR.favorecido favorecido, " & _
										"t_ORCAMENTISTA_E_INDICADOR.favorecido_cnpj_cpf favorecido_cnpj_cpf, " & _
										"t_ORCAMENTISTA_E_INDICADOR.agencia_dv agencia_dv, " & _
										"t_ORCAMENTISTA_E_INDICADOR.conta_operacao conta_operacao, " & _
										"t_ORCAMENTISTA_E_INDICADOR.conta_dv conta_dv, " & _
										"t_ORCAMENTISTA_E_INDICADOR.tipo_conta tipo_conta " & _
										"FROM t_PEDIDO INNER JOIN t_PEDIDO AS t_PEDIDO__BASE ON (t_PEDIDO.pedido_base=t_PEDIDO__BASE.pedido)" & _
										" INNER JOIN t_ORCAMENTISTA_E_INDICADOR ON (t_PEDIDO__BASE.indicador=t_ORCAMENTISTA_E_INDICADOR.apelido) " & _
										"WHERE (t_PEDIDO.pedido = '" & v_lista_completa_pedidos(i) & "')"

			set rs2 = cn.Execute(s_pedido_indicador)
			tem_desconto=false
		
			if Not rs2.Eof then
				if (vendedor_a <> Trim("" & rs2("vendedor"))) then
					if Not fin_gera_nsu(T_COMISSAO_INDICADOR_N2, intNsuNovoComissaoN2, msg_erro) then 
							alerta = "FALHA AO GERAR NSU PARA O NOVO REGISTRO (" & msg_erro & ")"
						else
							if intNsuNovoComissaoN2 <= 0 then
								alerta = "NSU GERADO � INV�LIDO (" & intNsuNovoComissaoN2 & ")"
								end if
							end if
					rs.Open sql_n2, cn
					rs.AddNew
					rs("id") = intNsuNovoComissaoN2
					rs("id_comissao_indicador_n1") = intNsuNovoComissaoN1
					rs("competencia_ano") = ano_competencia
					rs("competencia_mes") = mes_competencia
					rs("vendedor") = Trim("" & rs2("vendedor"))
					rs("proc_automatico_status") = 0
					rs("proc_automatico_qtde_tentativas") = 0
					On Error Resume Next
					rs.Update

					if Err <> 0 then 
							cn.RollbackTrans
							if instr(Err.Description, "insert duplicate key") > 0 then
							Response.Redirect("aviso.asp?id=" & ERR_INDICADORES_VENDEDOR_INFORMADO_JA_PROCESSADO)
							end if
					end if
					On Error GoTo 0
					if rs.State <> 0 then rs.Close
				 end if

					if (indicador_a <> Trim("" & rs2("indicador"))) then
					
						if Not fin_gera_nsu(T_COMISSAO_INDICADOR_N3, intNsuNovoComissaoN3, msg_erro) then 
							alerta = "FALHA AO GERAR NSU PARA O NOVO REGISTRO (" & msg_erro & ")"
						else
							if intNsuNovoComissaoN3 <= 0 then
								alerta = "NSU GERADO � INV�LIDO (" & intNsuNovoComissaoN3 & ")"
								end if
							end if
						rs.Open sql_n3, cn
						rs.AddNew 
						rs("id") = intNsuNovoComissaoN3
						rs("id_comissao_indicador_n2") = intNsuNovoComissaoN2
						rs("indicador") = Trim("" & rs2("indicador"))
				   
						for j=LBound(v_pedido) to UBound(v_pedido)
							marcado = false
							if (v_pedido(j) = v_lista_completa_pedidos(i)) then
								marcado = true
								Exit For
							end if
						Next
					
						if marcado=true then 
							rs("st_tratamento_manual") = 0 
						else
							ind_completo = rs2("indicador") & ","
							rs("st_tratamento_manual") = 1
							if instr(lista_indicador_com_desconto, ind_completo)>0 then 
								rs("cod_motivo_tratamento_manual") = 1
								rs("vl_total_descontos_planilha") = converte_numero(v_total_desconto_planilha(pos_v_planilha_desconto))
								rs("qtde_reg_descontos_planilha") = v_lista_qtde_reg_descontos(pos_v_planilha_desconto)
								tem_desconto=true
								pos_v_planilha_desconto = pos_v_planilha_desconto + 1
							elseif instr(lista_indicador_negativo, ind_completo)>0 then
								rs("cod_motivo_tratamento_manual") = 2
							else
								rs("cod_motivo_tratamento_manual") = 3
							end if
						end if
					
						rs("vl_total_comissao") = converte_numero(FormatNumber(v_lista_vl_total_comissao(n)))
						rs("vl_total_comissao_arredondado") = floor(converte_numero(FormatNumber(v_lista_vl_total_comissao_arredondado(n), 2)))
						rs("vl_total_RA") = converte_numero(v_lista_vl_total_RA(n))
						rs("vl_total_RA_arredondado") = floor(converte_numero(v_lista_vl_total_RA_arredondado(n)))
						rs("meio_pagto") = Trim(v_lista_meio_pagto(n))
						rs("banco") = Trim("" & rs2("banco"))
						rs("agencia") = Trim("" & rs2("agencia"))
						rs("conta") = Trim("" & rs2("conta"))
						rs("favorecido") = Trim("" & rs2("favorecido"))
						rs("favorecido_cnpj_cpf") = Trim("" & rs2("favorecido_cnpj_cpf"))
						rs("agencia_dv") = Trim("" & rs2("agencia_dv"))
						rs("conta_operacao") = Trim("" & rs2("conta_operacao"))
						rs("conta_dv") = Trim("" & rs2("conta_dv"))
						rs("tipo_conta") = Trim("" & rs2("tipo_conta"))
						rs("vl_total_pagto") = converte_numero(v_lista_vl_total_pagto(n))
						rs.Update
						 rs3.Open ("INSERT INTO t_COMISSAO_INDICADOR_N3_DESCONTO (" & _
									" id_comissao_indicador_n3," & _
									" id_orcamentista_e_indicador_desconto," & _
									" usuario," & _
									" dt_cadastro," & _
									" dt_hr_cadastro," & _
									" ordenacao," & _
									" descricao," & _
									" valor" & _
									" )" & _
									" SELECT" & _
									" "& intNsuNovoComissaoN3 &"," & _
									" id," & _
									" usuario," & _
									" dt_cadastro," & _
									" dt_hr_cadastro," & _
									" ordenacao," & _
									" descricao," & _
									" valor" & _ 
									" FROM t_ORCAMENTISTA_E_INDICADOR_DESCONTO" & _
									" WHERE apelido = '"& Trim("" & rs2("indicador")) &"'"),cn
						if rs.State <> 0 then rs.Close
						if rs3.State <> 0 then rs3.Close
						n = n + 1
					end if

					rs.Open sql_n4, cn
					rs.AddNew
					if Not fin_gera_nsu(T_COMISSAO_INDICADOR_N4, intNsuNovoComissaoN4, msg_erro) then 
							alerta = "FALHA AO GERAR NSU PARA O NOVO REGISTRO (" & msg_erro & ")"
						else
							if intNsuNovoComissaoN4 <= 0 then
								alerta = "NSU GERADO � INV�LIDO (" & intNsuNovoComissaoN4 & ")"
								end if
							end if
				
					rs("id") = intNsuNovoComissaoN4
					rs("id_comissao_indicador_n3") = intNsuNovoComissaoN3
					rs("pedido") = v_lista_completa_pedidos(i)
					rs("perc_RT") = rs2("perc_RT")
					rs("vl_pedido") = converte_numero(v_lista_vl_pedido(i))
					rs("vl_comissao") = converte_numero(FormatNumber(v_lista_vl_comissao(i),2))
					rs("vl_RA_bruto") = converte_numero(v_lista_vl_RA_bruto(i))
					rs("vl_RA_liq") = converte_numero(v_lista_vl_RA_liquido(i))
					rs("st_pagto") = "S"
					rs("tabela_origem") = v_lista_operacao(i)
					rs.Update
				
					if rs.State <> 0 then rs.Close
				
					if vendedor = "" then
						vendedor = vendedor & rs2("vendedor")
					else
						if rs2("vendedor") <> vendedor_a then 
						   vendedor = vendedor & "," & rs2("vendedor")
						end if
					end if
				
					vendedor_a = Trim("" & rs2("vendedor"))
					indicador_a = Trim("" & rs2("indicador"))
				
			end if
			
			   if rs2.State <> 0 then rs2.Close
			
		next
	
		'--- GRAVA O LOG 
		   mensagem = "VENDEDOR(ES) ESCOLHIDO(S): " & vendedor & "; "&"M�s de competencia: " & mes_competencia & "/"& ano_competencia 
		   grava_log usuario,"","","", OP_LOG_REL_COMISSAO_INDICADORES_PAGAMENTO , mensagem

		if alerta="" then
			cn.CommitTrans       
		else
			cn.RollbackTrans
			Response.Write alerta
		end if
		
		if Err=0 then Response.Redirect("RelComissaoIndicadoresFinaliza.asp?id=" & intNsuNovoComissaoN1 & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")) & "&rb_visao=" & rb_visao)
		end if 'if alerta = ""

'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>