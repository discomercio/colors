<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/Global.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ========================================================
'	  RelPedidoChamadoEstatisticasExec.asp
'     ========================================================
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

	dim usuario
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO)

'	CONECTA COM O BANCO DE DADOS
	dim cn, rs, msg_erro
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	if Not operacao_permitida(OP_CEN_REL_ESTATISTICAS_PEDIDO_CHAMADO, s_lista_operacoes_permitidas) then 
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if
	
	dim blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos
	blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos = isActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos

	dim alerta
	dim s, s_aux
	dim c_dt_periodo_inicio, c_dt_periodo_termino, c_motivo_finalizacao, c_motivo_abertura
	dim c_transportadora, c_vendedor, c_indicador, c_uf
	dim s_nome_vendedor, s_nome_indicador, s_nome_transportadora
	dim c_loja, lista_loja, s_filtro_loja, v_loja, v, i
    dim blnHasDepto, c_depto

	alerta = ""

	c_dt_periodo_inicio=Trim(Request("c_dt_cad_chamado_inicio"))
	c_dt_periodo_termino=Trim(Request("c_dt_cad_chamado_termino"))
	c_motivo_finalizacao=Trim(Request("c_motivo_finalizacao"))
    c_motivo_abertura=Trim(Request("c_motivo_abertura"))
	c_transportadora = Trim(Request.Form("c_transportadora"))
	c_vendedor = Ucase(Trim(Request.Form("c_vendedor")))
	c_indicador = Ucase(Trim(Request.Form("c_indicador")))
	c_uf = Ucase(Trim(Request.Form("c_uf")))
    blnHasDepto = CBool(Request.Form("blnHasDepto"))
    c_depto = Trim(Request.Form("c_depto"))
	
	c_loja = Trim(Request.Form("c_loja"))
	lista_loja = substitui_caracteres(c_loja,chr(10),"")
	v_loja = split(lista_loja,chr(13),-1)
	
	dim s_filtro, intQtdeChamados
	intQtdeChamados = 0
	
	if alerta = "" then
		s_nome_transportadora = ""
		if c_transportadora <> "" then
			s = "SELECT nome FROM t_TRANSPORTADORA WHERE (id='" & c_transportadora & "')"
			if rs.State <> 0 then rs.Close
			rs.open s, cn
			if rs.Eof then 
				alerta=texto_add_br(alerta)
				alerta = "TRANSPORTADORA " & c_transportadora & " NÃO ESTÁ CADASTRADA."
			else
				s_nome_transportadora = iniciais_em_maiusculas(Trim("" & rs("nome")))
				end if
			end if
		end if

	if alerta = "" then
		s_nome_vendedor = ""
		if c_vendedor <> "" then
			s = "SELECT nome_iniciais_em_maiusculas FROM t_USUARIO WHERE (usuario='" & c_vendedor & "')"
			if rs.State <> 0 then rs.Close
			rs.open s, cn
			if rs.Eof then 
				alerta=texto_add_br(alerta)
				alerta = "VENDEDOR " & c_vendedor & " NÃO ESTÁ CADASTRADO."
			else
				s_nome_vendedor = Trim("" & rs("nome_iniciais_em_maiusculas"))
				end if
			end if
		end if

	if alerta = "" then
		s_nome_indicador = ""
		if c_indicador <> "" then
			s = "SELECT razao_social_nome_iniciais_em_maiusculas FROM t_ORCAMENTISTA_E_INDICADOR WHERE (apelido='" & c_indicador & "')"
			if rs.State <> 0 then rs.Close
			rs.open s, cn
			if rs.Eof then 
				alerta=texto_add_br(alerta)
				alerta = "INDICADOR " & c_indicador & " NÃO ESTÁ CADASTRADO."
			else
				s_nome_indicador = Trim("" & rs("razao_social_nome_iniciais_em_maiusculas"))
				end if
			end if
		end if


    if Not operacao_permitida(OP_CEN_PEDIDO_CHAMADO_CONSULTA_CHAMADOS_TODOS_DEPTOS, s_lista_operacoes_permitidas) And _
    Not blnHasDepto then
        alerta = "Usuário não consta como gestor ou responsável por algum departamento e não possui acesso à consulta de chamado de todos os departamentos."
    end if


' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________

function monta_link_pedido(byval id_pedido)
dim strLink
	monta_link_pedido = ""
	id_pedido = Trim("" & id_pedido)
	if id_pedido = "" then exit function
	strLink = "<a href='javascript:fPEDConsulta(" & _
				chr(34) & id_pedido & chr(34) & _
				")' title='clique para consultar o pedido " & id_pedido & "'>" & _
				id_pedido & "</a>"
	monta_link_pedido=strLink
end function


' _____________________________________
' CONSULTA EXECUTA
'
sub consulta_executa
dim s, s2, s_aux, s_sql, x
dim s_where_loja
dim r
dim cab_table, cab
dim qtde_chamado_aberto, qtde_chamado_em_andamento, qtde_chamado_finalizado

'	MONTA SQL
	s_sql = _
		"SELECT" & _
			" tPC.id," & _
			" tPC.pedido," & _
			" tPC.finalizado_status," & _
			" tPC.finalizado_data_hora," & _
			" tPC.usuario_cadastro," & _
			" tPC.dt_cadastro," & _
			" tPC.dt_hr_cadastro," & _
			" tPC.contato," & _
			" tPC.ddd_1," & _
			" tPC.tel_1," & _
			" tPC.ddd_2," & _
			" tPC.tel_2," & _
			" tPC.texto_chamado," & _
			" tPC.cod_motivo_finalizacao," & _
            " tPC.cod_motivo_abertura," & _
			" tPC.texto_finalizacao," & _
			" tP.transportadora_id,"

	if blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos then
		s_sql = s_sql & _
				" tP.endereco_nome_iniciais_em_maiusculas AS nome_cliente, "
	else
		s_sql = s_sql & _
				" tC.nome_iniciais_em_maiusculas AS nome_cliente, "
		end if

	s_sql = s_sql & _
            " tPCD.descricao AS depto," & _
			" (" & _
				"SELECT" & _
					" TOP 1 NFe_numero_NF" & _
				" FROM t_NFe_EMISSAO tNE" & _
				" WHERE" & _
					" (tNE.pedido=tPC.pedido)" & _
					" AND (tipo_NF = '1')" & _
					" AND (st_anulado = 0)" & _
					" AND (codigo_retorno_NFe_T1 = 1)" & _
				" ORDER BY" & _
					" id DESC" & _
			") AS numeroNFe," & _
			" (" & _
				"SELECT" & _
					" Count(*)" & _
				" FROM t_PEDIDO_CHAMADO_MENSAGEM" & _
				" WHERE" & _
					" (id_chamado=tPC.id)" & _
					" AND (fluxo_mensagem='" & COD_FLUXO_MENSAGEM_CHAMADOS_EM_PEDIDOS__RX & "')" & _
			") AS qtde_msg_rx" & _
		" FROM t_PEDIDO_CHAMADO tPC" & _
			" INNER JOIN t_PEDIDO tP ON (tPC.pedido=tP.pedido)" & _
			" INNER JOIN t_CLIENTE tC ON (tP.id_cliente=tC.id)" & _
            " LEFT JOIN t_PEDIDO_CHAMADO_DEPTO tPCD ON (tPCD.id=tPC.id_depto)" & _
		" WHERE" & _
			" (tPC.finalizado_status <> 0)"

    if c_depto <> "" then
        s_sql = s_sql & " AND (tPC.id_depto = '" & c_depto & "')"
    else
        if Not operacao_permitida(OP_CEN_PEDIDO_CHAMADO_CONSULTA_CHAMADOS_TODOS_DEPTOS, s_lista_operacoes_permitidas) then
            s_sql = s_sql & " AND tPC.id_depto IN (SELECT id FROM t_PEDIDO_CHAMADO_DEPTO WHERE usuario_responsavel = '" & usuario & "' OR usuario_gestor = '" & usuario & "')"
        end if
    end if

	if IsDate(c_dt_periodo_inicio) then
		s_sql = s_sql & _
					" AND (tPC.dt_cadastro >= " & bd_formata_data(StrToDate(c_dt_periodo_inicio)) & ")"
		end if

	if IsDate(c_dt_periodo_termino) then
		s_sql = s_sql & _
					" AND (tPC.dt_cadastro < " & bd_formata_data(StrToDate(c_dt_periodo_termino)) & ")"
		end if
	
	if c_motivo_finalizacao <> "" then
		s_sql = s_sql & _
					" AND (tPC.cod_motivo_finalizacao = '" & c_motivo_finalizacao & "')"
		end if

    if c_motivo_abertura <> "" then
		s_sql = s_sql & _
					" AND (tPC.cod_motivo_abertura = '" & c_motivo_abertura & "')"
		end if

	if c_transportadora <> "" then
		s_sql = s_sql & _
					" AND (tP.transportadora_id = '" & c_transportadora & "')"
		end if

	if c_vendedor <> "" then
		s_sql = s_sql & _
					" AND (tP.vendedor = '" & c_vendedor & "')"
		end if
	
	if c_indicador <> "" then
		s_sql = s_sql & _
					" AND (tP.indicador = '" & c_indicador & "')"
		end if
	
	if c_uf <> "" then
		if blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos then
			s_sql = s_sql & _
						" AND " & _
						"(" & _
							"((tP.st_end_entrega <> 0) And (tP.EndEtg_uf = '" & c_uf & "'))" & _
							" OR " & _
							"((tP.st_end_entrega = 0) And (tP.endereco_uf = '" & c_uf & "'))" & _
						")"
		else
			s_sql = s_sql & _
						" AND " & _
						"(" & _
							"((tP.st_end_entrega <> 0) And (tP.EndEtg_uf = '" & c_uf & "'))" & _
							" OR " & _
							"((tP.st_end_entrega = 0) And (tC.uf = '" & c_uf & "'))" & _
						")"
			end if
		end if
	
'	LOJA(S)
	s_where_loja = ""
	for i=Lbound(v_loja) to Ubound(v_loja)
		if v_loja(i) <> "" then
			v = split(v_loja(i),"-",-1)
			if Ubound(v)=Lbound(v) then
				if s_where_loja <> "" then s_where_loja = s_where_loja & " OR"
				s_where_loja = s_where_loja & " (tP.numero_loja = " & v_loja(i) & ")"
			else
				s = ""
				if v(Lbound(v))<>"" then 
					if s <> "" then s = s & " AND"
					s = s & " (tP.numero_loja >= " & v(Lbound(v)) & ")"
					end if
				if v(Ubound(v))<>"" then
					if s <> "" then s = s & " AND"
					s = s & " (tP.numero_loja <= " & v(Ubound(v)) & ")"
					end if
				if s <> "" then 
					if s_where_loja <> "" then s_where_loja = s_where_loja & " OR"
					s_where_loja = s_where_loja & " (" & s & ")"
					end if
				end if
			end if
		next
	
	if s_where_loja <> "" then 
		s_sql = s_sql & _
					" AND (" & s_where_loja & ")"
		end if
	
	s_sql = "SELECT * FROM (" & s_sql & ") t ORDER BY dt_hr_cadastro, id"

	cab_table = "<TABLE cellSpacing=0 cellPadding=0>" & chr(13)
	cab = "	<TR style='background:azure' NOWRAP>" & chr(13) & _
		  "		<TD class='MDTE tdDataHora' style='vertical-align:bottom'><P class='Rc'>DT Chamado</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdDepartamento' style='vertical-align:bottom'><P class='R'>Depto Resp</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdUsuario' style='vertical-align:bottom'><P class='R'>Usuário</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdPedido' style='vertical-align:bottom'><P class='R'>Pedido</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdNF' style='vertical-align:bottom'><P class='Rc'>NF</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdTransp' style='vertical-align:bottom'><P class='R'>Transp</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdContato' style='vertical-align:bottom'><P class='R'>Contato</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdTel' style='vertical-align:bottom'><P class='R'>Telefone</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdDescricao' style='vertical-align:bottom'><P class='R'>Descrição Chamado</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdSolucao' style='vertical-align:bottom'><P class='R'>Solução</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdMotivoFinalizacao' style='vertical-align:bottom'><P class='R'>Motivo Finalização</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdStatus' style='vertical-align:bottom'><P class='R'>Status</P></TD>" & chr(13) & _
		  "		<TD style='background:white;'>&nbsp;</TD>" & chr(13) & _
		  "	</TR>" & chr(13)
	
	x = cab_table & cab
	intQtdeChamados = 0
	qtde_chamado_aberto = 0
	qtde_chamado_em_andamento = 0
	qtde_chamado_finalizado = 0
	
	set r = cn.execute(s_sql)
	do while Not r.Eof
	
	 ' CONTAGEM
		intQtdeChamados = intQtdeChamados + 1

		x = x & "	<TR NOWRAP>" & chr(13)

	'> DATA DO CHAMADO
		s = formata_data_hora_sem_seg(r("dt_hr_cadastro"))
		x = x & "		<TD class='MDTE tdDataHora'><P class='Cnc'>" & s & "</P></TD>" & chr(13)

    '> DEPTO RESPONSAVEL
		s = Trim("" & r("depto"))
		x = x & "		<TD class='MTD tdDepartamento'><P class='Cn'>" & s & "</P></TD>" & chr(13)

    '> USUÁRIO
		s = Trim("" & r("usuario_cadastro"))
		x = x & "		<TD class='MTD tdUsuario'><P class='Cn'>" & s & "</P></TD>" & chr(13)

	'> PEDIDO
		s = monta_link_pedido(Trim("" & r("pedido")))
		x = x & "		<TD class='MTD tdPedido'><P class='Cn'>" & s & "</P></TD>" & chr(13)

	'> NF
		s = Trim("" & r("numeroNFe"))
		if s = "" then s = "&nbsp;"
		x = x & "		<TD class='MTD tdNF'><P class='Cnc'>" & s & "</P></TD>" & chr(13)

	'> TRANSPORTADORA
		s = Trim("" & r("transportadora_id"))
		if s = "" then s = "&nbsp;"
		x = x & "		<TD class='MTD tdTransp'><P class='Cn'>" & s & "</P></TD>" & chr(13)

	'> CONTATO
		s = iniciais_em_maiusculas(Trim("" & r("contato")))
		if s = "" then s = "&nbsp;"
		x = x & "		<TD class='MTD tdContato'><P class='Cn'>" & s & "</P></TD>" & chr(13)

	'> TELEFONE(S)
		s = ""
		s2 = Trim("" & r("tel_1"))
		if s2 <> "" then
			s2 = telefone_formata(s2)
			s_aux = Trim("" & r("ddd_1"))
			if s_aux <> "" then s2 = "(" & s_aux & ")" & s2
			if s <> "" then s = s & "<br>"
			s = s & s2
			end if
		s2 = Trim("" & r("tel_2"))
		if s2 <> "" then
			s2 = telefone_formata(s2)
			s_aux = Trim("" & r("ddd_2"))
			if s_aux <> "" then s2 = "(" & s_aux & ")" & s2
			if s <> "" then s = s & "<br>"
			s = s & s2
			end if

		if s = "" then s = "&nbsp;"
		x = x & "		<TD class='MTD tdTel'><P class='Cn'>" & s & "</P></TD>" & chr(13)

	'> DESCRIÇÃO DO CHAMADO
        s = Trim("" & r("texto_chamado"))
		if Trim("" & r("cod_motivo_abertura")) = "" then 
            x = x & "		<TD class='MTD tdDescricao'><P class='Cn'>" & s & "</P></TD>" & chr(13)
        else
			s = iniciais_em_maiusculas(obtem_descricao_tabela_t_codigo_descricao(GRUPO_T_CODIGO_DESCRICAO__CHAMADOS_EM_PEDIDOS__MOTIVO_ABERTURA, Trim("" & r("cod_motivo_abertura")))) & "<br>" & substitui_caracteres(Trim("" & r("texto_chamado")), chr(13), "<br>")
            x = x & "		<TD class='MTD tdDescricao'><P class='Cn'>" & s & "</P></TD>" & chr(13)            
	    end if
		
	'> SOLUÇÃO
		s = substitui_caracteres(Trim("" & r("texto_finalizacao")), chr(13), "<br>")
		if s = "" then s = "&nbsp;"
		x = x & "		<TD class='MTD tdSolucao'><P class='Cn'>" & s & "</P></TD>" & chr(13)

	'> MOTIVO DA FINALIZAÇÃO DO CHAMADO
		s = Trim("" & r("cod_motivo_finalizacao"))
		if s <> "" then 
			s = iniciais_em_maiusculas(obtem_descricao_tabela_t_codigo_descricao(GRUPO_T_CODIGO_DESCRICAO__CHAMADOS_EM_PEDIDOS__MOTIVO_FINALIZACAO, Trim("" & r("cod_motivo_finalizacao"))))
			end if
		if s = "" then s = "&nbsp;"
		x = x & "		<TD class='MTD tdTransp'><P class='Cn'>" & s & "</P></TD>" & chr(13)

	'> STATUS
		if CInt(r("finalizado_status")) <> 0 then
			s = "Finalizado"
			qtde_chamado_finalizado = qtde_chamado_finalizado + 1
		else
			if CInt(r("qtde_msg_rx")) > 0 then
				s = "Em Andamento"
				qtde_chamado_em_andamento = qtde_chamado_em_andamento + 1
			else
				s = "Aberta"
				qtde_chamado_aberto = qtde_chamado_aberto + 1
				end if
			end if

		if s = "" then s = "&nbsp;"
		x = x & "		<TD class='MTD tdStatus'><P class='Cn'>" & s & "</P></TD>" & chr(13)

	'> BOTÃO P/ EXIBIR DEMAIS CAMPOS
		x = x & "		<TD valign='bottom' class='notPrint'>" & _
							"&nbsp;<a name='bExibeOcultaCampos' id='bExibeOcultaCampos' href='javascript:fExibeOcultaCampos(" & chr(34) & Cstr(intQtdeChamados) & chr(34) & ")' title='exibe ou oculta os campos adicionais'><img src='../botao/view_bottom.png' border='0'></a>" & _
						"</TD>" & chr(13)
		
		x = x & "	</TR>" & chr(13)

	'> MENSAGENS
		s_sql = _
			"SELECT " & _
				"*" & _
			" FROM t_PEDIDO_CHAMADO_MENSAGEM" & _
			" WHERE" & _
				" (id_chamado = " & Trim("" & r("id")) & ")" & _
			" ORDER BY" & _
				" dt_hr_cadastro," & _
				" id"
		if rs.State <> 0 then rs.Close
		rs.open s_sql, cn
		x = x & "	<TR style='display:none;' id='TR_MSGS_" & Cstr(intQtdeChamados) & "'>" & chr(13) & _
				"		<TD class='ME MD'>&nbsp;</TD>" & chr(13) & _
				"		<TD colspan='11' class='MC MD'>" & chr(13) & _
				"			<table width='100%' cellspacing='0' cellpadding='0'>" & chr(13) & _
				"				<TR>" & chr(13) & _
				"					<td class='Rf tdWithPadding'>MENSAGENS</td>" & chr(13) & _
				"				</TR>" & chr(13)
		if rs.Eof then
			x = x & _
				"				<TR>" & chr(13) & _
				"					<td>&nbsp;</td>" & chr(13) & _
				"				</TR>" & chr(13)
			end if
		
		do while Not rs.Eof
			x = x & _
				"				<TR>" & chr(13) & _
				"					<TD>" & chr(13) & _
				"						<table width='100%' cellSpacing='0' cellPadding='0'>" & chr(13) & _
				"							<TR>" & chr(13) & _
				"								<TD class='Cn MD MC tdWithPadding tdDataHoraMsg' align='center'>" & chr(13) & _
													formata_data_hora_sem_seg(rs("dt_hr_cadastro")) & _
				"								</TD>" & chr(13) & _
				"								<TD class='Cn MD MC tdWithPadding tdUsuarioMsg' align='center'>" & chr(13) & _
													rs("usuario_cadastro")
			if Trim("" & rs("loja")) <> "" then x = x & " (Loja&nbsp;" & Trim("" & rs("loja")) & ")"
			x = x & _
				"								</TD>" & chr(13) & _
				"								<TD class='Cn MC tdWithPadding tdTextoMensagem' align='left' valign='top'>" & chr(13) & _
													substitui_caracteres(Trim("" & rs("texto_mensagem")), chr(13), "<br>") & _
												"</TD>" & chr(13) & _
				"							</TR>" & chr(13) & _
				"						</table>" & chr(13) & _
				"					</TD>" & chr(13) & _
				"				</TR>" & chr(13)
			rs.MoveNext
			loop
			
		x = x & _
				"			</table>" & chr(13) & _
				"		</TD>" & chr(13) & _
				"	</TR>" & chr(13)

		if (intQtdeChamados mod 100) = 0 then
			Response.Write x
			x = ""
			end if
			
		r.MoveNext
		loop
	
	
'	TOTAL GERAL
	if intQtdeChamados > 0 then
		x = x & "	<TR>" & chr(13) & _
				"		<TD COLSPAN='12' class='MC' style='border-left:0px;border-right:0px;'>&nbsp;</TD>" & chr(13) & _
				"	</TR>" & chr(13) & _
				"	<TR NOWRAP style='background:honeydew'>" & chr(13) & _
				"		<TD COLSPAN='12' class='MT'><p class='C'>TOTAL: &nbsp; " & cstr(qtde_chamado_aberto+qtde_chamado_em_andamento+qtde_chamado_finalizado) & " chamado(s)</p></TD>" & chr(13) & _
				"	</TR>" & chr(13)
		end if
	
  ' MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	if intQtdeChamados = 0 then
		x = cab_table & cab
		x = x & "	<TR NOWRAP>" & chr(13) & _
				"		<TD class='MT' colspan='12'><P class='ALERTA'>&nbsp;NENHUM REGISTRO ENCONTRADO&nbsp;</P></TD>" & chr(13) & _
				"	</TR>" & chr(13)
		end if

  ' FECHA TABELA
	x = x & "</TABLE>" & chr(13)
	
	Response.write x

	if r.State <> 0 then r.Close
	set r=nothing
	
end sub

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
window.status = 'Aguarde, executando a consulta ...';

function fExibeOcultaCampos(indice_row) {
var row_MSGS;

	row_MSGS = document.getElementById("TR_MSGS_" + indice_row);
	if (row_MSGS.style.display.toString() == "none") {
		row_MSGS.style.display = "";
	}
	else {
		row_MSGS.style.display = "none";
	}
}

function fPEDConsulta(id_pedido) {
	window.status = "Aguarde ...";
	fPED.pedido_selecionado.value = id_pedido;
	fPED.action = "pedido.asp"
	fPED.submit();
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
<link href="<%=URL_FILE__EPRINTER_CSS%>" Rel="stylesheet" Type="text/css" media="print">
<link href="<%=URL_FILE__ESCREEN_CSS%>" Rel="stylesheet" Type="text/css" media="screen">

<style type="text/css">
html 
{
	overflow-y: scroll;
}
.tdWithPadding
{
	padding:1px;
}
.tdDataHora{
	vertical-align: top;
	width: 65px;
	}
.tdPedido{
	vertical-align: top;
	font-weight: bold;
	width: 65px;
	}
.tdNF{
	vertical-align: top;
	width: 60px;
	}
.tdTransp{
	vertical-align: top;
	width: 90px;
	}
.tdContato{
	vertical-align: top;
	width: 100px;
	}
.tdTel{
	vertical-align: top;
	width: 90px;
	}
.tdDescricao{
	vertical-align: top;
	width: 170px;
	}
.tdSolucao{
	vertical-align: top;
	width: 170px;
	}
.tdMotivoFinalizacao{
	vertical-align: top;
	width: 101px;
	}
.tdStatus{
	vertical-align: top;
	width: 60px;
	}
.tdDataHoraMsg{
	vertical-align: top;
	width: 63px;
	}
.tdUsuarioMsg{
	vertical-align: top;
	width: 70px;
	}
.tdTextoMensagem{
	vertical-align: top;
	width: 785px;
	}
.tdDepartamento{
    vertical-align:top;
    width: 85px;
}
.tdUsuario{
    vertical-align:top;
    width: 85px;
}
</style>


<body onload="window.status='Concluído';focus();" link=#000000 alink=#000000 vlink=#000000>
<center>

<form id="fPED" name="fPED" method="post">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="pedido_selecionado" id="pedido_selecionado" value="">
</form>


<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="1024" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Estatísticas de Chamados</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>

<!-- FILTROS -->
<% 
	s_filtro = "<table width='1024' cellPadding='0' CellSpacing='0' style='border-bottom:1px solid black' border='0'>" & chr(13)

'	PERÍODO DE ABERTURA DO CHAMADO
	s = ""
	s_aux = c_dt_periodo_inicio
	if s_aux = "" then s_aux = "N.I."
	s = s & s_aux & " a "
	s_aux = c_dt_periodo_termino
	if s_aux = "" then s_aux = "N.I."
	s = s & s_aux
	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' NOWRAP><p class='N'>Período Abertura Chamado:&nbsp;</p></td>" & chr(13) & _
				"		<td valign='top' width='99%'><p class='N'>" & s & "</p></td>" & chr(13) & _
				"	</tr>" & chr(13)

'	MOTIVO ABERTURA
	s = c_motivo_abertura
	if s = "" then
		s = "todos"
	else
		s = iniciais_em_maiusculas(obtem_descricao_tabela_t_codigo_descricao(GRUPO_T_CODIGO_DESCRICAO__CHAMADOS_EM_PEDIDOS__MOTIVO_ABERTURA, c_motivo_abertura))
		end if
	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' NOWRAP><p class='N'>Motivo Abertura:&nbsp;</p></td>" & chr(13) & _
				"		<td valign='top'><p class='N'>" & s & "</p></td>" & chr(13) & _
				"	</tr>" & chr(13)

'	MOTIVO DA FINALIZAÇÃO
	s = c_motivo_finalizacao
	if s = "" then
		s = "todos"
	else
		s = iniciais_em_maiusculas(obtem_descricao_tabela_t_codigo_descricao(GRUPO_T_CODIGO_DESCRICAO__CHAMADOS_EM_PEDIDOS__MOTIVO_FINALIZACAO, c_motivo_finalizacao))
		end if
	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' NOWRAP><p class='N'>Motivo Finalização:&nbsp;</p></td>" & chr(13) & _
				"		<td valign='top'><p class='N'>" & s & "</p></td>" & chr(13) & _
				"	</tr>" & chr(13)

	s = c_transportadora
	if s = "" then 
		s = "todas"
	else
		if (s_nome_transportadora <> "") And (s_nome_transportadora <> c_transportadora) then s = s & "  (" & s_nome_transportadora & ")"
		end if
	s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
				"<p class='N'>Transportadora:&nbsp;</p></td><td valign='top'>" & _
				"<p class='N'>" & s & "</p></td></tr>"

	s = c_vendedor
	if s = "" then 
		s = "todos"
	else
		if (s_nome_vendedor <> "") And (s_nome_vendedor <> c_vendedor) then s = s & "  (" & s_nome_vendedor & ")"
		end if
	s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
				"<p class='N'>Vendedor:&nbsp;</p></td><td valign='top'>" & _
				"<p class='N'>" & s & "</p></td></tr>"

	s = c_indicador
	if s = "" then 
		s = "todos"
	else
		if (s_nome_indicador <> "") And (s_nome_indicador <> c_indicador) then s = s & "  (" & s_nome_indicador & ")"
		end if
	s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
				"<p class='N'>Indicador:&nbsp;</p></td><td valign='top'>" & _
				"<p class='N'>" & s & "</p></td></tr>"

	s = c_uf
	if s = "" then s = "todos"
	s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
				"<p class='N'>UF:&nbsp;</p></td><td valign='top'>" & _
				"<p class='N'>" & s & "</p></td></tr>"

'	LOJA(S)
	s_filtro_loja = ""
	for i = Lbound(v_loja) to Ubound(v_loja)
		if v_loja(i) <> "" then
			v = split(v_loja(i),"-",-1)
			if Ubound(v)=Lbound(v) then
				if s_filtro_loja <> "" then s_filtro_loja = s_filtro_loja & ", "
				s_filtro_loja = s_filtro_loja & v_loja(i)
			else
				if (v(Lbound(v))<>"") And (v(Ubound(v))<>"") then 
					if s_filtro_loja <> "" then s_filtro_loja = s_filtro_loja & ", "
					s_filtro_loja = s_filtro_loja & v(Lbound(v)) & " a " & v(Ubound(v))
				elseif (v(Lbound(v))<>"") And (v(Ubound(v))="") then
					if s_filtro_loja <> "" then s_filtro_loja = s_filtro_loja & ", "
					s_filtro_loja = s_filtro_loja & v(Lbound(v)) & " e acima"
				elseif (v(Lbound(v))="") And (v(Ubound(v))<>"") then
					if s_filtro_loja <> "" then s_filtro_loja = s_filtro_loja & ", "
					s_filtro_loja = s_filtro_loja & v(Ubound(v)) & " e abaixo"
					end if
				end if
			end if
		next
	s = s_filtro_loja
	if s = "" then s = "todas"
	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' NOWRAP><p class='N'>Loja(s):&nbsp;</p></td>" & chr(13) & _
				"		<td valign='top'><p class='N'>" & s & "</p></td>" & chr(13) & _
				"	</tr>" & chr(13)

	s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='top' NOWRAP>" & _
					"<p class='N'>Emissão:&nbsp;</p></td><td valign='top' width='99%'>" & _
					"<p class='N'>" & formata_data_hora(Now) & "</p></td>" & chr(13) & _
					"	</tr>" & chr(13)
	
	s_filtro = s_filtro & "</table>" & chr(13)
	Response.Write s_filtro
%>

<!--  RELATÓRIO  -->
<br>

<% if alerta = "" then    
     consulta_executa
   else %>
    <div class="MtAlerta" style="width:400px;FONT-WEIGHT:bold;" align="CENTER">
        <P style='margin:5px 2px 5px 2px;'><%=alerta%></P>
    </div>
<% end if %>

<!-- ************   SEPARADOR   ************ -->
<table width="1024" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>
<br>


<table class="notPrint" width="1024" cellSpacing="0">
<tr>
	<td align="center"><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()" title="volta para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a>
	</td>
</tr>
</table>

</center>
</body>
</html>


<%
	if rs.State <> 0 then rs.Close
	set rs = nothing
	
'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>
