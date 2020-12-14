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
'	  RelAcompanhamentoChamadosExec.asp
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
	if Not operacao_permitida(OP_CEN_PEDIDO_CHAMADO_CADASTRAMENTO, s_lista_operacoes_permitidas) then 
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if
	
	dim blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos
	blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos = isActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos

	dim s_filtro, intQtdeChamados
	dim s, s_aux, rb_status, origem, rb_posicao, c_motivo_abertura
    dim c_dt_cad_chamado_inicio, c_dt_cad_chamado_termino
    dim blnHaSomenteFinalizados
	origem = ucase(Trim(request("origem")))
	intQtdeChamados = 0


	if origem="A" then
	'	PARÂMETRO VEM PELA QUERYSTRING
		rb_status = Trim(Request("rb_status"))
        rb_posicao = Trim(Request("rb_posicao"))
        c_dt_cad_chamado_inicio=Trim(Request("c_dt_cad_chamado_inicio"))
	    c_dt_cad_chamado_termino=Trim(Request("c_dt_cad_chamado_termino"))
        c_motivo_abertura=Trim(Request("c_motivo_abertura"))
	else
		rb_status = Trim(Request.Form("rb_status"))
        rb_posicao = Trim(Request.Form("rb_posicao"))
        c_dt_cad_chamado_inicio=Trim(Request.Form("c_dt_cad_chamado_inicio"))
	    c_dt_cad_chamado_termino=Trim(Request.Form("c_dt_cad_chamado_termino"))
        c_motivo_abertura=Trim(Request.Form("c_motivo_abertura"))
		end if

    dim nivel_acesso_chamado
	nivel_acesso_chamado = Session("nivel_acesso_chamado")
	if Trim(nivel_acesso_chamado) = "" then
		nivel_acesso_chamado = obtem_nivel_acesso_chamado_pedido(cn, usuario)
		Session("nivel_acesso_chamado") = nivel_acesso_chamado
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


' _____________________________________________
' MOTIVO_FINALIZACAO_CHAMADO_MONTA_ITENS_SELECT

function motivo_finalizacao_chamado_monta_itens_select(byval depto, byval id_default)
dim x, r, strResp, strSql, idDefault

    idDefault = Trim("" & id_default)
    strSql = "SELECT * FROM t_CODIGO_DESCRICAO" & _
                " WHERE grupo='" & GRUPO_T_CODIGO_DESCRICAO__CHAMADOS_EM_PEDIDOS__MOTIVO_FINALIZACAO & "' AND codigo_pai = '" & Cstr(depto) & "' AND st_inativo=0" & _
                " ORDER BY ordenacao"

    set r = cn.Execute(strSql)
	strResp = "<option value='' selected>&nbsp;</option>"
	do while Not r.EOF 
        x = r("codigo")
        strResp = strResp & "<option"

        if idDefault <> "" then
            if idDefault = x then strResp = strResp & " selected"
        end if

	    strResp = strResp & " value='" & x & "'>"
        strResp = strResp & r("descricao")
        strResp = strResp & "</option>"
		r.MoveNext        
    loop

    motivo_finalizacao_chamado_monta_itens_select = strResp
	r.Close
	set r=nothing
end function



' _____________________________________
' CONSULTA EXECUTA
'
sub consulta_executa
dim s, s2, s_aux, s_sql, x
dim r
dim cab_table, cab
dim qtde_chamado_aberto, qtde_chamado_em_andamento, qtde_chamados_finalizados
dim s_disabled

	s_sql = _
		"SELECT" & _
			" tPC.id," & _
			" tPC.pedido," & _
			" tPC.usuario_cadastro," & _
			" tPC.dt_cadastro," & _
			" tPC.dt_hr_cadastro," & _
			" tPC.contato," & _
			" tPC.ddd_1," & _
			" tPC.tel_1," & _
			" tPC.ddd_2," & _
			" tPC.tel_2," & _
			" tPC.texto_chamado," & _
            " tPC.cod_motivo_abertura," & _
            " tPC.finalizado_status," & _
            " tPC.cod_motivo_finalizacao," & _
            " tPC.texto_finalizacao," & _
			" tP.transportadora_id,"

	if blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos then
		s_sql = s_sql & _
			" tP.endereco_nome_iniciais_em_maiusculas AS nome_cliente,"
	else
		s_sql = s_sql & _
			" tC.nome_iniciais_em_maiusculas AS nome_cliente,"
		end if

	s_sql = s_sql & _
			" tCD.codigo, tCD.descricao," & _
            " tPCD.descricao AS depto," & _
            " tPCD.id AS cod_depto," & _
            " tPCD.usuario_responsavel," & _
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
			") AS qtde_msg_rx," & _
            " (" & _
				"SELECT" & _
					" Count(*)" & _
				" FROM t_PEDIDO_CHAMADO_MENSAGEM" & _
				" WHERE" & _
					" (id_chamado=tPC.id)" & _
					" AND (usuario_cadastro ='" & usuario & "')" & _
			") AS qtde_msg_usuario" & _
	   " FROM t_PEDIDO_CHAMADO tPC" & _
			" INNER JOIN t_PEDIDO tP ON (tPC.pedido=tP.pedido)" & _
			" INNER JOIN t_CLIENTE tC ON (tP.id_cliente=tC.id)" & _
            " LEFT JOIN t_PEDIDO_CHAMADO_DEPTO tPCD ON (tPCD.id=tPC.id_depto)" & _
            " LEFT JOIN t_CODIGO_DESCRICAO tCD ON (tPC.cod_motivo_abertura=tCD.codigo) AND (tCD.grupo='" & GRUPO_T_CODIGO_DESCRICAO__CHAMADOS_EM_PEDIDOS__MOTIVO_ABERTURA & "')" & _
	   " WHERE" & _
			" (1 = 1)"

    if IsDate(c_dt_cad_chamado_inicio) then
	s_sql = s_sql & _
				" AND (tPC.dt_cadastro >= " & bd_formata_data(StrToDate(c_dt_cad_chamado_inicio)) & ")"
	end if

	if IsDate(c_dt_cad_chamado_termino) then
		s_sql = s_sql & _
				" AND (tPC.dt_cadastro < " & bd_formata_data(StrToDate(c_dt_cad_chamado_termino)+1) & ")"
	end if

    if c_motivo_abertura <> "" then
        s_sql = s_sql & " AND (tPC.cod_motivo_abertura = '" & c_motivo_abertura & "')"
    end if

	if rb_status = "ABERTO" then
		s_sql = "SELECT * FROM (" & s_sql & ") t WHERE ((qtde_msg_rx = 0) AND (finalizado_status = 0))"
	elseif rb_status = "EM_ANDAMENTO" then
		s_sql = "SELECT * FROM (" & s_sql & ") t WHERE ((qtde_msg_rx > 0) AND (finalizado_status = 0))"
    elseif rb_status = "FINALIZADO" then
		s_sql = "SELECT * FROM (" & s_sql & ") t WHERE (finalizado_status <> 0)"
	else
		s_sql = "SELECT * FROM (" & s_sql & ") t WHERE (1 = 1)"
		end if

    if rb_posicao = "USUARIO_TX" then
            s_sql = s_sql & " AND (t.usuario_cadastro = '" & usuario & "')"
        elseif rb_posicao = "USUARIO_RX" then
            s_sql = s_sql & " AND (t.usuario_responsavel = '" & usuario & "')"
        elseif rb_posicao = "USUARIO_MSG" then
            s_sql = s_sql & " AND (t.qtde_msg_usuario > 0)"
        elseif rb_posicao = "" then
            s_sql = s_sql & " AND ((t.usuario_cadastro = '" & usuario & "') OR (t.usuario_responsavel = '" & usuario & "') OR (t.qtde_msg_usuario > 0))"
        end if

    s_sql = s_sql & " ORDER BY dt_hr_cadastro, id"

	cab_table = "<TABLE cellSpacing=0 cellPadding=0>" & chr(13)
	cab = "	<TR style='background:azure' NOWRAP>" & chr(13) & _
		  "		<TD class='MDTE tdDataHora' style='vertical-align:bottom'><P class='Rc'>DT Chamado</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdDepartamento' style='vertical-align:bottom'><P class='R'>Depto Resp</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdUsuario' style='vertical-align:bottom'><P class='R'>Usuário</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdPedido' style='vertical-align:bottom'><P class='R'>Pedido</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdNF' style='vertical-align:bottom'><P class='Rc'>NF</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdCliente' style='vertical-align:bottom'><P class='R'>Cliente</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdTransp' style='vertical-align:bottom'><P class='R'>Transp</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdContato' style='vertical-align:bottom'><P class='R'>Contato</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdTel' style='vertical-align:bottom'><P class='R'>Telefone</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdChamado' style='vertical-align:bottom'><P class='R'>Motivo de Abertura</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdStatus' style='vertical-align:bottom'><P class='R'>Status</P></TD>" & chr(13) & _
		  "		<TD style='background:white;'>&nbsp;</TD>" & chr(13) & _
		  "	</TR>" & chr(13)
	
	x = cab_table & cab
	intQtdeChamados = 0
	qtde_chamado_aberto = 0
	qtde_chamado_em_andamento = 0
    qtde_chamados_finalizados = 0
	
	set r = cn.execute(s_sql)
	do while Not r.Eof
	
	 ' CONTAGEM
		intQtdeChamados = intQtdeChamados + 1

		x = x & "	<TR NOWRAP>" & chr(13)

	'> ID DO CHAMADO (HIDDEN)
		x = x & "		<input type=hidden name='c_id_chamado_" & Cstr(intQtdeChamados) & "' id='c_id_chamado_" & Cstr(intQtdeChamados) & "' value='" & Trim("" & r("id")) & "'>" & chr(13)

	'> Nº PEDIDO (HIDDEN)
		x = x & "		<input type=hidden name='c_pedido_" & Cstr(intQtdeChamados) & "' id='c_pedido_" & Cstr(intQtdeChamados) & "' value='" & Trim("" & r("pedido")) & "'>" & chr(13)
		
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

	'> CLIENTE
		s = Trim("" & r("nome_cliente"))
		if s = "" then s = "&nbsp;"
		x = x & "		<TD class='MTD tdCliente'><P class='Cn'>" & s & "</P></TD>" & chr(13)

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

	'> MOTIVO ABERTURA DO CHAMADO
		s = obtem_descricao_tabela_t_codigo_descricao(GRUPO_T_CODIGO_DESCRICAO__CHAMADOS_EM_PEDIDOS__MOTIVO_ABERTURA, Trim("" & r("cod_motivo_abertura")))
		if s = "" then s = "&nbsp;"
		x = x & "		<TD class='MTD tdChamado'><P class='Cn'>" & s & "</P></TD>" & chr(13)

	'> STATUS
        if CInt(r("finalizado_status")) = 0 then
		    if CInt(r("qtde_msg_rx")) > 0 then
			    s = "Em Andamento"
			    qtde_chamado_em_andamento = qtde_chamado_em_andamento + 1
		    else
			    s = "Aberto"
			    qtde_chamado_aberto = qtde_chamado_aberto + 1
			end if
        else
            s = "Finalizado"
            qtde_chamados_finalizados = qtde_chamados_finalizados + 1
        end if

		if s = "" then s = "&nbsp;"
		x = x & "		<TD class='MTD tdStatus'><P class='Cn'>" & s & "</P></TD>" & chr(13)

	'> BOTÃO P/ EXIBIR DEMAIS CAMPOS
		x = x & "		<TD valign='bottom' class='notPrint'>" & _
							"&nbsp;<a name='bExibeOcultaCampos' id='bExibeOcultaCampos' href='javascript:fExibeOcultaCampos(" & chr(34) & Cstr(intQtdeChamados) & chr(34) & ")' title='exibe ou oculta os campos adicionais'><img src='../botao/view_bottom.png' border='0'></a>" & _
						"</TD>" & chr(13)
		
		x = x & "	</TR>" & chr(13)

    '> DESCRIÇÃO DO CHAMADO
        x = x & "	<TR style='display:none;' id='TR_DESCRICAO_" & Cstr(intQtdeChamados) & "'>" & chr(13) & _
				"		<TD class='ME MD'>&nbsp;</TD>" & chr(13) & _
				"		<TD colspan='10' class='MC MD'>" & chr(13) & _
				"			<table width='100%' cellspacing='0' cellpadding='0'>" & chr(13) & _
				"				<TR>" & chr(13) & _
				"					<td class='Rf tdWithPadding'>DESCRIÇÃO DO CHAMADO</td>" & chr(13) & _
				"					<td>&nbsp;</td>" & chr(13) & _
				"					<td>&nbsp;</td>" & chr(13) & _
				"				</TR>" & chr(13) & _
				"				<TR>" & chr(13) & _
				"					<TD colspan='4'>" & chr(13) & _
				"						<table width='100%' cellSpacing='0'>" & chr(13) & _
				"							<TR>" & chr(13) & _
				"								<td>" & chr(13) & _
				"                                   <textarea name='c_descricao_" & Cstr(intQtdeChamados) & "' id='c_descricao_" & Cstr(intQtdeChamados) & "' class='PLLe' rows='7' style='width:100%;margin-left:2pt;' disabled>" & Trim("" & r("texto_chamado")) & "</textarea>" & chr(13) & _
				"								</td>" & chr(13) & _
				"							</TR>" & chr(13) & _
				"						</table>" & chr(13) & _
				"					</TD>" & chr(13) & _
				"				</TR>" & chr(13) & _
				"			</table>" & chr(13) & _
				"		</TD>" & chr(13) & _
				"	</TR>" & chr(13)

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
				"		<TD colspan='10' class='MC MD'>" & chr(13) & _
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
		        "                                <TD class='Cn MC MD' align='center' valign='top' style='width:50px;'>" & nivel_acesso_chamado_pedido_descricao(rs("nivel_acesso")) & "</TD>" & chr(13) & _
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

	'> NOVA MENSAGEM
        if CInt(r("finalizado_status")) = 0 then 
		    x = x & "	<TR style='display:none;' id='TR_NEW_MSG_" & Cstr(intQtdeChamados) & "'>" & chr(13) & _
				    "		<TD class='ME MD'>&nbsp;</TD>" & chr(13) & _
				    "		<TD colspan='10' class='MC MD'>" & chr(13) & _
				    "			<table width='100%' cellspacing='0' cellpadding='0'>" & chr(13) & _
				    "				<TR>" & chr(13) & _
				    "					<td>&nbsp;</td>" & chr(13) & _
				    "					<td class='Rf tdWithPadding'>NOVA MENSAGEM</td>" & chr(13) & _
				    "					<td align='right' valign='bottom'>" & chr(13) & _
										    "<span class='PLLd' style='font-weight:normal;'>Tamanho restante:</span><input name='c_tamanho_restante_nova_msg_" & Cstr(intQtdeChamados) & "' id='c_tamanho_restante_nova_msg_" & Cstr(intQtdeChamados) & "' tabindex=-1 readonly class='TA' style='width:35px;text-align:right;font-size:8pt;font-weight:normal;' value='" & Cstr(MAX_TAM_MENSAGEM_CHAMADOS_EM_PEDIDOS) & "' />" & chr(13) & _
				    "					</td>" & chr(13) & _
				    "				</TR>" & chr(13) & _
				    "				<TR>" & chr(13) & _
				    "					<TD colspan='4'>" & chr(13) & _
				    "						<table width='100%' cellSpacing='0'>" & chr(13) & _
				    "							<TR>" & chr(13) & _
				    "								<td>" & chr(13) & _
													    "<textarea name='c_nova_msg_" & Cstr(intQtdeChamados) & "' id='c_nova_msg_" & Cstr(intQtdeChamados) & "' class='PLLe' rows='3' style='width:100%;margin-left:2pt;' onkeypress='limita_tamanho(this,MAX_TAM_MENSAGEM_CHAMADOS_EM_PEDIDOS);' onblur='this.value=trim(this.value);calcula_tamanho_restante_nova_msg(" & chr(34) & Cstr(intQtdeChamados) & chr(34) & ");' onkeyup='calcula_tamanho_restante_nova_msg(" & chr(34) & Cstr(intQtdeChamados) & chr(34) & ");'></textarea>" & _
				    "								</td>" & chr(13) & _
				    "							</TR>" & chr(13) & _
				    "						</table>" & chr(13) & _
				    "					</TD>" & chr(13) & _
				    "				</TR>" & chr(13) & _
                    "				<TR>" & chr(13) & _
				    "					<TD colspan='4'>" & chr(13) & _
				    "						<table width='100%' cellSpacing='0'>" & chr(13) & _
				    "							<TR>" & chr(13) & _
				    "								<td>" & chr(13) & _
				    "									<p class='Rf'>NÍVEL DE ACESSO PARA LEITURA</p>" & chr(13) & _
		            "                                   <select id='c_nivel_acesso_chamado_" & Cstr(intQtdeChamados) & "' name='c_nivel_acesso_chamado_" & Cstr(intQtdeChamados) & "' style='margin-top:3px;margin-bottom:4px;margin-left:2pt;width:180px;' onkeyup='if (window.event.keyCode==KEYCODE_DELETE) {this.options[0].selected=true;}'>" & chr(13) & _
		                                                nivel_acesso_chamado_pedido_monta_itens_select(Null, nivel_acesso_chamado, True) & chr(13) & _
		            "                                   </select>" & chr(13) & _
				    "								</td>" & chr(13) & _
				    "							</TR>" & chr(13) & _
				    "						</table>" & chr(13) & _
				    "					</TD>" & chr(13) & _
				    "				</TR>" & chr(13) & _
				    "			</table>" & chr(13) & _
				    "		</TD>" & chr(13) & _
				    "	</TR>" & chr(13)
        else
            x = x & "	<TR style='display:none;' id='TR_NEW_MSG_" & Cstr(intQtdeChamados) & "'>" & chr(13) & _
				    "		<TD class='ME MD'>&nbsp;</TD>" & chr(13) & _
				    "		<TD colspan='10' class='MC MD'>&nbsp;</TD>" & chr(13) & _
				    "	</TR>" & chr(13)
        end if

	'> MOTIVO FINALIZAÇÃO
        s_disabled = ""
        if Trim("" & r("finalizado_status")) = "1" then s_disabled = " disabled"
		x = x & "	<TR style='display:none;' id='TR_MOTIVO_FINALIZACAO_" & Cstr(intQtdeChamados) & "'>" & chr(13) & _
				"		<TD class='ME MD'>&nbsp;</TD>" & chr(13) & _
				"		<TD colspan='10' class='MC MD'>" & chr(13) & _
				"			<table width='100%' cellspacing='0' cellpadding='0'>" & chr(13) & _
				"				<TR>" & chr(13) & _
				"					<td class='Rf tdWithPadding'>MOTIVO DA FINALIZAÇÃO</td>" & chr(13) & _
				"				</TR>" & chr(13) & _
				"				<TR>" & chr(13) & _
				"					<TD>" & chr(13) & _
				"                       <select id='c_motivo_finalizacao_" & Cstr(intQtdeChamados) & "' name='c_motivo_finalizacao_" & Cstr(intQtdeChamados) & "' style='width:600px;margin-left:4pt;margin-top:4pt;margin-bottom:4pt;' onkeyup='if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;'" & s_disabled & ">" & chr(13) & _
										motivo_finalizacao_chamado_monta_itens_select(Trim("" & r("cod_depto")), Trim("" & r("cod_motivo_finalizacao"))) & chr(13) & _
				"                       </select>" & chr(13) & _
				"					</TD>" & chr(13) & _
				"				</TR>" & chr(13) & _
				"			</table>" & chr(13) & _
				"		</TD>" & chr(13) & _
				"	</TR>" & chr(13)

	'> SOLUÇÃO
        s_disabled = ""
        if Trim("" & r("finalizado_status")) then s_disabled = " disabled"
		x = x & "	<TR style='display:none;' id='TR_SOLUCAO_" & Cstr(intQtdeChamados) & "'>" & chr(13) & _
				"		<TD class='ME MD'>&nbsp;</TD>" & chr(13) & _
				"		<TD colspan='10' class='MC MD'>" & chr(13) & _
				"			<table width='100%' cellspacing='0' cellpadding='0'>" & chr(13) & _
				"				<TR>" & chr(13) & _
				"					<td class='Rf tdWithPadding'>SOLUÇÃO</td>" & chr(13) & _
				"					<td>&nbsp;</td>" & chr(13) & _
				"					<td align='right' valign='bottom'>" & chr(13) & _
				"                       <span class='PLLd' style='font-weight:normal;'>Tamanho restante:</span><input name='c_tamanho_restante_solucao_" & Cstr(intQtdeChamados) & "' id='c_tamanho_restante_solucao_" & Cstr(intQtdeChamados) & "' tabindex=-1 readonly class='TA' style='width:35px;text-align:right;font-size:8pt;font-weight:normal;' value='" & Cstr(MAX_TAM_DESCRICAO_CHAMADO_EM_PEDIDOS) & "' />" & chr(13) & _
				"					</td>" & chr(13) & _
				"				</TR>" & chr(13) & _
				"				<TR>" & chr(13) & _
				"					<TD colspan='4'>" & chr(13) & _
				"						<table width='100%' cellSpacing='0'>" & chr(13) & _
				"							<TR>" & chr(13) & _
				"								<td>" & chr(13) & _
				"                                   <textarea name='c_solucao_" & Cstr(intQtdeChamados) & "' id='c_solucao_" & Cstr(intQtdeChamados) & "' class='PLLe' rows='7' style='width:100%;margin-left:2pt;' onkeypress='limita_tamanho(this,MAX_TAM_DESCRICAO_CHAMADO_EM_PEDIDOS);' onblur='this.value=trim(this.value);calcula_tamanho_restante_solucao(" & chr(34) & Cstr(intQtdeChamados) & chr(34) & ");' onkeyup='calcula_tamanho_restante_solucao(" & chr(34) & Cstr(intQtdeChamados) & chr(34) & ");'" & s_disabled & ">" & Trim("" & r("texto_finalizacao")) & "</textarea>" & chr(13) & _
				"								</td>" & chr(13) & _
				"							</TR>" & chr(13) & _
				"						</table>" & chr(13) & _
				"					</TD>" & chr(13) & _
				"				</TR>" & chr(13) & _
				"			</table>" & chr(13) & _
				"		</TD>" & chr(13) & _
				"	</TR>" & chr(13)

        ' MOTIVO DE ABERTURA (HIDDEN)
        x = x  & "  <input type='hidden' name='c_motivo_abertura_" & Cstr(intQtdeChamados) & "' id='c_motivo_abertura_" & Cstr(intQtdeChamados) & "' value='" & r("cod_motivo_abertura") & "' />" & chr(13)
	    x = x & "   <input type='hidden' name='c_finalizado_status_" & CStr(intQtdeChamados) & "' id='c_finalizado_status_" & CStr(intQtdeChamados) &	"' value='" & r("finalizado_status") & "' />" & chr(13)

		if (intQtdeChamados mod 100) = 0 then
			Response.Write x
			x = ""
			end if
			
		r.MoveNext
		loop
	
    ' verifica se a consulta retornou apenas chamados finalizados para ocultar o botão de 'Confirmar'
    blnHaSomenteFinalizados = False
    if CInt(qtde_chamado_aberto)+CInt(qtde_chamado_em_andamento) = 0 then
        if CInt(qtde_chamados_finalizados) > 0 then blnHaSomenteFinalizados = True
    end if
        
	
'	TOTAL GERAL
	if intQtdeChamados > 0 then
		x = x & "	<TR>" & chr(13) & _
				"		<TD COLSPAN='11' class='MC' style='border-left:0px;border-right:0px;'>&nbsp;</TD>" & chr(13) & _
				"	</TR>" & chr(13) & _
				"	<TR NOWRAP style='background:honeydew'>" & chr(13) & _
				"		<TD COLSPAN='11' class='MT'><p class='C'>TOTAL: &nbsp; " & cstr(qtde_chamado_aberto+qtde_chamado_em_andamento+qtde_chamados_finalizados) & " chamado(s)</p></TD>" & chr(13) & _
				"	</TR>" & chr(13)
		end if
	
  ' MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	if intQtdeChamados = 0 then
		x = cab_table & cab
		x = x & "	<TR NOWRAP>" & chr(13) & _
				"		<TD class='MT' colspan='11'><P class='ALERTA'>&nbsp;NENHUM REGISTRO ENCONTRADO&nbsp;</P></TD>" & chr(13) & _
				"	</TR>" & chr(13)
		end if

  ' FECHA TABELA
	x = x & "</TABLE>" & chr(13)
	
	x = x & "<input type=HIDDEN name='c_qtde_chamados' id='c_qtde_chamados' value='" & Cstr(intQtdeChamados) & "'>" & chr(13)

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
window.status = 'Aguarde, executando a consulta...';

function calcula_tamanho_restante_nova_msg(indice_row) {
var ctr, cnm, s;
	ctr = document.getElementById("c_tamanho_restante_nova_msg_" + indice_row);
	cnm = document.getElementById("c_nova_msg_" + indice_row);
	s = "" + cnm.value;
	ctr.value = MAX_TAM_MENSAGEM_CHAMADOS_EM_PEDIDOS - s.length;
}

function calcula_tamanho_restante_solucao(indice_row) {
	var ctr, cnm, s;
	ctr = document.getElementById("c_tamanho_restante_solucao_" + indice_row);
	cnm = document.getElementById("c_solucao_" + indice_row);
	s = "" + cnm.value;
	ctr.value = MAX_TAM_DESCRICAO_CHAMADO_EM_PEDIDOS - s.length;
}

function fExibeOcultaCampos(indice_row) {
var row_MSGS, row_NEW_MSG, row_SOLUCAO, row_MOTIVO_FINALIZACAO, row_DESCRICAO;

	row_MSGS = document.getElementById("TR_MSGS_" + indice_row);
	row_NEW_MSG = document.getElementById("TR_NEW_MSG_" + indice_row);
	row_SOLUCAO = document.getElementById("TR_SOLUCAO_" + indice_row);
	row_DESCRICAO = document.getElementById("TR_DESCRICAO_" + indice_row);
	row_MOTIVO_FINALIZACAO = document.getElementById("TR_MOTIVO_FINALIZACAO_" + indice_row);

	if (row_MSGS.style.display.toString() == "none") {
		row_MSGS.style.display = "";
		row_NEW_MSG.style.display = "";
		row_SOLUCAO.style.display = "";
		row_DESCRICAO.style.display = "";
		row_MOTIVO_FINALIZACAO.style.display = "";
	}
	else {
		row_MSGS.style.display = "none";
		row_NEW_MSG.style.display = "none";
		row_SOLUCAO.style.display = "none";
		row_DESCRICAO.style.display = "none";
		row_MOTIVO_FINALIZACAO.style.display = "none";
	}
}

function fPEDConsulta(id_pedido) {
	window.status = "Aguarde ...";
	fPED.pedido_selecionado.value = id_pedido;
	fPED.action = "pedido.asp"
	fPED.submit();
}

function fRELConfirma(f) {
var cqo, cto, cs, cp, i, n, s, cmsg, cmsgna, fs;
	cqo = document.getElementById("c_qtde_chamados");
	n = parseInt(cqo.value);
	for (i = 1; i <= n; i++) {
	    cto = document.getElementById("c_motivo_finalizacao_" + i.toString());
		cs = document.getElementById("c_solucao_" + i.toString());
		cp = document.getElementById("c_pedido_" + i.toString());
		cmsg = document.getElementById("c_nova_msg_" + i.toString());
		cmsgna = document.getElementById("c_nivel_acesso_chamado_" + i.toString());
		fs = document.getElementById("c_finalizado_status_" + i.toString());
		if ((trim(cto.value) != "") || (trim(cs.value))) {
			//if (trim(cto.value) == "") {
			//	s = "Não foi selecionado o motivo da finalização para o pedido " + cp.value + "!!\nAo finalizar um chamado, é necessário informar o motivo da finalização e o texto descrevendo a solução.";
			//	alert(s);
			//	return;
			//}
			if (trim(cs.value) == "") {
				s = "Não foi informado o texto descrevendo a solução do chamado do pedido " + cp.value + "!!\nAo finalizar um chamado, é necessário informar o motivo da finalização e o texto descrevendo a solução.";
				alert(s);
				return;
			}
		}

		if (fs.value == "0") {
		    if (cmsg.value != "") {
		        if (cmsgna.value == "") {
		            s = "Não foi informado o nível de acesso para a mensagem do pedido " + cp.value + "!!";
		            alert(s);
		            cmsgna.focus();
		            return;
		        }
		    }
		}
	}
	dCONFIRMA.style.visibility = "hidden";
	window.status = "Aguarde ...";
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
.tdDepartamento{
    vertical-align:top;
    width: 85px;
}
.tdUsuario{
    vertical-align:top;
    width: 85px;
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
.tdCliente{
	vertical-align: top;
	width: 120px;
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
.tdChamado{
	vertical-align: top;
	width: 280px;
	}
.tdStatus{
	vertical-align: top;
	width: 90px;
	}
.tdDataHoraMsg{
	vertical-align: top;
	width: 63px;
	}
.tdUsuarioMsg{
	vertical-align: top;
	width: 80px;
	}
.tdTextoMensagem{
	vertical-align: top;
	width: 785px;
	}
</style>


<body onload="window.status='Concluído';focus();" link=#000000 alink=#000000 vlink=#000000>
<center>

<form id="fPED" name="fPED" method="post">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="pedido_selecionado" id="pedido_selecionado" value="">
</form>


<form id="fREL" name="fREL" method="post" action="RelPedidoChamadoGravaDados.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="rb_status" id="rb_status" value="<%=rb_status%>">
<input type="hidden" name="rb_posicao" id="rb_posicao" value="<%=rb_posicao %>" />
<input type="hidden" name="c_rel_origem" id="c_rel_origem" value="ACOMPANHAMENTO_CHAMADOS" />
<input type="hidden" name="c_dt_cad_chamado_inicio" id="c_dt_cad_chamado_inicio" value="<%=c_dt_cad_chamado_inicio%>" />
<input type="hidden" name="c_dt_cad_chamado_termino" id="c_dt_cad_chamado_termino" value="<%=c_dt_cad_chamado_termino%>" />
<input type="hidden" name="c_motivo_abertura" id="c_motivo_abertura" value="<%=c_motivo_abertura%>" />

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="1024" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Acompanhamento de Chamados</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>

<!-- FILTROS -->
<% 
	s_filtro = "<table width='1024' cellPadding='0' CellSpacing='0' style='border-bottom:1px solid black' border='0'>" & chr(13)

	s = ""
	s_aux = c_dt_cad_chamado_inicio
	if s_aux = "" then s_aux = "N.I."
	s = s & s_aux & " a "
	s_aux = c_dt_cad_chamado_termino
	if s_aux = "" then s_aux = "N.I."
	s = s & s_aux
	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' NOWRAP><p class='N'>Período abertura chamado:&nbsp;</p></td>" & chr(13) & _
				"		<td valign='top' width='99%'><p class='N'>" & s & "</p></td>" & chr(13) & _
				"	</tr>" & chr(13)

	s = rb_status
	if s = "ABERTO" then
		s = "Aberto"
	elseif s = "EM_ANDAMENTO" then
		s = "Em Andamento"
    elseif s = "FINALIZADO" then
		s = "Finalizado"
	elseif s = "" then
		s = "Aberto, Em Andamento, Finalizado"
	else
		s = "Parâmetro Desconhecido"
		end if

	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' NOWRAP>" & _
				"<p class='N'>Status do chamado:&nbsp;</p></td><td valign='top'>" & _
				"<p class='N'>" & s & "</p></td>" & chr(13) & _
				"	</tr>" & chr(13)

    s = rb_posicao
	if s = "USUARIO_TX" then
		s = "Abertos por mim"
	elseif s = "USUARIO_RX" then
		s = "Destinados ao meu departamento"
    elseif s = "USUARIO_MSG" then
		s = "Em que interagi com mensagens"
	elseif s = "" then
		s = "Todos"
	else
		s = "Parâmetro Desconhecido"
		end if

	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' NOWRAP>" & _
				"<p class='N'>Selecionar chamados:&nbsp;</p></td><td valign='top'>" & _
				"<p class='N'>" & s & "</p></td>" & chr(13) & _
				"	</tr>" & chr(13)

    s = c_motivo_abertura
    if s = "" then
        s = "Todos"
    else
        s = iniciais_em_maiusculas(obtem_descricao_tabela_t_codigo_descricao(GRUPO_T_CODIGO_DESCRICAO__CHAMADOS_EM_PEDIDOS__MOTIVO_ABERTURA, c_motivo_abertura))
    end if

    s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' NOWRAP>" & _
				"<p class='N'>Motivo da Abertura:&nbsp;</p></td><td valign='top'>" & _
				"<p class='N'>" & s & "</p></td>" & chr(13) & _
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

<% consulta_executa %>

<!-- ************   SEPARADOR   ************ -->
<table width="1024" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>
<br>


<table class="notPrint" width="1024" cellSpacing="0">
<tr>
	<td><a name="bVOLTAR" id="bVOLTAR"
		<% if origem="A" then %>
			href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>"
		<% else %>
			href="javascript:history.back()"
		<% end if %>
	title="volta para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
	<td align="right"><div name="dCONFIRMA" id="dCONFIRMA">
        <% if Not blnHaSomenteFinalizados then %>
        <a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fRELConfirma(fREL)" title="executa a consulta">
		<img src="../botao/confirmar.gif" width="176" height="55" border="0"></a></div>
        <% else %>
        &nbsp;
        <%end if %>
	</td>
</tr>
</table>
</form>

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
