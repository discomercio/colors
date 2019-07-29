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
'	  RelPedidoOcorrencia.asp
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
	if Not operacao_permitida(OP_CEN_REL_OCORRENCIAS_EM_PEDIDOS, s_lista_operacoes_permitidas) then 
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if
	
	dim s_filtro, intQtdeOcorrencias
	dim s, rb_status, origem, c_loja
	origem = ucase(Trim(request("origem")))
	intQtdeOcorrencias = 0

	if origem="A" then
	'	PARÂMETRO VEM PELA QUERYSTRING
		rb_status = Trim(Request("rb_status"))
		c_loja = retorna_so_digitos(Trim(Request("c_loja")))
	else
		rb_status = Trim(Request.Form("rb_status"))
		c_loja = retorna_so_digitos(Trim(Request.Form("c_loja")))
		end if

	dim alerta
	alerta = ""

	if alerta = "" then
		if c_loja <> "" then
			s = "SELECT loja FROM t_LOJA WHERE (loja = '" & c_loja & "')"
			if rs.State <> 0 then rs.Close
			rs.open s, cn
			if rs.Eof then
				alerta = "LOJA Nº " & c_loja & " NÃO ESTÁ CADASTRADA."
				end if
			end if
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
dim r
dim cab_table, cab
dim qtde_ocorrencia_aberta, qtde_ocorrencia_em_andamento
dim s_link_rastreio

	s_sql = _
		"SELECT" & _
			" tPO.id," & _
			" tPO.pedido," & _
			" tPO.usuario_cadastro," & _
			" tPO.dt_cadastro," & _
			" tPO.dt_hr_cadastro," & _
			" tPO.contato," & _
			" tPO.ddd_1," & _
			" tPO.tel_1," & _
			" tPO.ddd_2," & _
			" tPO.tel_2," & _
			" tPO.texto_ocorrencia," & _
			" tP.loja," & _
            " tP.loja AS pedido_loja," & _
			" tP.transportadora_id," & _
			" tC.nome_iniciais_em_maiusculas AS nome_cliente, tCD.codigo, tCD.descricao," & _
			" (" & _
				"SELECT" & _
					" TOP 1 NFe_numero_NF" & _
				" FROM t_NFe_EMISSAO tNE" & _
				" WHERE" & _
					" (tNE.pedido=tPO.pedido)" & _
					" AND (tipo_NF = '1')" & _
					" AND (st_anulado = 0)" & _
					" AND (codigo_retorno_NFe_T1 = 1)" & _
				" ORDER BY" & _
					" id DESC" & _
			") AS numeroNFe," & _
			" (" & _
				"SELECT" & _
					" Count(*)" & _
				" FROM t_PEDIDO_OCORRENCIA_MENSAGEM" & _
				" WHERE" & _
					" (id_ocorrencia=tPO.id)" & _
					" AND (fluxo_mensagem='" & COD_FLUXO_MENSAGEM_OCORRENCIAS_EM_PEDIDOS__CENTRAL_PARA_LOJA & "')" & _
			") AS qtde_msg_central," & _
            " (" & _
                " SELECT Count(*)" & _
		           " FROM t_PEDIDO_OCORRENCIA_MENSAGEM INNER JOIN t_PEDIDO_OCORRENCIA ON (t_PEDIDO_OCORRENCIA_MENSAGEM.id_ocorrencia=t_PEDIDO_OCORRENCIA.id)" & _
                   " INNER JOIN t_PEDIDO ON (t_PEDIDO_OCORRENCIA.pedido=t_PEDIDO.pedido)" & _ 
		           " WHERE (id_ocorrencia = tPO.id)" & _
                   " AND (t_PEDIDO.loja = '" & NUMERO_LOJA_ECOMMERCE_AR_CLUBE & "')" & _ 
            ") AS qtde_msg" & _
	   " FROM t_PEDIDO_OCORRENCIA tPO" & _
			" INNER JOIN t_PEDIDO tP ON (tPO.pedido=tP.pedido)" & _
			" INNER JOIN t_CLIENTE tC ON (tP.id_cliente=tC.id)" & _
            " LEFT JOIN t_CODIGO_DESCRICAO tCD ON (tPO.cod_motivo_abertura=tCD.codigo) AND (tCD.grupo='" & GRUPO_T_CODIGO_DESCRICAO__OCORRENCIAS_EM_PEDIDOS__MOTIVO_ABERTURA & "')" & _
	   " WHERE" & _
			" (finalizado_status = 0)"
	
	if c_loja <> "" then
		s_sql = s_sql & " AND (tP.numero_loja = " & c_loja & ")"
		end if

	if rb_status = "ABERTA" then
		s_sql = "SELECT * FROM (" & s_sql & ") t WHERE (qtde_msg_central = 0 AND qtde_msg = 0) ORDER BY dt_hr_cadastro, id"
	elseif rb_status = "EM_ANDAMENTO" then
		s_sql = "SELECT * FROM (" & s_sql & ") t WHERE (qtde_msg_central > 0 OR qtde_msg > 0) ORDER BY dt_hr_cadastro, id"
	else
		s_sql = "SELECT * FROM (" & s_sql & ") t ORDER BY dt_hr_cadastro, id"
		end if

	cab_table = "<TABLE cellSpacing=0 cellPadding=0>" & chr(13)
	cab = "	<TR style='background:azure' NOWRAP>" & chr(13) & _
		  "		<TD class='MDTE tdDataHora' style='vertical-align:bottom'><P class='Rc'>DT Ocorr</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdLoja' style='vertical-align:bottom'><P class='R'>Loja</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdPedido' style='vertical-align:bottom'><P class='R'>Pedido</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdNF' style='vertical-align:bottom'><P class='Rc'>NF</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdCliente' style='vertical-align:bottom'><P class='R'>Cliente</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdTransp' style='vertical-align:bottom'><P class='R'>Transp</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdContato' style='vertical-align:bottom'><P class='R'>Contato</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdTel' style='vertical-align:bottom'><P class='R'>Telefone</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdOcorrencia' style='vertical-align:bottom'><P class='R'>Ocorrência</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdStatus' style='vertical-align:bottom'><P class='R'>Status</P></TD>" & chr(13) & _
		  "		<TD style='background:white;'>&nbsp;</TD>" & chr(13) & _
		  "	</TR>" & chr(13)
	
	x = cab_table & cab
	intQtdeOcorrencias = 0
	qtde_ocorrencia_aberta = 0
	qtde_ocorrencia_em_andamento = 0
	
	set r = cn.execute(s_sql)
	do while Not r.Eof
	
	 ' CONTAGEM
		intQtdeOcorrencias = intQtdeOcorrencias + 1

		x = x & "	<TR NOWRAP>" & chr(13)

	'> ID DA OCORRÊNCIA (HIDDEN)
		x = x & "		<input type=hidden name='c_id_ocorrencia_" & Cstr(intQtdeOcorrencias) & "' id='c_id_ocorrencia_" & Cstr(intQtdeOcorrencias) & "' value='" & Trim("" & r("id")) & "'>" & chr(13)

	'> Nº PEDIDO (HIDDEN)
		x = x & "		<input type=hidden name='c_pedido_" & Cstr(intQtdeOcorrencias) & "' id='c_pedido_" & Cstr(intQtdeOcorrencias) & "' value='" & Trim("" & r("pedido")) & "'>" & chr(13)
		
	'> DATA DA OCORRÊNCIA
		s = formata_data_hora_sem_seg(r("dt_hr_cadastro"))
		x = x & "		<TD class='MDTE tdDataHora'><P class='Cnc'>" & s & "</P></TD>" & chr(13)

	'> LOJA
		x = x & "		<TD class='MTD tdLoja'><P class='Cn'>" & Trim("" & r("loja")) & "</P></TD>" & chr(13)

	'> PEDIDO
		s = monta_link_pedido(Trim("" & r("pedido")))
		x = x & "		<TD class='MTD tdPedido'><P class='Cn'>" & s & "</P></TD>" & chr(13)

	'> NF
        s_link_rastreio = ""
		s = Trim("" & r("numeroNFe"))
		if s = "" then
            s = "&nbsp;"
        else
            s_link_rastreio = monta_link_rastreio(Trim("" & r("pedido")), Trim("" & r("numeroNFe")), Trim("" & r("transportadora_id")), Trim("" & r("pedido_loja")))
        end if
        if s_link_rastreio <> "" then s_link_rastreio = "&nbsp;" & s_link_rastreio
		x = x & "		<TD class='MTD tdNF'><P class='Cnc'>" & s & s_link_rastreio & "</P></TD>" & chr(13)

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

	'> OCORRÊNCIA
        if Trim("" & r("codigo")) = "" then
		    s = substitui_caracteres(Trim("" & r("texto_ocorrencia")), chr(13), "<br>")
		    if s = "" then s = "&nbsp;"
		    x = x & "		<TD class='MTD tdOcorrencia'><P class='Cn'>" & s & "</P></TD>" & chr(13)
        else
            s = Trim("" & r("descricao"))
            if Trim("" & r("texto_ocorrencia")) <> "" then 
                x = x & "   <TD class='MTD tdOcorrencia'><P class='Cn'>" & s & "<br>" & substitui_caracteres(Trim("" & r("texto_ocorrencia")), chr(13), "<br>") & "</P></TD>" & chr(13)
            else
                x = x & "   <TD class='MTD tdOcorrencia'><P class='Cn'>" & s & "</P></TD>" & chr(13)
            end if
        end if

	'> STATUS
		if CInt(r("qtde_msg_central")) > 0 Or _
                        (Trim("" & r("pedido_loja")) = NUMERO_LOJA_ECOMMERCE_AR_CLUBE And CInt(r("qtde_msg")) > 0) then
			s = "Em Andamento"
			qtde_ocorrencia_em_andamento = qtde_ocorrencia_em_andamento + 1
		else
			s = "Aberta"
			qtde_ocorrencia_aberta = qtde_ocorrencia_aberta + 1
			end if

		if s = "" then s = "&nbsp;"
		x = x & "		<TD class='MTD tdStatus'><P class='Cn'>" & s & "</P></TD>" & chr(13)

	'> BOTÃO P/ EXIBIR DEMAIS CAMPOS
		x = x & "		<TD valign='bottom' class='notPrint'>" & _
							"&nbsp;<a name='bExibeOcultaCampos' id='bExibeOcultaCampos' href='javascript:fExibeOcultaCampos(" & chr(34) & Cstr(intQtdeOcorrencias) & chr(34) & ")' title='exibe ou oculta os campos adicionais'><img src='../botao/view_bottom.png' border='0'></a>" & _
						"</TD>" & chr(13)
		
		x = x & "	</TR>" & chr(13)

	'> MENSAGENS
		s_sql = _
			"SELECT " & _
				"*" & _
		   " FROM t_PEDIDO_OCORRENCIA_MENSAGEM" & _
		   " WHERE" & _
				" (id_ocorrencia = " & Trim("" & r("id")) & ")" & _
		   " ORDER BY" & _
				" dt_hr_cadastro," & _
				" id"
		if rs.State <> 0 then rs.Close
		rs.open s_sql, cn
		x = x & "	<TR style='display:none;' id='TR_MSGS_" & Cstr(intQtdeOcorrencias) & "'>" & chr(13) & _
				"		<TD class='ME MD'>&nbsp;</TD>" & chr(13) & _
				"		<TD colspan='9' class='MC MD'>" & chr(13) & _
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

	'> NOVA MENSAGEM
		x = x & "	<TR style='display:none;' id='TR_NEW_MSG_" & Cstr(intQtdeOcorrencias) & "'>" & chr(13) & _
				"		<TD class='ME MD'>&nbsp;</TD>" & chr(13) & _
				"		<TD colspan='9' class='MC MD'>" & chr(13) & _
				"			<table width='100%' cellspacing='0' cellpadding='0'>" & chr(13) & _
				"				<TR>" & chr(13) & _
				"					<td class='Rf tdWithPadding'>NOVA MENSAGEM</td>" & chr(13) & _
				"					<td>&nbsp;</td>" & chr(13) & _
				"					<td align='right' valign='bottom'>" & chr(13) & _
										"<span class='PLLd' style='font-weight:normal;'>Tamanho restante:</span><input name='c_tamanho_restante_nova_msg_" & Cstr(intQtdeOcorrencias) & "' id='c_tamanho_restante_nova_msg_" & Cstr(intQtdeOcorrencias) & "' tabindex=-1 readonly class='TA' style='width:35px;text-align:right;font-size:8pt;font-weight:normal;' value='" & Cstr(MAX_TAM_MENSAGEM_OCORRENCIAS_EM_PEDIDOS) & "' />" & chr(13) & _
				"					</td>" & chr(13) & _
				"				</TR>" & chr(13) & _
				"				<TR>" & chr(13) & _
				"					<TD colspan='3'>" & chr(13) & _
				"						<table width='100%' cellSpacing='0'>" & chr(13) & _
				"							<TR>" & chr(13) & _
				"								<td>" & chr(13) & _
													"<textarea name='c_nova_msg_" & Cstr(intQtdeOcorrencias) & "' id='c_nova_msg_" & Cstr(intQtdeOcorrencias) & "' class='PLLe' rows='3' style='width:100%;margin-left:2pt;' onkeypress='limita_tamanho(this,MAX_TAM_MENSAGEM_OCORRENCIAS_EM_PEDIDOS);' onblur='this.value=trim(this.value);calcula_tamanho_restante_nova_msg(" & chr(34) & Cstr(intQtdeOcorrencias) & chr(34) & ");' onkeyup='calcula_tamanho_restante_nova_msg(" & chr(34) & Cstr(intQtdeOcorrencias) & chr(34) & ");'></textarea>" & _
				"								</td>" & chr(13) & _
				"							</TR>" & chr(13) & _
				"						</table>" & chr(13) & _
				"					</TD>" & chr(13) & _
				"				</TR>" & chr(13) & _
				"			</table>" & chr(13) & _
				"		</TD>" & chr(13) & _
				"	</TR>" & chr(13)

	'> TIPO OCORRÊNCIA
		x = x & "	<TR style='display:none;' id='TR_TIPO_OCORR_" & Cstr(intQtdeOcorrencias) & "'>" & chr(13) & _
				"		<TD class='ME MD'>&nbsp;</TD>" & chr(13) & _
				"		<TD colspan='9' class='MC MD'>" & chr(13) & _
				"			<table width='100%' cellspacing='0' cellpadding='0'>" & chr(13) & _
				"				<TR>" & chr(13) & _
				"					<td class='Rf tdWithPadding'>TIPO DE OCORRÊNCIA</td>" & chr(13) & _
				"				</TR>" & chr(13) & _
				"				<TR>" & chr(13) & _
				"					<TD>" & chr(13) & _
										"<select id='c_tipo_ocorrencia_" & Cstr(intQtdeOcorrencias) & "' name='c_tipo_ocorrencia_" & Cstr(intQtdeOcorrencias) & "' style='width:600px;margin-left:4pt;margin-top:4pt;margin-bottom:4pt;' onkeyup='if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;'>" & chr(13) & _
										codigo_descricao_monta_itens_select(GRUPO_T_CODIGO_DESCRICAO__OCORRENCIAS_EM_PEDIDOS__TIPO_OCORRENCIA, "") & chr(13) & _
										"</select>" & chr(13) & _
				"					</TD>" & chr(13) & _
				"				</TR>" & chr(13) & _
				"			</table>" & chr(13) & _
				"		</TD>" & chr(13) & _
				"	</TR>" & chr(13)

	'> SOLUÇÃO
		x = x & "	<TR style='display:none;' id='TR_SOLUCAO_" & Cstr(intQtdeOcorrencias) & "'>" & chr(13) & _
				"		<TD class='ME MD'>&nbsp;</TD>" & chr(13) & _
				"		<TD colspan='9' class='MC MD'>" & chr(13) & _
				"			<table width='100%' cellspacing='0' cellpadding='0'>" & chr(13) & _
				"				<TR>" & chr(13) & _
				"					<td class='Rf tdWithPadding'>SOLUÇÃO</td>" & chr(13) & _
				"					<td>&nbsp;</td>" & chr(13) & _
				"					<td align='right' valign='bottom'>" & chr(13) & _
										"<span class='PLLd' style='font-weight:normal;'>Tamanho restante:</span><input name='c_tamanho_restante_solucao_" & Cstr(intQtdeOcorrencias) & "' id='c_tamanho_restante_solucao_" & Cstr(intQtdeOcorrencias) & "' tabindex=-1 readonly class='TA' style='width:35px;text-align:right;font-size:8pt;font-weight:normal;' value='" & Cstr(MAX_TAM_DESCRICAO_OCORRENCIAS_EM_PEDIDOS) & "' />" & chr(13) & _
				"					</td>" & chr(13) & _
				"				</TR>" & chr(13) & _
				"				<TR>" & chr(13) & _
				"					<TD colspan='3'>" & chr(13) & _
				"						<table width='100%' cellSpacing='0'>" & chr(13) & _
				"							<TR>" & chr(13) & _
				"								<td>" & chr(13) & _
													"<textarea name='c_solucao_" & Cstr(intQtdeOcorrencias) & "' id='c_solucao_" & Cstr(intQtdeOcorrencias) & "' class='PLLe' rows='3' style='width:100%;margin-left:2pt;' onkeypress='limita_tamanho(this,MAX_TAM_DESCRICAO_OCORRENCIAS_EM_PEDIDOS);' onblur='this.value=trim(this.value);calcula_tamanho_restante_solucao(" & chr(34) & Cstr(intQtdeOcorrencias) & chr(34) & ");' onkeyup='calcula_tamanho_restante_solucao(" & chr(34) & Cstr(intQtdeOcorrencias) & chr(34) & ");'></textarea>" & _
				"								</td>" & chr(13) & _
				"							</TR>" & chr(13) & _
				"						</table>" & chr(13) & _
				"					</TD>" & chr(13) & _
				"				</TR>" & chr(13) & _
				"			</table>" & chr(13) & _
				"		</TD>" & chr(13) & _
				"	</TR>" & chr(13)
		
		if (intQtdeOcorrencias mod 100) = 0 then
			Response.Write x
			x = ""
			end if
			
		r.MoveNext
		loop
	
	
'	TOTAL GERAL
	if intQtdeOcorrencias > 0 then
		x = x & "	<TR>" & chr(13) & _
				"		<TD COLSPAN='10' class='MC' style='border-left:0px;border-right:0px;'>&nbsp;</TD>" & chr(13) & _
				"	</TR>" & chr(13) & _
				"	<TR NOWRAP style='background:honeydew'>" & chr(13) & _
				"		<TD COLSPAN='10' class='MT'><p class='C'>TOTAL: &nbsp; " & cstr(qtde_ocorrencia_aberta+qtde_ocorrencia_em_andamento) & " ocorrência(s)</p></TD>" & chr(13) & _
				"	</TR>" & chr(13)
		end if
	
  ' MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	if intQtdeOcorrencias = 0 then
		x = cab_table & cab
		x = x & "	<TR NOWRAP>" & chr(13) & _
				"		<TD class='MT' colspan='10'><P class='ALERTA'>&nbsp;NENHUM REGISTRO ENCONTRADO&nbsp;</P></TD>" & chr(13) & _
				"	</TR>" & chr(13)
		end if

  ' FECHA TABELA
	x = x & "</TABLE>" & chr(13)
	
	x = x & "<input type=HIDDEN name='c_qtde_ocorrencias' id='c_qtde_ocorrencias' value='" & Cstr(intQtdeOcorrencias) & "'>" & chr(13)

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
<script src="<%=URL_FILE__JQUERY%>" language="JavaScript" type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
window.status = 'Aguarde, executando a consulta...';

function calcula_tamanho_restante_nova_msg(indice_row) {
var ctr, cnm, s;
	ctr = document.getElementById("c_tamanho_restante_nova_msg_" + indice_row);
	cnm = document.getElementById("c_nova_msg_" + indice_row);
	s = "" + cnm.value;
	ctr.value = MAX_TAM_MENSAGEM_OCORRENCIAS_EM_PEDIDOS - s.length;
}

function calcula_tamanho_restante_solucao(indice_row) {
	var ctr, cnm, s;
	ctr = document.getElementById("c_tamanho_restante_solucao_" + indice_row);
	cnm = document.getElementById("c_solucao_" + indice_row);
	s = "" + cnm.value;
	ctr.value = MAX_TAM_DESCRICAO_OCORRENCIAS_EM_PEDIDOS - s.length;
}

function fExibeOcultaCampos(indice_row) {
var row_MSGS, row_NEW_MSG, row_SOLUCAO, row_TIPO_OCORR;

	row_MSGS = document.getElementById("TR_MSGS_" + indice_row);
	row_NEW_MSG = document.getElementById("TR_NEW_MSG_" + indice_row);
	row_SOLUCAO = document.getElementById("TR_SOLUCAO_" + indice_row);
	row_TIPO_OCORR = document.getElementById("TR_TIPO_OCORR_" + indice_row);

	if (row_MSGS.style.display.toString() == "none") {
		row_MSGS.style.display = "";
		row_NEW_MSG.style.display = "";
		row_SOLUCAO.style.display = "";
		row_TIPO_OCORR.style.display = "";
	}
	else {
		row_MSGS.style.display = "none";
		row_NEW_MSG.style.display = "none";
		row_SOLUCAO.style.display = "none";
		row_TIPO_OCORR.style.display = "none";
	}
}

function fPEDConsulta(id_pedido) {
	window.status = "Aguarde ...";
	fPED.pedido_selecionado.value = id_pedido;
	fPED.action = "pedido.asp"
	fPED.submit();
}

function fRELConfirma(f) {
var cqo, cto, cs, cp, i, n, s;
	cqo = document.getElementById("c_qtde_ocorrencias");
	n = parseInt(cqo.value);
	for (i = 1; i <= n; i++) {
		cto = document.getElementById("c_tipo_ocorrencia_" + i.toString());
		cs = document.getElementById("c_solucao_" + i.toString());
		cp = document.getElementById("c_pedido_" + i.toString());
		if ((trim(cto.value) != "") || (trim(cs.value))) {
			if (trim(cto.value) == "") {
				s = "Não foi selecionado o tipo de ocorrência para o pedido " + cp.value + "!!\nAo finalizar uma ocorrência, é necessário informar o tipo de ocorrência e o texto descrevendo a solução.";
				alert(s);
				return;
			}
			if (trim(cs.value) == "") {
				s = "Não foi informado o texto descrevendo a solução da ocorrência do pedido " + cp.value + "!!\nAo finalizar uma ocorrência, é necessário informar o tipo de ocorrência e o texto descrevendo a solução.";
				alert(s);
				return;
			}
		}
	}
	dCONFIRMA.style.visibility = "hidden";
	window.status = "Aguarde ...";
	f.submit();
}
</script>
<script type="text/javascript">
    $(document).ready(function () {
        $("#divRastreioConsultaView").hide();
        $('#divInternoRastreioConsultaView').addClass('divFixo');
        sizeDivRastreioConsultaView();
        $(document).keyup(function (e) {
            if (e.keyCode == 27) {
                fechaDivRastreioConsultaView();
            }
        });
        $("#divRastreioConsultaView").click(function () {
            fechaDivRastreioConsultaView();
        });
        $("#imgFechaDivRastreioConsultaView").click(function () {
            fechaDivRastreioConsultaView();
        });
    });
    //Every resize of window
    $(window).resize(function () {
        sizeDivRastreioConsultaView();
    });
    function fRastreioConsultaView(url) {
        sizeDivRastreioConsultaView();
        $("#divRastreioConsultaView").fadeIn();
        frame = document.getElementById("iframeRastreioConsultaView");
        frame.contentWindow.location.replace(url);
    }
    function fechaDivRastreioConsultaView() {
        $("#divRastreioConsultaView").fadeOut();
        //$("#iframeRastreioConsultaView").attr("src", "");
    }
    function sizeDivRastreioConsultaView() {
        var newHeight = $(document).height() + "px";
        $("#divRastreioConsultaView").css("height", newHeight);
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
.tdLoja{
	vertical-align: top;
	text-align:center;
	font-weight: bold;
	width: 30px;
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
.tdOcorrencia{
	vertical-align: top;
	width: 314px;
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
#divRastreioConsultaView
{
	position:absolute;
	top:0;
	left:0;
	width:100%;
	z-index:1000;
	background-color:#808080;
	opacity: 1;
}
#divInternoRastreioConsultaView
{
	position:absolute;
	top:6%;
	left:5%;
	width:90%;
	height:90%;
	z-index:1000;
	background-color:#fff;
	opacity: 1;
}
#divInternoRastreioConsultaView.divFixo
{
	position:fixed;
	top:6%;
}
#imgFechaDivRastreioConsultaView
{
	position:fixed;
	top:6%;
	left: 50%;
	margin-left: -16px; /* -1 * image width / 2 */
	margin-top: -32px;
	z-index:1001;
}
#iframeRastreioConsultaView
{
	position:absolute;
	top:0;
	left:0;
	width:100%;
	height:100%;
	border: solid 4px black;
}
</style>


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



<% else %>
<body onload="window.status='Concluído';focus();" link=#000000 alink=#000000 vlink=#000000>
<center>

<form id="fPED" name="fPED" method="post">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="pedido_selecionado" id="pedido_selecionado" value="">
</form>


<form id="fREL" name="fREL" method="post" action="RelPedidoOcorrenciaGravaDados.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="rb_status" id="rb_status" value="<%=rb_status%>">
<input type="hidden" name="c_loja" id="c_loja" value="<%=c_loja%>">


<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="1024" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Relatório de Ocorrências</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>

<!-- FILTROS -->
<% 
	s_filtro = "<table width='1024' cellPadding='0' CellSpacing='0' style='border-bottom:1px solid black' border='0'>" & chr(13)

	s = rb_status
	if s = "ABERTA" then
		s = "Aberta"
	elseif s = "EM_ANDAMENTO" then
		s = "Em Andamento"
	elseif s = "" then
		s = "Aberta, Em Andamento"
	else
		s = "Parâmetro Desconhecido"
		end if

	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' NOWRAP>" & _
				"<p class='N'>Status da Ocorrência:&nbsp;</p></td><td valign='top'>" & _
				"<p class='N'>" & s & "</p></td>" & chr(13) & _
				"	</tr>" & chr(13)

	s = c_loja
	if s = "" then s = "N.I."
	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' NOWRAP>" & _
				"<p class='N'>Loja:&nbsp;</p></td><td valign='top'>" & _
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
	<td align="right"><div name="dCONFIRMA" id="dCONFIRMA"><a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fRELConfirma(fREL)" title="executa a consulta">
		<img src="../botao/confirmar.gif" width="176" height="55" border="0"></a></div>
	</td>
</tr>
</table>
</form>

</center>

<div id="divRastreioConsultaView"><center><div id="divInternoRastreioConsultaView"><img id="imgFechaDivRastreioConsultaView" src="../imagem/close_button_32.png" title="clique para fechar o painel de consulta" /><iframe id="iframeRastreioConsultaView"></iframe></div></center></div>

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
