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
'	  RelPesquisaOrdemServicoExec.asp
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
	if Not operacao_permitida(OP_CEN_REL_PESQUISA_ORDEM_SERVICO, s_lista_operacoes_permitidas) then 
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if
	
	dim alerta
	dim s, s_aux
	dim c_dt_abertura_inicio, c_dt_abertura_termino
	dim c_dt_encerramento_inicio, c_dt_encerramento_termino
	dim c_fabricante, c_produto
	dim c_pedido
	dim c_vendedor, c_indicador
	dim s_nome_vendedor, s_nome_indicador
	dim s_nome_fabricante, s_nome_produto
	dim c_lista_loja, s_lista_loja, v_loja, v, i
	dim flag_ok

	alerta = ""

	c_dt_abertura_inicio=Trim(Request("c_dt_abertura_inicio"))
	c_dt_abertura_termino=Trim(Request("c_dt_abertura_termino"))
	c_dt_encerramento_inicio=Trim(Request("c_dt_encerramento_inicio"))
	c_dt_encerramento_termino=Trim(Request("c_dt_encerramento_termino"))
	c_fabricante=Trim(Request("c_fabricante"))
	c_produto=Trim(Request("c_produto"))
	c_pedido=Trim(Request("c_pedido"))
	c_vendedor = Ucase(Trim(Request("c_vendedor")))
	c_indicador = Ucase(Trim(Request("c_indicador")))
	c_lista_loja = Request("c_lista_loja")
	
	s_lista_loja = substitui_caracteres(c_lista_loja,chr(10),"")
	v_loja = split(s_lista_loja,chr(13),-1)
	
	dim FLAG_EXIBIR_DETALHES
	FLAG_EXIBIR_DETALHES = False

	dim s_filtro, s_filtro_loja, intQtdeOS
	intQtdeOS = 0
	
'	Fabricante
	if alerta = "" then
		s_nome_fabricante = ""
		if c_fabricante <> "" then
			s = "SELECT fabricante, nome, razao_social FROM t_FABRICANTE WHERE (fabricante='" & c_fabricante & "')"
			if rs.State <> 0 then rs.Close
			rs.open s, cn
			if rs.Eof then 
				alerta=texto_add_br(alerta)
				alerta = "FABRICANTE " & c_fabricante & " NÃO ESTÁ CADASTRADO."
			else
				s_nome_fabricante = Trim("" & rs("nome"))
				if s_nome_fabricante = "" then s_nome_fabricante = Trim("" & rs("razao_social"))
				end if
			end if
		end if
	
'	Produto
	if alerta = "" then
		s_nome_produto = ""
		if c_produto <> "" then
			if (Not IsEAN(c_produto)) And (c_fabricante="") then
				alerta=texto_add_br(alerta)
				alerta=alerta & "NÃO FOI ESPECIFICADO O FABRICANTE DO PRODUTO A SER CONSULTADO."
			else
				s = "SELECT * FROM t_PRODUTO WHERE"
				if IsEAN(c_produto) then
					s = s & " (ean='" & c_produto & "')"
				else
					s = s & " (fabricante='" & c_fabricante & "') AND (produto='" & c_produto & "')"
					end if
				if rs.State <> 0 then rs.Close
				rs.open s, cn
				if Not rs.Eof then
					flag_ok = True
					if IsEAN(c_produto) And (c_fabricante<>"") then
						if (c_fabricante<>Trim("" & rs("fabricante"))) then
							flag_ok = False
							alerta=texto_add_br(alerta)
							alerta=alerta & "Produto a ser consultado " & c_produto & " NÃO pertence ao fabricante " & c_fabricante & "."
							end if
						end if
					if flag_ok then
					'	CARREGA CÓDIGO INTERNO DO PRODUTO
						c_fabricante = Trim("" & rs("fabricante"))
						c_produto = Trim("" & rs("produto"))
						s_nome_produto = Trim("" & rs("descricao"))
						end if
					end if
				end if
			end if
		end if

'	Pedido
	if alerta = "" then
		if c_pedido <> "" then
			s = "SELECT pedido FROM t_PEDIDO WHERE (pedido='" & c_pedido & "')"
			if rs.State <> 0 then rs.Close
			rs.open s, cn
			if rs.Eof then 
				alerta=texto_add_br(alerta)
				alerta = "PEDIDO " & c_pedido & " NÃO ESTÁ CADASTRADO."
				end if
			end if
		end if

'	Vendedor
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

'	Indicador
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





' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________

function monta_link_pedido(byval id_pedido, byval texto_label)
dim strLink
	monta_link_pedido = ""
	id_pedido = Trim("" & id_pedido)
	texto_label = Trim("" & texto_label)
	if texto_label = "" then texto_label = id_pedido
	if id_pedido = "" then exit function
	strLink = "<a href='javascript:fPEDConsulta(" & _
				chr(34) & id_pedido & chr(34) & _
				")' title='clique para consultar o pedido " & id_pedido & "'>" & _
				texto_label & "</a>"
	monta_link_pedido=strLink
end function


function monta_link_OS(byval id_OS, byval texto_label)
dim strLink
	monta_link_OS = ""
	id_OS = Trim("" & id_OS)
	texto_label = Trim("" & texto_label)
	if texto_label = "" then texto_label = formata_num_OS_tela(id_OS)
	if id_OS = "" then exit function
	strLink = "<a href='javascript:fOSConsulta(" & _
				chr(34) & id_OS & chr(34) & _
				")' title='clique para consultar a ordem de serviço " & formata_num_OS_tela(id_OS) & "'>" & _
				texto_label & "</a>"
	monta_link_OS=strLink
end function


' _____________________________________
' CONSULTA EXECUTA
'
sub consulta_executa
dim s, s_plural, s_sql, s_where, s_where_loja, x
dim r
dim cab_table, cab
dim s_volume, s_tipo_volume, s_num_serie, s_obs_problema
dim intQtdeEmAndamento, intQtdeCancelado, intQtdeEncerrado

	s_sql = _
		"SELECT " & _
			"tOS.*, " & _
			"tP.vendedor, " & _
			"tP.loja, " & _
			"tC.cnpj_cpf AS cnpj_cpf_cliente, " &_
			"tU.nome_iniciais_em_maiusculas AS nome_vendedor, " & _
			"tOI.razao_social_nome_iniciais_em_maiusculas AS nome_indicador" & _
		" FROM t_ORDEM_SERVICO tOS" & _
			" LEFT JOIN t_PEDIDO tP ON (tOS.pedido=tP.pedido)" & _
			" LEFT JOIN t_CLIENTE tC ON (tP.id_cliente=tC.id)" & _
			" LEFT JOIN t_USUARIO tU ON (tP.vendedor=tU.usuario)" & _
			" LEFT JOIN t_ORCAMENTISTA_E_INDICADOR tOI ON (tOS.indicador=tOI.apelido)"
	
	s_where = ""
	
	if IsDate(c_dt_abertura_inicio) then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (tOS.data >= " & bd_formata_data(StrToDate(c_dt_abertura_inicio)) & ")"
		end if
	
	if IsDate(c_dt_abertura_termino) then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (tOS.data < " & bd_formata_data(StrToDate(c_dt_abertura_termino)+1) & ")"
		end if
	
	if IsDate(c_dt_encerramento_inicio) then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " ((tOS.situacao_data >= " & bd_formata_data(StrToDate(c_dt_encerramento_inicio)) & ") AND ((tOS.situacao_status = '" & ST_OS_ENCERRADA & "') OR (tOS.situacao_status = '" & ST_OS_CANCELADA & "')))"
		end if
	
	if IsDate(c_dt_encerramento_termino) then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " ((tOS.situacao_data < " & bd_formata_data(StrToDate(c_dt_encerramento_termino)+1) & ") AND ((tOS.situacao_status = '" & ST_OS_ENCERRADA & "') OR (tOS.situacao_status = '" & ST_OS_CANCELADA & "')))"
		end if
	
	if c_fabricante <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (tOS.fabricante = '" & c_fabricante & "')"
		end if
	
	if c_produto <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (tOS.produto = '" & c_produto & "')"
		end if
	
	if c_pedido <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (tOS.pedido = '" & c_pedido & "')"
		end if
	
	if c_vendedor <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (tP.vendedor = '" & c_vendedor & "')"
		end if
	
	if c_indicador <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (tOS.indicador = '" & c_indicador & "')"
		end if
	
'	LOJAS
	s_where_loja = ""
	for i=Lbound(v_loja) to Ubound(v_loja)
		if v_loja(i) <> "" then
			v = split(v_loja(i),"-",-1)
			if Ubound(v)=Lbound(v) then
				if s_where_loja <> "" then s_where_loja = s_where_loja & " OR"
				s_where_loja = s_where_loja & " (CONVERT(smallint, tP.loja) = " & v_loja(i) & ")"
			else
				s = ""
				if v(Lbound(v))<>"" then 
					if s <> "" then s = s & " AND"
					s = s & " (CONVERT(smallint, tP.loja) >= " & v(Lbound(v)) & ")"
					end if
				if v(Ubound(v))<>"" then
					if s <> "" then s = s & " AND"
					s = s & " (CONVERT(smallint, tP.loja) <= " & v(Ubound(v)) & ")"
					end if
				if s <> "" then 
					if s_where_loja <> "" then s_where_loja = s_where_loja & " OR"
					s_where_loja = s_where_loja & " (" & s & ")"
					end if
				end if
			end if
		next
		
	if s_where_loja <> "" then 
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (" & s_where_loja & ")"
		end if
	
	if s_where <> "" then s_sql = s_sql & " WHERE" & s_where

	s_sql = "SELECT * FROM (" & s_sql & ") t ORDER BY data, hora, ordem_servico"

	cab_table = "<TABLE cellSpacing=0 cellPadding=0>" & chr(13)
	cab = "	<TR style='background:azure' NOWRAP>" & chr(13) & _
		  "		<TD valign='bottom' style='background:white;' NOWRAP>&nbsp;</TD>" & chr(13) & _
		  "		<TD class='MDTE tdDataHora' style='vertical-align:bottom'><P class='Rc'>DT Abert</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdDataHora' style='vertical-align:bottom'><P class='Rc'>DT Encerr</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdPedido' style='vertical-align:bottom'><P class='R'>Pedido</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdOS' style='vertical-align:bottom'><P class='Rc'>O.S.</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdProd' style='vertical-align:bottom'><P class='R'>Produto</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdCliente' style='vertical-align:bottom'><P class='R'>Cliente</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdVendedor' style='vertical-align:bottom'><P class='R'>Vendedor</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdParceiro' style='vertical-align:bottom'><P class='R'>Parceiro</P></TD>" & chr(13) & _
		  "		<TD style='background:white;'>&nbsp;</TD>" & chr(13) & _
		  "	</TR>" & chr(13)
	
	intQtdeOS = 0
	intQtdeEmAndamento = 0
	intQtdeCancelado = 0
	intQtdeEncerrado = 0
	
	x = cab_table & cab
	
	set r = cn.execute(s_sql)
	do while Not r.Eof
	
	'	CONTAGEM
		intQtdeOS = intQtdeOS + 1

		if Trim("" & r("situacao_status")) = ST_OS_EM_ANDAMENTO then
			intQtdeEmAndamento = intQtdeEmAndamento + 1
		elseif Trim("" & r("situacao_status")) = ST_OS_ENCERRADA then
			intQtdeEncerrado = intQtdeEncerrado + 1
		elseif Trim("" & r("situacao_status")) = ST_OS_CANCELADA then
			intQtdeCancelado = intQtdeCancelado + 1
			end if
		
		x = x & "	<TR NOWRAP>" & chr(13)

	'> Nº DA LINHA
		x = x & "		<TD valign='top' align='right' NOWRAP><P class='Rd' style='margin-right:2px;'>" & Cstr(intQtdeOS) & ".</P></TD>" & chr(13)

	'> DATA ABERTURA
		s = formata_data(r("data")) & " " & formata_hhnnss_para_hh_nn(Trim("" & r("hora")))
		s = monta_link_OS(Trim("" & r("ordem_servico")), s)
		x = x & "		<TD class='MDTE tdDataHora'><P class='Cnc'>" & s & "</P></TD>" & chr(13)

	'> DATA ENCERRAMENTO
		if (Trim("" & r("situacao_status")) = ST_OS_ENCERRADA) Or (Trim("" & r("situacao_status")) = ST_OS_CANCELADA) then
			s = formata_data_hora_sem_seg(r("situacao_data"))
		else
			s = "&nbsp;"
			end if
		s = monta_link_OS(Trim("" & r("ordem_servico")), s)
		x = x & "		<TD class='MTD tdDataHora'><P class='Cnc'>" & s & "</P></TD>" & chr(13)

	'> PEDIDO
		s = monta_link_pedido(Trim("" & r("pedido")), Trim("" & r("pedido")))
		x = x & "		<TD class='MTD tdPedido'><P class='Cn'>" & s & "</P></TD>" & chr(13)

	'> NÚMERO DA OS
		s = monta_link_OS(Trim("" & r("ordem_servico")), formata_num_OS_tela(Trim("" & r("ordem_servico"))))
		x = x & "		<TD class='MTD tdOS'><P class='Cnc'>" & s & "</P></TD>" & chr(13)

	'> CÓDIGO E DESCRIÇÃO DO PRODUTO
		s = "(" & Trim("" & r("fabricante")) & ") " & Trim("" & r("produto"))
		if Trim("" & r("descricao_html")) <> "" then
			s = s & " - " & produto_formata_descricao_em_html(Trim("" & r("descricao_html")))
		else
			s = s & " - " & Trim("" & r("descricao"))
			end if
		if s = "" then s = "&nbsp;"
		x = x & "		<TD class='MTD tdProd'><P class='Cn'>" & s & "</P></TD>" & chr(13)

	'> CLIENTE
		s = iniciais_em_maiusculas(Trim("" & r("nome_cliente"))) & " (" & cnpj_cpf_formata(Trim("" & r("cnpj_cpf_cliente"))) & ")"
		if s = "" then s = "&nbsp;"
		x = x & "		<TD class='MTD tdCliente'><P class='Cn'>" & s & "</P></TD>" & chr(13)

	'> VENDEDOR
		s = Trim("" & r("vendedor"))
		if Ucase(s) <> Ucase(Trim("" & r("nome_vendedor"))) then s = s & " - " & Trim("" & r("nome_vendedor"))
		if s = "" then s = "&nbsp;"
		x = x & "		<TD class='MTD tdVendedor'><P class='Cn'>" & s & "</P></TD>" & chr(13)

	'> PARCEIRO
		s = Trim("" & r("indicador"))
		if Ucase(s) <> Ucase(Trim("" & r("nome_indicador"))) then s = s & " - " & Trim("" & r("nome_indicador"))
		if s = "" then s = "&nbsp;"
		x = x & "		<TD class='MTD tdParceiro'><P class='Cn'>" & s & "</P></TD>" & chr(13)

	'> BOTÃO P/ EXIBIR DEMAIS CAMPOS
		if FLAG_EXIBIR_DETALHES then
			x = x & "		<TD valign='bottom' class='notPrint'>" & _
								"&nbsp;<a name='bExibeOcultaCampos' id='bExibeOcultaCampos' href='javascript:fExibeOcultaCampos(" & chr(34) & Cstr(intQtdeOS) & chr(34) & ")' title='exibe ou oculta os campos adicionais'><img src='../botao/view_bottom.png' border='0'></a>" & _
							"</TD>" & chr(13)
		else
			x = x & "		<TD valign='bottom' class='notPrint'>&nbsp;" & "</TD>" & chr(13)
			end if

		x = x & "	</TR>" & chr(13)

		if FLAG_EXIBIR_DETALHES then
		'> VOLUMES
			s_sql = _
				"SELECT " & _
					"*" & _
				" FROM t_ORDEM_SERVICO_ITEM" & _
				" WHERE" & _
					" (ordem_servico = '" & Trim("" & r("ordem_servico")) & "')" & _
					" AND (excluido_status = 0)" & _
				" ORDER BY" & _
					" sequencia"
			if rs.State <> 0 then rs.Close
			rs.open s_sql, cn
			x = x & "	<TR style='display:none;' id='TR_MORE_INFO_" & Cstr(intQtdeOS) & "'>" & chr(13) & _
					"		<TD style='background:white;'>&nbsp;</td>" & chr(13) & _
					"		<TD class='ME MD'>&nbsp;</TD>" & chr(13) & _
					"		<TD class='MD'>&nbsp;</TD>" & chr(13) & _
					"		<TD colspan='6' class='MC MD'>" & chr(13) & _
					"			<table width='100%' cellspacing='0' cellpadding='0'>" & chr(13) & _
					"				<TR style='background:whitesmoke;'>" & chr(13) & _
					"					<td class='Rf tdWithPadding' align='center'>VOLUMES</td>" & chr(13) & _
					"				</TR>" & chr(13)
			if rs.Eof then
				x = x & _
					"				<TR>" & chr(13) & _
					"					<td>&nbsp;</td>" & chr(13) & _
					"				</TR>" & chr(13)
				end if
		
			if Not rs.Eof then
				x = x & _
					"				<TR>" & chr(13) & _
					"					<TD>" & chr(13) & _
					"						<table width='100%' cellSpacing='0' cellPadding='0'>" & chr(13) & _
					"							<TR>" & chr(13) & _
					"								<TD class='Rf MD MC tdWithPadding tdVolume' align='center'>" & chr(13) & _
														"Volume" & _
					"								</TD>" & chr(13) & _
					"								<TD class='Rf MD MC tdWithPadding tdTipoVolume' align='center'>" & chr(13) & _
														"Tipo" & _
					"								</TD>" & chr(13) & _
					"								<TD class='Rf MD MC tdWithPadding tdNumSerie' align='center' valign='top'>" & chr(13) & _
														"Nº Série" & _
													"</TD>" & chr(13) & _
					"								<TD class='Rf MC tdWithPadding tdProblema' align='left' valign='top'>" & chr(13) & _
														"Problema" & _
													"</TD>" & chr(13) & _
					"							</TR>" & chr(13) & _
					"						</table>" & chr(13) & _
					"					</TD>" & chr(13) & _
					"				</TR>" & chr(13)
				end if
		
			do while Not rs.Eof
				s_volume = TrimRightCrLf(Trim("" & rs("descricao_volume")))
				if s_volume = "" then s_volume = "&nbsp;"
				s_tipo_volume = TrimRightCrLf(Trim("" & rs("tipo")))
				if s_tipo_volume = "" then s_tipo_volume = "&nbsp;"
				s_num_serie = TrimRightCrLf(Trim("" & rs("num_serie")))
				if s_num_serie = "" then s_num_serie = "&nbsp;"
				s_obs_problema = substitui_caracteres(TrimRightCrLf(Trim("" & rs("obs_problema"))), chr(13), "<br>")
				if s_obs_problema = "" then s_obs_problema = "&nbsp;"
				x = x & _
					"				<TR>" & chr(13) & _
					"					<TD>" & chr(13) & _
					"						<table width='100%' cellSpacing='0' cellPadding='0'>" & chr(13) & _
					"							<TR>" & chr(13) & _
					"								<TD class='Cn MD MC tdWithPadding tdVolume' align='center'>" & chr(13) & _
														s_volume & _
					"								</TD>" & chr(13) & _
					"								<TD class='Cn MD MC tdWithPadding tdTipoVolume' align='center'>" & chr(13) & _
														s_tipo_volume & _
					"								</TD>" & chr(13) & _
					"								<TD class='Cn MD MC tdWithPadding tdNumSerie' align='center' valign='top'>" & chr(13) & _
														s_num_serie & _
													"</TD>" & chr(13) & _
					"								<TD class='Cn MC tdWithPadding tdProblema' align='left' valign='top'>" & chr(13) & _
														s_obs_problema & _
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
			end if 'if FLAG_EXIBIR_DETALHES
		
		if (intQtdeOS mod 100) = 0 then
			Response.Write x
			x = ""
			end if
			
		r.MoveNext
		loop
		
		
'	TOTAL GERAL
	if intQtdeOS > 0 then
		s_plural = "TOTAL: &nbsp; " & Cstr(intQtdeOS)
		if intQtdeOS = 1 then
			s_plural = s_plural & " ordem de serviço"
		else
			s_plural = s_plural & " ordens de serviço"
			end if
		s_plural = s_plural & " &nbsp; ("
	'	EM ANDAMENTO
		s_plural = s_plural & Cstr(intQtdeEmAndamento) & " em andamento"
	'	ENCERRADO
		s_plural = s_plural & "; &nbsp; " & Cstr(intQtdeEncerrado)
		if intQtdeEncerrado = 1 then
			s_plural = s_plural & " encerrado"
		else
			s_plural = s_plural & " encerrados"
			end if
	'	CANCELADO
		s_plural = s_plural & "; &nbsp; " & Cstr(intQtdeCancelado)
		if intQtdeCancelado = 1 then
			s_plural = s_plural & " cancelado"
		else
			s_plural = s_plural & " cancelados"
			end if
		s_plural = s_plural & ")"
		
		x = x & "	<TR>" & chr(13) & _
				"		<TD style='background:white;'>&nbsp;</td>" & chr(13) & _
				"		<TD COLSPAN='8' class='MC' style='border-left:0px;border-right:0px;'>&nbsp;</TD>" & chr(13) & _
				"	</TR>" & chr(13) & _
				"	<TR NOWRAP style='background:honeydew'>" & chr(13) & _
				"		<TD style='background:white;'>&nbsp;</td>" & chr(13) & _
				"		<TD COLSPAN='8' class='MT'><p class='C'>" & s_plural & "</p></TD>" & chr(13) & _
				"	</TR>" & chr(13)
		end if
	
'	MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	if intQtdeOS = 0 then
		x = cab_table & cab
		x = x & "	<TR NOWRAP>" & chr(13) & _
				"		<TD style='background:white;'>&nbsp;</td>" & chr(13) & _
				"		<TD class='MT' colspan='8'><P class='ALERTA'>&nbsp;NENHUM REGISTRO ENCONTRADO&nbsp;</P></TD>" & chr(13) & _
				"	</TR>" & chr(13)
		end if

'	FECHA TABELA
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

function expandirTudo() {
var i;
var row_MORE_INFO;
	for (i = 1; i <= intQtdeRows; i++) {
		row_MORE_INFO = document.getElementById("TR_MORE_INFO_" + i);
		row_MORE_INFO.style.display = "";
	}
}

function recolherTudo() {
var i;
var row_MORE_INFO;
	for (i = 1; i <= intQtdeRows; i++) {
		row_MORE_INFO = document.getElementById("TR_MORE_INFO_" + i);
		row_MORE_INFO.style.display = "none";
	}
}

function fExibeOcultaCampos(indice_row) {
var row_MORE_INFO;

	row_MORE_INFO = document.getElementById("TR_MORE_INFO_" + indice_row);
	if (row_MORE_INFO.style.display.toString() == "none") {
		row_MORE_INFO.style.display = "";
	}
	else {
		row_MORE_INFO.style.display = "none";
	}
}

function fPEDConsulta(id_pedido) {
	window.status = "Aguarde ...";
	fPED.pedido_selecionado.value = id_pedido;
	fPED.action = "pedido.asp"
	fPED.submit();
}

function fOSConsulta(id_OS) {
	window.status = "Aguarde ...";
	fOS.num_OS.value = id_OS;
	fOS.action = "OrdemServico.asp"
	fOS.submit();
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
.tdOS{
	vertical-align: top;
	font-weight: bold;
	width: 60px;
	}
.tdProd{
	vertical-align: top;
	width: 160px;
	}
.tdCliente{
	vertical-align: top;
	width: 160px;
	}
.tdVendedor{
	vertical-align: top;
	width: 120px;
	}
.tdParceiro{
	vertical-align: top;
	width: 140px;
	}
.tdVolume{
	vertical-align: top;
	width: 93px;
	}
.tdTipoVolume{
	vertical-align: top;
	width: 93px;
	}
.tdNumSerie{
	vertical-align: top;
	width: 93px;
	}
.tdProblema{
	vertical-align: top;
	width: 419px;
	}
</style>


<body onload="window.status='Concluído';focus();" link=#000000 alink=#000000 vlink=#000000>
<center>

<form id="fPED" name="fPED" method="post">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="pedido_selecionado" id="pedido_selecionado" value="">
</form>

<form id="fOS" name="fOS" method="post">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="num_OS" id="num_OS" value="">
</form>



<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="900" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Pesquisa de Ordem de Serviço</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>

<!-- FILTROS -->
<% 
	s_filtro = "<table width='900' cellPadding='0' CellSpacing='0' style='border-bottom:1px solid black' border='0'>" & chr(13)

'	PERÍODO DE ABERTURA
	s = ""
	s_aux = c_dt_abertura_inicio
	if s_aux = "" then s_aux = "N.I."
	s = s & s_aux & " a "
	s_aux = c_dt_abertura_termino
	if s_aux = "" then s_aux = "N.I."
	s = s & s_aux
	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' NOWRAP><p class='N'>Período de Abertura:&nbsp;</p></td>" & chr(13) & _
				"		<td valign='top' width='99%'><p class='N'>" & s & "</p></td>" & chr(13) & _
				"	</tr>" & chr(13)

'	PERÍODO DE ENCERRAMENTO
	s = ""
	s_aux = c_dt_encerramento_inicio
	if s_aux = "" then s_aux = "N.I."
	s = s & s_aux & " a "
	s_aux = c_dt_encerramento_termino
	if s_aux = "" then s_aux = "N.I."
	s = s & s_aux
	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' NOWRAP><p class='N'>Período de Encerramento:&nbsp;</p></td>" & chr(13) & _
				"		<td valign='top' width='99%'><p class='N'>" & s & "</p></td>" & chr(13) & _
				"	</tr>" & chr(13)

'	FABRICANTE
	s = c_fabricante
	if s = "" then
		s = "N.I."
	else
		if s_nome_fabricante <> "" then s = c_fabricante & " - " & s_nome_fabricante
		end if
	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' NOWRAP><p class='N'>Fabricante:&nbsp;</p></td>" & chr(13) & _
				"		<td valign='top'><p class='N'>" & s & "</p></td>" & chr(13) & _
				"	</tr>" & chr(13)

'	PRODUTO
	s = c_produto
	if s = "" then 
		s = "N.I."
	else
		if s_nome_produto <> "" then s = c_produto & " - " & s_nome_produto
		end if
	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' NOWRAP><p class='N'>Produto:&nbsp;</p></td>" & chr(13) & _
				"		<td valign='top'><p class='N'>" & s & "</p></td>" & chr(13) & _
				"	</tr>" & chr(13)

'	PEDIDO
	s = c_pedido
	if s = "" then s = "N.I."
	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' NOWRAP><p class='N'>Pedido:&nbsp;</p></td>" & chr(13) & _
				"		<td valign='top'><p class='N'>" & s & "</p></td>" & chr(13) & _
				"	</tr>" & chr(13)
	
'	VENDEDOR
	s = c_vendedor
	if s = "" then
		s = "N.I."
	else
		if (s_nome_vendedor <> "") And (s_nome_vendedor <> c_vendedor) then s = s & "  (" & s_nome_vendedor & ")"
		end if
	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' NOWRAP><p class='N'>Vendedor:&nbsp;</p></td>" & chr(13) & _
				"		<td valign='top'><p class='N'>" & s & "</p></td>" & chr(13) & _
				"	</tr>" & chr(13)

'	INDICADOR
	s = c_indicador
	if s = "" then 
		s = "N.I."
	else
		if (s_nome_indicador <> "") And (s_nome_indicador <> c_indicador) then s = s & "  (" & s_nome_indicador & ")"
		end if
	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' NOWRAP><p class='N'>Indicador:&nbsp;</p></td>" & chr(13) & _
				"		<td valign='top'><p class='N'>" & s & "</p></td>" & chr(13) & _
				"	</tr>" & chr(13)

'	LISTA DE LOJAS
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
	if s = "" then s = "N.I."
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

<% consulta_executa %>

<script language="JavaScript" type="text/javascript">
var intQtdeRows=<%=Cstr(intQtdeOS)%>;
</script>

<!-- ************   SEPARADOR   ************ -->
<table width="900" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>

<table class="notPrint" width='853' cellPadding='0' CellSpacing='0' border='0' style="margin-top:5px;">
<tr>
	<td width="50%" nowrap>&nbsp;</td>
	<td align="right" nowrap><a id="bImprimir" href="javascript:window.print();"><p class="Button" style="margin-bottom:0px;">Imprimir...</p></a></td>
	<td nowrap>&nbsp;</td>
	<td align="right" nowrap><a id="linkExpandirTudo" href="javascript:expandirTudo();"><p class="Button" style="margin-bottom:0px;">Expandir Tudo</p></a></td>
	<td nowrap>&nbsp;</td>
	<td align="right" nowrap><a id="linkRecolherTudo" href="javascript:recolherTudo();"><p class="Button" style="margin-bottom:0px;">Recolher Tudo</p></a></td>
</tr>
</table>

<table class="notPrint" width="900" cellSpacing="0">
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
