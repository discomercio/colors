<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ======================================================
'	  RelSeparacaoZonaConsultaDetalhe.asp
'     ======================================================
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

	Server.ScriptTimeout = MAX_SERVER_SCRIPT_TIMEOUT_EM_SEG
	
	dim usuario
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	CONECTA COM O BANCO DE DADOS
	dim cn, tN1
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	if Not operacao_permitida(OP_CEN_REL_SEPARACAO_ZONA, s_lista_operacoes_permitidas) then 
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	dim blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos
	blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos = isActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos

	dim alerta
	alerta = ""

	dim s, s_aux, s_table_produtos_sem_zona, s_filtro
	dim nsu_selecionado
	dim c_usuario_emissao, c_dt_hr_emissao, c_dt_inicio, c_dt_termino, rb_nfe, c_transportadora, c_qtde_max_pedidos, lista_zonas_cadastradas
	c_usuario_emissao = ""
	c_dt_hr_emissao = ""
	c_dt_inicio = ""
	c_dt_termino = ""
	rb_nfe = ""
	c_transportadora = ""
	c_qtde_max_pedidos = ""
	lista_zonas_cadastradas = ""

	nsu_selecionado = Trim(Request.Form("nsu_selecionado"))
	
	if nsu_selecionado = "" then
		alerta = "Não foi informado o NSU do relatório!!"
		end if
	
	if alerta = "" then
		s = "SELECT " & _
				"*" & _
			" FROM t_WMS_ETQ_N1_SEPARACAO_ZONA_RELATORIO" & _
			" WHERE" & _
				" (id = " & nsu_selecionado & ")"
		set tN1 = cn.Execute(s)
		if tN1.Eof then
			alerta = "Relatório de Separação (Zona) com NSU=" & nsu_selecionado & " não foi encontrado!!"
			end if
		end if
	
	if alerta = "" then
		c_usuario_emissao = Trim("" & tN1("usuario"))
		c_dt_hr_emissao = formata_data_hora_sem_seg(tN1("dt_hr_emissao"))
		c_dt_inicio = Trim("" & tN1("filtro_dt_inicio"))
		c_dt_termino = Trim("" & tN1("filtro_dt_termino"))
		rb_nfe = Trim("" & tN1("filtro_NFe_emitida"))
		c_transportadora = Trim("" & tN1("filtro_transportadora"))
		c_qtde_max_pedidos = Trim("" & tN1("filtro_qtde_max_pedidos"))
		lista_zonas_cadastradas = Trim("" & tN1("lista_zonas_cadastradas"))
		end if
	
	dim strScriptJS
	strScriptJS = ""





' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________

' _____________________________________
' CONSULTA RELATORIO
' Subsídios: o depósito está dividido em zonas, cada qual com uma equipe de separadores.
' ========== Cada equipe de separadores realiza o "picking" de produtos e os leva até uma
' área de consolidação. A equipe de consolidação irá reunir os itens que compõem o pedido
' e levá-los até o box da transportadora adequada.
' Inicialmente, o filtro deste relatório possuía um campo destinado a restringir por zona,
' para que o relatório fosse executado uma vez p/ cada uma das zonas e mais uma consulta
' abrangendo todas as zonas p/ o consolidador.
' O problema dessa abordagem é que os pedidos podem ser alterados de modo que passem a
' constar no relatório durante esse processo de impressão, causando inconsistência entre os
' relatórios de cada uma das zonas com o relatório do consolidador.
' A solução implementada é realizar uma única consulta ao banco de dados e, a partir dela,
' gerar todos os relatórios de uma vez.
sub consulta_relatorio
const COD_ZONA_ID__TODOS = 999999
const COD_ZONA_CODIGO__TODOS = "***TODOS***"
dim r
dim s, s_aux, s_zona, s_sql, pedido_a, fabricante_a, s_numero_NF, s_transportadora
dim i, iZona, iRel, n_qtde_zona, n_reg_total, idx
dim vProd(), vZona(), vRel()
dim v, vz
dim x, xRel
dim s_linha_branco, s_tit_zona
dim blnProcessaRegistro
dim strJS_AllTablesCollapse, strJS_AllTablesNotPrint

'	VETOR QUE ARMAZENA TODOS OS REGISTROS (A ORDENAÇÃO DEVE SER FEITA NA CONSULTA SQL)
	redim vRel(0)
	set vRel(0) = New cl_REL_SEPARACAO_ZONA
	vRel(0).produto = ""

'	ARMAZENA A LISTA DE ZONAS CADASTRADAS
	redim vZona(0)
	set vZona(0) = New cl_ZONA_DEPOSITO
	vZona(0).zona_id = 0
	
	vz = Split(lista_zonas_cadastradas, "|")
	for i=LBound(vz) to UBound(vz)
		if Trim("" & vz(i)) <> "" then
			if vZona(Ubound(vZona)).zona_id <> 0 then
				redim preserve vZona(Ubound(vZona)+1)
				set vZona(Ubound(vZona)) = New cl_ZONA_DEPOSITO
				end if
			v=Split(vz(i), "§")
			with vZona(Ubound(vZona))
				.zona_id = CLng(v(0))
				.zona_codigo = Trim("" & v(1))
				end with
			end if
		next
	
	s_sql = "SELECT" & _
				" tN2.pedido," & _
				" tN2.obs_2," & _
				" tN2.obs_3," & _
				" tN2.loja," & _
				" tN2.transportadora_id," & _
				" tN2.id_cliente,"

	if blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos then
		s_sql = s_sql & _
				" tPed.endereco_nome_iniciais_em_maiusculas AS nome_cliente,"
	else
		s_sql = s_sql & _
				" tCli.nome_iniciais_em_maiusculas AS nome_cliente,"
		end if

	s_sql = s_sql & _
				" tFab.nome AS nome_fabricante," & _
				" tFab.razao_social AS razao_social_fabricante," & _
				" tN3.fabricante," & _
				" tN3.produto," & _
				" tN3.qtde," & _
				" tN3.qtde_volumes," & _
				" tPedItem.descricao," & _
				" tPedItem.descricao_html," & _
				" tN3.zona_id," & _
				" tN3.zona_codigo," & _
				" tN2.numeroNFe," & _
				" tN2.qtde_volumes_pedido" & _
			" FROM t_WMS_ETQ_N1_SEPARACAO_ZONA_RELATORIO tN1" & _
				" INNER JOIN t_WMS_ETQ_N2_SEPARACAO_ZONA_PEDIDO tN2 ON (tN1.id = tN2.id_wms_etq_n1)" & _
				" INNER JOIN t_WMS_ETQ_N3_SEPARACAO_ZONA_PRODUTO tN3 ON (tN2.id = tN3.id_wms_etq_n2)" & _
				" INNER JOIN t_CLIENTE tCli ON (tN2.id_cliente = tCli.id)" & _
				" INNER JOIN t_PEDIDO tPed ON (tN2.pedido = tPed.pedido)" & _
				" INNER JOIN t_PEDIDO_ITEM tPedItem ON (tPed.pedido = tPedItem.pedido) AND (tN3.fabricante = tPedItem.fabricante) AND (tN3.produto = tPedItem.produto)" & _
				" INNER JOIN t_FABRICANTE tFab ON (tN3.fabricante = tFab.fabricante)" & _
				" INNER JOIN t_PRODUTO tProd ON ((tN3.fabricante = tProd.fabricante) AND (tN3.produto = tProd.produto))" & _
			" WHERE" & _
				" (tN1.id = " & nsu_selecionado & ")" & _
			" ORDER BY" & _
				" tN1.id," & _
				" tN2.id," & _
				" tN3.id"
	
	set r = cn.execute(s_sql)

'	ARMAZENA TODOS OS REGISTROS NO VETOR
	do while Not r.Eof
		if Trim("" & vRel(Ubound(vRel)).produto) <> "" then
			redim preserve vRel(Ubound(vRel)+1)
			set vRel(Ubound(vRel)) = New cl_REL_SEPARACAO_ZONA
			end if
		
		with vRel(Ubound(vRel))
			.pedido = Trim("" & r("pedido"))
			.obs_2 = Trim("" & r("obs_2"))
			.obs_3 = Trim("" & r("obs_3"))
			.loja = Trim("" & r("loja"))
			.id_cliente = Trim("" & r("id_cliente"))
			.nome_cliente = Trim("" & r("nome_cliente"))
			.transportadora_id = Trim("" & r("transportadora_id"))
			.fabricante = Trim("" & r("fabricante"))
			.produto = Trim("" & r("produto"))
			.qtde = r("qtde")
			.qtde_volumes = r("qtde_volumes")
			.descricao = Trim("" & r("descricao"))
			.descricao_html = Trim("" & r("descricao_html"))
			.zona_id = r("zona_id")
			.zona_codigo = Trim("" & r("zona_codigo"))
			.numeroNFe = Trim("" & r("numeroNFe"))
			.qtde_volumes_pedido = r("qtde_volumes_pedido")
			.nome_fabricante = Trim("" & r("nome_fabricante"))
			if .nome_fabricante = "" then .nome_fabricante = Trim("" & r("razao_social_fabricante"))
			end with
		
		r.MoveNext
		loop

	if r.State <> 0 then r.Close
	set r=nothing

	n_qtde_zona = 0
	xRel = ""
	x = ""
	strJS_AllTablesCollapse = ""
	strJS_AllTablesNotPrint = ""

'	GERA O RELATÓRIO PARA CADA UMA DAS ZONAS
	for iZona=Lbound(vZona) to Ubound(vZona)
		if converte_numero(vZona(iZona).zona_id) <> 0 then
		
		'	AS ROTINAS DE ORDENAÇÃO USAM VETORES QUE SE INICIAM NA POSIÇÃO 1
			redim vProd(1)
			for i = Lbound(vProd) to Ubound(vProd)
				set vProd(i) = New cl_CINCO_COLUNAS
				with vProd(i)
					.c1 = ""
					.c2 = 0
					.c3 = ""
					.c4 = ""
					.c5 = ""
					end with
				next
			
			n_reg_total = 0
			pedido_a = "XXXXXXXXXX"
			n_qtde_zona = n_qtde_zona + 1
			
			for iRel=Lbound(vRel) to Ubound(vRel)
				blnProcessaRegistro = False
				if Trim("" & vRel(iRel).produto) <> "" then
					if Trim("" & vRel(iRel).zona_id) = Trim("" & vZona(iZona).zona_id) then blnProcessaRegistro = True
					if vZona(iZona).zona_id = COD_ZONA_ID__TODOS then blnProcessaRegistro = True
					end if
				
				if blnProcessaRegistro then
					with vRel(iRel)
						if Trim("" & .pedido) <> pedido_a then
							pedido_a = Trim("" & .pedido)
						'	FECHA TABELA DO PEDIDO ANTERIOR
							if n_reg_total > 0 then
								x = x & "		</table>" & chr(13) & _
										"	</td></tr>" & chr(13) & _
										"</table>" & chr(13) & _
										"<br>" & chr(13)
								end if

						'	Nº DA NF
							s_numero_NF = Trim("" & .obs_2)
							if Trim("" & .obs_3) <> "" then s_numero_NF = Trim("" & .obs_3)
							if s_numero_NF = "" then s_numero_NF = "&nbsp;"
							
						'	TRANSPORTADORA
							s_transportadora = iniciais_em_maiusculas(Trim("" & .transportadora_id))
							if s_transportadora = "" then s_transportadora = "&nbsp;"
							
						'	TABELA P/ O PRÓXIMO PEDIDO
							x = x & chr(13) & _
								"<table cellspacing='0' cellpadding='0'>" & chr(13) & _
								"	<tr><td align='left'>" & chr(13) & _
								"		<table class='Q' cellspacing=0 cellpadding=0>" & chr(13) & _
								"			<tr style='background:#FFF0E0' nowrap>" & chr(13)& _
								"				<td class='tdPedido' align='left' nowrap><span class='C'><a href='javascript:fRELConcluir(" & _
												chr(34) & Trim("" & .pedido) & chr(34) & ")' title='clique para consultar o pedido'>" & _
													Trim("" & .pedido) & "</a></span></td>" & chr(13) & _
								"				<td class='tdQtVol' align='left' nowrap><span class='C'>" & _
													Trim("" & .qtde_volumes_pedido) & "</span></td>" & chr(13) & _
								"				<td class='tdNF' align='left' nowrap><span class='C'>" & _
													s_numero_NF & "</span></td>" & chr(13) & _
								"				<td class='tdLoja' align='left' nowrap><span class='C'>Lj&nbsp;" & _
													Trim("" & .loja) & "</span></td>" & chr(13) & _
								"				<td class='tdCliente' align='left'><span class='C'>" & _
													Trim("" & .nome_cliente) & "</span></td>" & chr(13) & _
								"				<td class='tdTransp' align='left' nowrap><span class='C'>" & _
													s_transportadora & "</span></td>" & chr(13) & _
								"			</tr>" & chr(13) & _
								"		</table>" & chr(13) & _
								"	</td></tr>" & chr(13) & _
								"	<tr><td align='left'>" & chr(13) & _
								"		<table width='100%' cellspacing='0' cellpadding='0'>" & chr(13)
							end if

					  ' CONTAGEM
						n_reg_total = n_reg_total + 1
						
					'	LISTAGEM
						x = x & "			<tr nowrap>" & chr(13)

					 '> QTDE
						x = x & "				<td class='MDBE tdQtde' align='left'><span class='Cd'>&nbsp;" & _
							formata_inteiro(.qtde) & "</span></td>"  & chr(13)

 					 '> PRODUTO
						x = x & "				<td class='MDB tdProd' align='left' nowrap><span class='C' nowrap>&nbsp;" & _
							produto_formata_descricao_em_html(Trim("" & .descricao_html)) & _
							"&nbsp;&nbsp;(Cód:&nbsp;" & Trim("" & .produto) & _
							"&nbsp;&nbsp;Fabr:&nbsp;" & Trim("" & .fabricante) & ")" & _
							"</span></td>" & chr(13)

 					 '> ZONA
 						s = Trim("" & .zona_codigo)
 						if s = "" then s = "&nbsp;"
 						x = x & "				<td class='MDB tdZona' align='left' nowrap><span class='Cc' nowrap>" & s & "</span></td>" & chr(13)

						x = x & "			</tr>" & chr(13)

					'	TOTALIZAÇÃO
						s = Trim("" & .fabricante) & "|" & Trim("" & .produto)
						if localiza_cl_cinco_colunas(vProd, s, idx) then
							with vProd(idx)
								.c2 = .c2 + CLng(vRel(iRel).qtde)
								end with
						else
							if (vProd(Ubound(vProd)).c1<>"") then
								redim preserve vProd(Ubound(vProd)+1)
								set vProd(Ubound(vProd)) = New cl_CINCO_COLUNAS
								end if
							with vProd(Ubound(vProd))
								.c1 = Trim("" & vRel(iRel).fabricante) & "|" & Trim("" & vRel(iRel).produto)
								.c2 = CLng(vRel(iRel).qtde)
								.c3 = Trim("" & vRel(iRel).zona_codigo)
								.c4 = Trim("" & vRel(iRel).descricao_html)
								.c5 = Trim("" & vRel(iRel).nome_fabricante)
								end with
							ordena_cl_cinco_colunas vProd, 1, Ubound(vProd)
							end if

						end with
					end if ' if blnProcessaRegistro
				next
			
		'	FINALIZAÇÃO
			if n_reg_total <> 0 then 
			'	FECHA ÚLTIMA TABELA
				x = x & "		</table>" & chr(13) & _
						"	</td></tr>" & chr(13) & _
						"</table>" & chr(13)
			'	TOTAIS
				x = x & chr(13) & "<br><br>" & chr(13) & _
					"<table class='Q' style='border-bottom:0px;' cellspacing=0 cellpadding=0>" & chr(13) & _
					"	<tr style='background:azure' nowrap>" & chr(13) & _
					"		<td class='MB' align='center' colspan='4'><span class='C'>&nbsp;TOTAL</span></td>" & chr(13) & _
					"	</tr>" & chr(13)
			'	LEMBRANDO QUE O VETOR ESTÁ ORDENADO
				fabricante_a = "XXXXX"
				for i = Lbound(vProd) to Ubound(vProd)
					with vProd(i)
						if Trim("" & .c1) <> "" then
							v = Split(.c1, "|", -1)
							if Trim("" & v(0)) <> fabricante_a then
								fabricante_a = Trim("" & v(0))
								s = Trim("" & v(0))
								s_aux = ucase(.c5)
								if (s<>"") And (s_aux<>"") then s = s & " - "
								s = s & s_aux
								x = x & "	<tr nowrap>" & chr(13) & _
									"		<td colspan='4' class='MB' align='left'><span class='Cc'>" & _
									s & _
									"</span>" & chr(13) & _
									"	</tr>" & chr(13)
								end if
							
							s_zona = Trim(.c3)
							if s_zona = "" then s_zona = "&nbsp;"
							
							x = x & "	<tr nowrap>" & chr(13) & _
								"		<td class='MDB tdTotCodProd' align='left'><span class='C'>&nbsp;" & _
								Trim("" & v(1)) & "</span></td>"  & chr(13) & _
								"		<td class='MDB tdTotDescrProd' align='left' nowrap><span class='C' nowrap>&nbsp;" & _
								produto_formata_descricao_em_html(produto_descricao_html(v(0), v(1))) & _
								"</span></td>" & chr(13) & _
								"		<td class='MDB tdTotQtde' align='left'><span class='Cd'>&nbsp;" & _
								formata_inteiro(.c2) & "</span></td>"  & chr(13) & _
								"		<td class='MB tdTotZona' align='left'><span class='Cc'>" & s_zona & "</span></td>" & chr(13) & _
								"	</tr>" & chr(13)
							end if
						end with
					next
				x = x & "</table>"
				end if

		  ' MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
			if n_reg_total = 0 then
				x = x & _
					"<table class='Q' width='100%' cellspacing='0'>" & chr(13) & _
					"	<tr>" & chr(13) & _
					"		<td class='ALERTA' align='center'><span class='ALERTA'>&nbsp;NÃO HÁ PEDIDOS PARA SEPARAR.&nbsp;</span></td>" & chr(13) & _
					"	</tr>" & chr(13) & _
					"</table>"
				end if
			
			s_linha_branco = ""
			if n_qtde_zona > 1 then
				s_linha_branco = "	<tr class='notPrint'><td colspan='2' align='left'>&nbsp;</td></tr>" & chr(13) & _
								 "	<tr class='notPrint'><td colspan='2' align='left'>&nbsp;</td></tr>" & chr(13)
				end if
			
			if vZona(iZona).zona_id = COD_ZONA_ID__TODOS then
				s_tit_zona = "Todas"
			else
				s_tit_zona = vZona(iZona).zona_codigo
				end if
			
			x = s_linha_branco & _
				"	<tr id='trTitZona_" & vZona(iZona).zona_id & "' class='notPrint'>" & chr(13) & _
				"		<td nowrap style='width:120px;' align='left'><span class='STP'>Zona: " & s_tit_zona & "</span></td>" & chr(13) & _
				"		<td align='left' class='notPrint' width='90%'><a name='bImprime' href='javascript:fImprime(" & vZona(iZona).zona_id & ")' title='Imprime o relatório da zona: " & s_tit_zona & "' style='margin:3px;'><img name='imgPrinter' src='../botao/Printer.png' border='0'></a>" & _
							"&nbsp;<a name='bExibe' href='javascript:fExibeRelZona(" & vZona(iZona).zona_id & ")' title='Exibe o relatório da zona: " & s_tit_zona & "' style='margin:3px;'><img name='imgView' src='../botao/view_bottom.png' border='0'></a></td>" & chr(13) & _
				"	</tr>" & chr(13) & _
				"	<tr id='trRelZona_" & vZona(iZona).zona_id & "' class='notPrint'>" & chr(13) & _
				"		<td colspan='2' align='center'>" & chr(13) & _
						x & _
				"		</td>" & chr(13) & _
				"	</tr>" & chr(13)

			xRel = xRel & x
			x = ""
			
			strJS_AllTablesCollapse = strJS_AllTablesCollapse & _
									  "	oTrRel = document.getElementById('trRelZona_" & vZona(iZona).zona_id & "');" & chr(13) & _
									  "	oTrRel.style.display = 'none';" & chr(13)
			strJS_AllTablesNotPrint = strJS_AllTablesNotPrint & _
									  "	oTrRel = document.getElementById('trRelZona_" & vZona(iZona).zona_id & "');" & chr(13) & _
									  "	oTrRel.className = 'notPrint';" & chr(13) & _
									  "	oTrTitZona = document.getElementById('trTitZona_" & vZona(iZona).zona_id & "');" & chr(13) & _
									  "	oTrTitZona.className = 'notPrint';" & chr(13)
			end if ' if converte_numero(vZona(iZona).zona_id) <> 0
		next
	
	
'	EXIBE O RELATÓRIO
	xRel = "<table width='649' cellspacing='0' cellpadding='0' border='0'>" & chr(13) & _
			xRel & _
			"</table>" & chr(13)
	Response.write xRel

	if strJS_AllTablesCollapse <> "" then
		strJS_AllTablesCollapse = "function AllTablesCollapse() {" & chr(13) & _
								  "var oTrRel;" & chr(13) & _
								  strJS_AllTablesCollapse & _
								  "}" & chr(13)
		end if
		
	if strJS_AllTablesNotPrint <> "" then
		strJS_AllTablesNotPrint = "function AllTablesNotPrint() {" & chr(13) & _
								  "var divAlertaImpressao, divBody, oTrRel, oTrTitZona;" & chr(13) & _
								  "	divAlertaImpressao = document.getElementById('divAlertaImpressao');" & chr(13) & _
								  "	divAlertaImpressao.className = 'notVisible';" & chr(13) & _
								  "	divBody = document.getElementById('divBody');" & chr(13) & _
								  "	divBody.className = 'notPrint';" & chr(13) & _
								  strJS_AllTablesNotPrint & _
								  "}" & chr(13)
		end if
		
	strScriptJS = "<script language='JavaScript' type='text/javascript'>" & chr(13) & _
				  strJS_AllTablesCollapse & _
				  strJS_AllTablesNotPrint & _
				  "</script>" & chr(13)
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


<%=DOCTYPE_LEGADO%>

<html>


<head>
	<title>CENTRAL</title>
	</head>


<% if False then 'APENAS P/ HABILITAR O INTELLISENSE DURANTE O DESENVOLVIMENTO!! %>
<script src="../Global/jquery.js" language="JavaScript" type="text/javascript"></script>
<% end if %>

<script src="<%=URL_FILE__JQUERY%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
window.status='Aguarde, executando a consulta ...';

function fRELConcluir( id_pedido ) {
	fREL.action = "pedido.asp";
	fREL.pedido_selecionado.value = id_pedido;
	fREL.submit();
}

function fExibeRelZona(idx) {
var oTrRel, oTrTitZona, divAlertaImpressao, divBody, strStyleDisplay;
	oTrRel = document.getElementById("trRelZona_" + idx);
	strStyleDisplay = oTrRel.style.display;
	
	AllTablesNotPrint();
	AllTablesCollapse();

	if (strStyleDisplay == "none") {
		divAlertaImpressao = document.getElementById("divAlertaImpressao");
		divAlertaImpressao.className = "notVisible notPrint";
		divBody = document.getElementById("divBody");
		divBody.className = "";
		oTrTitZona = document.getElementById("trTitZona_" + idx);
		oTrTitZona.className = "";
		oTrRel.style.display = "";
		oTrRel.className = "";
	}
}

function fImprime(idx) {
var oTrRel, oTrTitZona, divAlertaImpressao, divBody;
	AllTablesNotPrint();
	AllTablesCollapse();
	divAlertaImpressao = document.getElementById("divAlertaImpressao");
	divAlertaImpressao.className = "notVisible notPrint";
	divBody = document.getElementById("divBody");
	divBody.className = "";
	oTrTitZona = document.getElementById("trTitZona_" + idx);
	oTrTitZona.className = "";
	oTrRel = document.getElementById("trRelZona_" + idx);
	oTrRel.style.display = "";
	oTrRel.className = "";
	print();
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
P.C { font-size:10pt; }
P.Cc { font-size:10pt; }
P.Cd { font-size:10pt; }
.Nni
{
	font-family: Arial, Helvetica, sans-serif;
	font-size: 10pt;
	font-weight: normal;
	font-style: italic;
}
.tdPedido
{
	vertical-align:bottom;
	width:75px;
	border-right:1px solid;
}
.tdQtVol
{
	vertical-align:bottom;
	text-align:center;
	border-right:1px solid;
	width:35px;
}
.tdNF
{
	vertical-align:bottom;
	width:70px;
	border-right:1px solid;
}
.tdLoja
{
	vertical-align:bottom;
	width:45px;
	border-right:1px solid;
}
.tdCliente
{
	vertical-align:bottom;
	width:337px;
	border-right:1px solid;
}
.tdTransp
{
	vertical-align:bottom;
	width:80px;
}
.tdQtde
{
	vertical-align:bottom;
	width:40px;
}
.tdProd
{
	vertical-align:bottom;
}
.tdZona
{
	vertical-align:bottom;
	width:30px;
}
.tdTotCodProd
{
	vertical-align:bottom;
	width:65px;
}
.tdTotDescrProd
{
	vertical-align:bottom;
	width:400px;
}
.tdTotQtde
{
	vertical-align:bottom;
	width:50px;
}
.tdTotZona
{
	vertical-align:bottom;
	width:30px;
}
</style>


<% if alerta <> "" then %>
<!-- ************************************************************ -->
<!-- **********  PÁGINA PARA EXIBIR MENSAGENS DE ERRO  ********** -->
<!-- ************************************************************ -->
<body>
<center>
<br>
<!--  T E L A  -->
<p class="T">A V I S O</p>
<div class="MtAlerta" style="width:600px;font-weight:bold;" align="center"><p style='margin:5px 2px 5px 2px;'><%=alerta%></p></div>
<% if s_table_produtos_sem_zona <> "" then Response.Write "<br>" & s_table_produtos_sem_zona%>
<br><br>
<p class="TracoBottom"></p>
<table cellspacing="0">
<tr>
	<td align="center"><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()"><img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
</tr>
</table>
</center>
</body>



<% else %>
<!-- ***************************************************** -->
<!-- **********  PÁGINA PARA EXIBIR RESULTADO   ********** -->
<!-- ***************************************************** -->
<body onload="window.status='Concluído';AllTablesCollapse();AllTablesNotPrint();" link=#000000 alink=#000000 vlink=#000000>
<center>
<div id="divAlertaImpressao" class="notVisible"><p class="ALERTA">Para imprimir os dados do relatório,<br />clique no ícone de impressora ao lado do título de cada zona.</p></div>

<div id="divBody" class="notPrint">
<form id="fREL" name="fREL" method="post">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="pedido_selecionado" id="pedido_selecionado" value=''>
<input type="hidden" name="nsu_selecionado" id="nsu_selecionado" value="<%=nsu_selecionado%>">

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Separação (Zona) - Consulta</span>
	<br class="notPrint"><span class="Rc notPrint">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>

<%
	s_filtro ="<table width='649' cellpadding='0' cellspacing='0' style='border-bottom:1px solid black'>" & chr(13)
	
'	PERÍODO
	s = ""
	s_aux = c_dt_inicio
	if s_aux = "" then s_aux = "N.I."
	s = s & s_aux & " a "
	s_aux = c_dt_termino
	if s_aux = "" then s_aux = "N.I."
	s = s & s_aux
	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' nowrap><span class='Nni'>Período:&nbsp;</span></td>" & chr(13) & _
				"		<td align='left' valign='top' width='99%'><span class='N'>" & s & "</span></td>" & chr(13) & _
				"	</tr>" & chr(13)
'	NFe EMITIDA
	s = ""
	s_aux = rb_nfe
	if s_aux = "EMITIDA" then
		s = "Somente com NFe já emitida"
	elseif s_aux = "NAO_EMITIDA" then
		s = "Somente com NFe não emitida"
	else
		s = "Ambos"
		end if
	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' nowrap><span class='Nni'>NFe Emitida:&nbsp;</span></td>" & chr(13) & _
				"		<td align='left' valign='top' width='99%'><span class='N'>" & s & "</span></td>" & chr(13) & _
				"	</tr>" & chr(13)
'	TRANSPORTADORA
	s = ""
	s_aux = c_transportadora
	if s_aux = "" then s_aux = "N.I."
	s = s & s_aux
	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' nowrap><span class='Nni'>Transportadora:&nbsp;</span></td>" & chr(13) & _
				"		<td align='left' valign='top' width='99%'><span class='N'>" & s & "</span></td>" & chr(13) & _
				"	</tr>" & chr(13)
'	QTDE MÁXIMA DE PEDIDOS
	s = ""
	s_aux = c_qtde_max_pedidos
	if s_aux = "" then s_aux = "N.I."
	s = s & s_aux
	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' nowrap><span class='Nni'>Qtde Máxima de Pedidos:&nbsp;</span></td>" & chr(13) & _
				"		<td align='left' valign='top' width='99%'><span class='N'>" & s & "</span></td>" & chr(13) & _
				"	</tr>" & chr(13)
'	NSU RELATÓRIO
	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' nowrap><span class='Nni'>NSU do Relatório:&nbsp;</span></td>" & chr(13) & _
				"		<td align='left' valign='top' width='99%'><span id='spanFiltroNsuRelatorio' class='N'>" & nsu_selecionado & "</span></td>" & chr(13) & _
				"	</tr>" & chr(13)
'	EMITIDO EM
	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' nowrap><span class='Nni'>Emitido em:&nbsp;</span></td>" & chr(13) & _
				"		<td align='left' valign='top' width='99%'><span id='spanFiltroDataEmissao' class='N'>" & c_dt_hr_emissao & "</span></td>" & chr(13) & _
				"	</tr>" & chr(13)
'	EMITIDO POR
	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' nowrap><span class='Nni'>Emitido por:&nbsp;</span></td>" & chr(13) & _
				"		<td align='left' valign='top' width='99%'><span id='spanFiltroUsuarioEmissao' class='N'>" & c_usuario_emissao & "</span></td>" & chr(13) & _
				"	</tr>" & chr(13)

	s_filtro = s_filtro & "</table>" & chr(13)

	Response.Write s_filtro
%>
<br>


<!--  RELATÓRIO  -->
<% consulta_relatorio %>

<!-- ************   SEPARADOR   ************ -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc" align="left">&nbsp;</td></tr>
</table>
<br>


<table class="notPrint" width="649" cellspacing="0">
<tr>
	<td align="center"><a name="bVOLTA" id="bVOLTA" href="javascript:history.back()" title="volta para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
</tr>
</table>

</form>
</div>

</center>
</body>

<%=strScriptJS%>

<% end if %>

</html>


<%
'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>
