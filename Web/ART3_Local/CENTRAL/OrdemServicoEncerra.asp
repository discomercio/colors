<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =============================================================
'	  O R D E M S E R V I C O E N C E R R A . A S P
'     =============================================================
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

	dim s, usuario, i, j, n, flag_ok, msg_erro
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

	dim url_back
	dim s_id_nfe_emitente
	dim s_tipo, s_op_descricao, s_num_OS, s_chave_OS
	dim s_descricao_volume, s_num_serie, s_tipo_vol, s_obs_problema
	dim s_ckb_spe, s_ckb_spe_descricao
	dim v_aux, s_cod_estoque_origem, s_cod_estoque_destino, s_fluxo
	dim s_loja, s_nome_loja, s_pedido
	
'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	
	dim max_qtde_itens
	max_qtde_itens = obtem_parametro_OrdemServico_Volumes_MaxQtdeItens
	
'	OBTÉM DADOS DO FORMULÁRIO
	url_back = Trim(Request("url_back"))
	s_id_nfe_emitente = Trim(Request.Form("c_id_nfe_emitente"))
	s_tipo = Ucase(Trim(Request.Form("rb_tipo")))
	s_op_descricao = Trim(Request.Form("op_selecionada_descricao"))
	s_num_OS = Ucase(Trim(Request.Form("c_num_OS")))
	s_chave_OS = normaliza_codigo(retorna_so_digitos(s_num_OS), TAM_MAX_NSU)
	s_ckb_spe = Ucase(Trim(Request.Form("ckb_spe")))
	s_ckb_spe_descricao = Trim(Request.Form("ckb_spe_descricao"))
	s_loja = Trim(Request.Form("c_loja"))
	s_pedido = Ucase(Trim(Request.Form("c_pedido")))
	if s_pedido <> "" then s_pedido = normaliza_num_pedido(s_pedido)

	dim alerta
	alerta=""
	
	dim observacao
	observacao=""

	if s_id_nfe_emitente = "" then
		alerta=texto_add_br(alerta)
		alerta=alerta & "Não foi informada a empresa (CD)."
		end if

	if InStr(s_tipo, "TRANSF_") > 0 then
		v_aux = Split(s_tipo, "_")
		s_cod_estoque_origem = v_aux(Ubound(v_aux)-1)
		s_cod_estoque_destino = v_aux(Ubound(v_aux))
		s_fluxo = "TRANSF"
		if (s_cod_estoque_destino<>ID_ESTOQUE_SHOW_ROOM)And(s_cod_estoque_destino<>ID_ESTOQUE_DEVOLUCAO) then s_loja=""
		if s_cod_estoque_destino<>ID_ESTOQUE_DEVOLUCAO then s_pedido = ""
	elseif InStr(s_tipo, "ENT_") > 0 then
		s_cod_estoque_origem = ID_ESTOQUE_DANIFICADOS
		s_cod_estoque_destino = ID_ESTOQUE_VENDA
		s_fluxo = Left(s_tipo, 3)
		s_loja=""
		s_pedido=""
	else
		alerta = "Operação desconhecida."
		end if
	
'	CONSISTE DADOS DIGITADOS
	if alerta = "" then
		if s_chave_OS = "" then
			alerta = "Nº da Ordem de Serviço não foi informado."
			end if
		end if

	if alerta = "" then
		if s_tipo = "" then
			alerta = "Não foi indicado o tipo de transferência/movimentação do estoque a ser efetuado."
		elseif s_fluxo = "" then
			alerta = "Operação selecionada é inválida."
			end if
		end if

	if alerta = "" then
		if s_cod_estoque_destino = "" then
			alerta = "Falha ao determinar o código do estoque de destino."
			end if
		end if
		
	if alerta = "" then
		if s_fluxo = "TRANSF" then
			if (s_loja="") And ((s_cod_estoque_destino=ID_ESTOQUE_SHOW_ROOM)Or(s_cod_estoque_destino=ID_ESTOQUE_DEVOLUCAO)) then
				alerta = "Número da loja não especificado."
			elseif (s_pedido="") And (s_cod_estoque_destino=ID_ESTOQUE_DEVOLUCAO) then
				alerta = "Número do pedido não foi especificado."
				end if
		else
			if s_fluxo <> "ENT" then
				alerta = "Operação selecionada é desconhecida."
				end if
			end if
		end if

	if alerta = "" then
		if s_cod_estoque_origem <> ID_ESTOQUE_DANIFICADOS then
			alerta=texto_add_br(alerta)
			alerta=alerta & "Estoque de origem é inválido para esta operação."
			end if
		end if

	if alerta = "" then
	'	VERIFICA A LOJA
		if s_loja <> "" then
			s = "SELECT " & _
					"*" & _
				" FROM t_LOJA" & _
				" WHERE" & _
					" (loja = '" & s_loja & "')"
			if rs.State <> 0 then rs.Close
			rs.open s, cn
			if rs.Eof then
				alerta=texto_add_br(alerta)
				alerta=alerta & "Loja " & s_loja & " NÃO está cadastrada."
			else
				s_nome_loja = Trim("" & rs("nome"))
				if s_nome_loja = "" then s_nome_loja = Trim("" & rs("razao_social"))
				end if
			end if
		end if

	dim r_OS, r_OS_item
	if alerta = "" then
		if Not le_ordem_servico(s_chave_OS, r_OS, msg_erro) then 
			alerta = msg_erro
		else
			if Not le_ordem_servico_item(s_chave_OS, r_OS_item, msg_erro) then alerta = msg_erro
			'Assegura que dados cadastrados anteriormente sejam exibidos corretamente, mesmo se o parâmetro da quantidade máxima de itens tiver sido reduzido
			if VectorLength(r_OS_item) > max_qtde_itens then max_qtde_itens = VectorLength(r_OS_item)
			end if
		
		if converte_numero(s_id_nfe_emitente) <> converte_numero(r_OS.id_nfe_emitente) then
			alerta=texto_add_br(alerta)
			alerta=alerta & "A ordem de serviço nº " & formata_num_OS_tela(s_chave_OS) & " não está vinculada ao CD '" & obtem_apelido_empresa_NFe_emitente(s_id_nfe_emitente) & "'"
			end if
		end if

	if alerta = "" then
		if r_OS.situacao_status <> ST_OS_EM_ANDAMENTO then
			alerta = "A ordem de serviço nº " & formata_num_OS_tela(s_chave_OS) & " não pode ser encerrada porque está com status: " & Ucase(x_OS_status(r_OS.situacao_status))
			end if
		end if

	dim r_cliente, s_nome_contato
	set r_cliente = New cl_CLIENTE
	s_nome_contato = ""
	if alerta = "" then
		if r_OS.id_cliente <> "" then
			if x_cliente_bd(r_OS.id_cliente, r_cliente) then
				s_nome_contato = Trim(r_cliente.contato)
				if s_nome_contato <> "" then 
					s_nome_contato = "  (contato: " & s_nome_contato & ")"
					end if
				end if
			end if
		end if
	
	dim r_orcamentista_e_indicador
	dim s_telefone_indicador, s_tel_aux_1, s_tel_aux_2
	s_telefone_indicador = ""
	if alerta = "" then
		if r_OS.indicador <> "" then
			if le_orcamentista_e_indicador(r_OS.indicador, r_orcamentista_e_indicador, msg_erro) then
				with r_orcamentista_e_indicador
					s_tel_aux_1 = formata_ddd_telefone_ramal(.ddd, .telefone, "")
					s_tel_aux_2 = formata_ddd_telefone_ramal(.ddd_cel, .tel_cel, "")
					if (s_tel_aux_1 <> "") And (s_tel_aux_2 <> "") then
						s_telefone_indicador = s_tel_aux_1 & " / " & s_tel_aux_2
					else
						s_telefone_indicador = s_tel_aux_1 & s_tel_aux_2
						end if
					if s_telefone_indicador <> "" then s_telefone_indicador = "  (Tel: " & s_telefone_indicador & ")"
					end with
				end if
			end if
		end if

	dim r_item
	set r_item = New cl_ITEM_PEDIDO
	if alerta = "" then
		s=retorna_so_digitos(Request.Form("c_fabricante")(1))
		if s = "" then
			alerta=texto_add_br(alerta)
			alerta=alerta & "Não foi informado o fabricante."
		else
			r_item.fabricante = normaliza_codigo(s, TAM_MIN_FABRICANTE)
			end if
			
		s=Trim(Request.Form("c_codigo")(1))
		if s = "" then
			alerta=texto_add_br(alerta)
			alerta=alerta & "Não foi informado o produto."
		else
			r_item.produto=Ucase(Trim(s))
			end if
			
		s = Trim(Request.Form("c_qtde")(1))
		if IsNumeric(s) then r_item.qtde = CLng(s) else r_item.qtde = 0
		if r_item.qtde <> 1 then
			alerta=texto_add_br(alerta)
			alerta=alerta & "Quantidade inválida: uma única unidade deve ser transferida por vez."
			end if
		end if
	
	if alerta = "" then
	'	VERIFICA O PRODUTO A SER TRANSFERIDO
		with r_item
			if .qtde <= 0 then
				alerta=texto_add_br(alerta)
				alerta=alerta & "Produto " & .produto & ": quantidade " & cstr(.qtde) & " é inválida."
				end if
				
			if (Not IsEAN(.produto)) And (.fabricante="") then
				alerta=texto_add_br(alerta)
				alerta=alerta & "Não foi especificado o fabricante do produto " & .produto & "."
			else
				s = "SELECT * FROM t_PRODUTO WHERE"
				if IsEAN(.produto) then
					s = s & " (ean='" & .produto & "')"
				else
					s = s & " (fabricante='" & .fabricante & "') AND (produto='" & .produto & "')"
					end if
				
				if rs.State <> 0 then rs.Close
				rs.open s, cn
				if rs.Eof then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Produto " & .produto & " NÃO está cadastrado."
				else
					flag_ok = True
					if IsEAN(.produto) And (.fabricante<>"") then
						if (.fabricante<>Trim("" & rs("fabricante"))) then
							flag_ok = False
							alerta=texto_add_br(alerta)
							alerta=alerta & "Produto " & .produto & " NÃO pertence ao fabricante " & .fabricante & "."
							end if
						end if
					if flag_ok then
					'	CARREGA CÓDIGO INTERNO DO PRODUTO
						.fabricante = Trim("" & rs("fabricante"))
						.produto = Trim("" & rs("produto"))	
						.ean = Trim("" & rs("ean"))
						.descricao = Trim("" & rs("descricao"))
						.descricao_html = Trim("" & rs("descricao_html"))
						end if
					end if
				end if
			end with
		end if

'	VERIFICA SE A ORDEM DE SERVIÇO SE REFERE AO PRODUTO INDICADO
	if alerta = "" then
		if (r_OS.fabricante <> r_item.fabricante) Or (r_OS.produto <> r_item.produto) then
			alerta=texto_add_br(alerta)
			alerta=alerta & "A ordem de serviço nº " & formata_num_OS_tela(s_chave_OS) & " não se refere ao produto " & r_item.produto & "."
			end if
		end if
		
'	VERIFICA DISPONIBILIDADE NO ESTOQUE
	if alerta = "" then
		with r_item
		'	QUANTIDADE DE PRODUTOS A SER TRANSFERIDA
		'	IMPORTANTE: NA TABELA T_ESTOQUE_MOVIMENTO, SOMENTE O ESTOQUE LÓGICO 'SPE' (SEM PRESENÇA NO ESTOQUE) NÃO POSSUI CONTEÚDO NO CAMPO 'id_estoque'.
			s = "SELECT" & _
					" SUM(qtde) AS total" & _
				" FROM t_ESTOQUE_MOVIMENTO tEM" & _
					" INNER JOIN t_ESTOQUE tE ON (tEM.id_estoque = tE.id_estoque)" & _
				" WHERE" & _
					" (tEM.anulado_status=0)" & _
					" AND (tE.id_nfe_emitente = " & s_id_nfe_emitente & ")" & _
					" AND (tEM.fabricante='" & Trim(.fabricante) & "')" & _
					" AND (tEM.produto='" & Trim(.produto) & "')" & _
					" AND (tEM.estoque='" & ID_ESTOQUE_DANIFICADOS & "')" & _
					" AND (tEM.id_ordem_servico='" & s_chave_OS & "')"
			if rs.State <> 0 then rs.Close
			rs.open s, cn
			j=0
			if Not rs.Eof then 
				if Not IsNull(rs("total")) then j = CLng(rs("total"))
				end if
			if .qtde > j then
				alerta=texto_add_br(alerta)
				alerta=alerta & "Faltam " & CStr(.qtde-j) & " unidades do produto " & .produto & " do fabricante " & .fabricante & "."
				end if
			end with
		end if

	if alerta <> "" then 
		alerta = texto_add_br(s_op_descricao) & alerta
		end if

	if alerta = "" then
		if s_pedido <> "" then
			s = "SELECT " & _
					"*" & _
				" FROM t_PEDIDO" & _
				" WHERE" & _
					" (pedido = '" & s_pedido & "')"
			if rs.State <> 0 then rs.Close
			rs.open s, cn
			if Not rs.Eof then
				if converte_numero(s_id_nfe_emitente) <> converte_numero(rs("id_nfe_emitente")) then
					alerta=texto_add_br(alerta)
					alerta=alerta & "O pedido " & s_pedido & " não está vinculado ao CD '" & obtem_apelido_empresa_NFe_emitente(s_id_nfe_emitente) & "'"
					end if
				end if
			end if
		end if

'	O DESTINO DO PRODUTO É DIFERENTE DE ONDE ELE SE ORIGINOU AO ABRIR A O.S.?
	if alerta = "" then
		if r_OS.cod_estoque_origem <> s_cod_estoque_destino then
			observacao=texto_add_br(observacao)
			observacao=observacao & "O produto desta O.S. veio do estoque '" & x_estoque(r_OS.cod_estoque_origem) & "' e está sendo transferido para o estoque '" & x_estoque(s_cod_estoque_destino) & "'"
			end if
	'	O ESTOQUE DE ORIGEM E DE DESTINO ESTÃO RELACIONADOS A UMA LOJA?
		if (Trim(r_OS.loja_estoque_origem) <> "") And (s_loja <> "") And (Trim(r_OS.loja_estoque_origem) <> s_loja) then
			observacao=texto_add_br(observacao)
			observacao=observacao & "O produto desta O.S. estava vinculado à loja '" & r_OS.loja_estoque_origem & "' e agora será vinculado à loja '" & s_loja & "'"
			end if
	'	O ESTOQUE DE ORIGEM E DE DESTINO ESTÃO RELACIONADOS A UM PEDIDO?
		if (Trim(r_OS.pedido) <> "") And (s_pedido <> "") And (Trim(r_OS.pedido) <> s_pedido) then
			observacao=texto_add_br(observacao)
			observacao=observacao & "O produto desta O.S. estava vinculado ao pedido '" & Trim(r_OS.pedido) & "' e agora será vinculado ao pedido '" & s_pedido & "'"
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
function fOPConfirma( f ) {
	dCONFIRMA.style.visibility="hidden";
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
<!-- *************************************************************** -->
<!-- **********  PÁGINA PARA EXIBIR DADOS DE CONFIRMAÇÃO  ********** -->
<!-- *************************************************************** -->
<body onload="focus();">
<center>

<form id="fOP" name="fOP" method="post" action="OrdemServicoEncerraConfirma.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="c_num_OS" id="c_num_OS" value="<%=s_num_OS%>">
<input type="hidden" name="rb_tipo" id="rb_tipo" value="<%=s_tipo%>">
<input type="hidden" name="op_selecionada_descricao" id="op_selecionada_descricao" value="<%=s_op_descricao%>">
<input type="hidden" name="ckb_spe" id="ckb_spe" value="<%=s_ckb_spe%>">
<input type="hidden" name="ckb_spe_descricao" id="ckb_spe_descricao" value="<%=s_ckb_spe_descricao%>">
<input type="hidden" name="c_loja" id="c_loja" value="<%=s_loja%>">
<input type="hidden" name="c_pedido" id="c_pedido" value="<%=s_pedido%>">
<input type="hidden" name="c_id_nfe_emitente" id="c_id_nfe_emitente" value="<%=s_id_nfe_emitente%>" />
<input type="hidden" name="url_back" id="url_back" value="<%=url_back%>" />

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="749" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">

<tr>
	<td align="right" valign="bottom"><p class="PEDIDO"><%=s_op_descricao%><br>Encerra Ordem de Serviço Nº <%=formata_num_OS_tela(s_chave_OS)%><span class="C">&nbsp;</span></p></td>
</tr>
</table>

<% if observacao <> "" then %>
	<br><br>
	<span class="Lbl">OBSERVAÇÕES</span>
	<div class='MtAviso' style="width:749px;font-weight:bold;border:1pt solid black;" align="center"><P style='margin:5px 2px 5px 2px;'><%=observacao%></p></div>
	<br>
<% end if %>

<br>
<%if s_fluxo="ENT" then%>
	<input type="checkbox" tabindex="-1" id="ckb_spe_aux" name="ckb_spe_aux" value="SPE_ON" disabled
		<%if s_ckb_spe<>"" then Response.Write " checked"%>
	><span class="C" id="sckb_spe" name="sckb_spe" style="cursor:default;font-weight:normal;font-style:normal;" 
		><%=s_ckb_spe_descricao%></span>
	<br><br>
<%end if%>


<table class="Qx" cellSpacing="0">
	<!--  TÍTULO  -->
	<tr bgColor="#FFFFFF">
	<td colspan="5" class="MT" valign="middle" align="center" NOWRAP style="background:azure;"><span class="PLTc" style="vertical-align:middle;"
		>ORDEM DE SERVIÇO</span></td>
	</tr>
<!--  CADASTRADO POR  -->
	<tr bgColor="#FFFFFF">
	<td class="MDBE" NOWRAP align="right"><span class="PLTe" style="vertical-align:middle;">Cadastrado por</span></td>
	<td class="MDB" colspan="4">
		<input name="c_cadastrado_por" id="c_cadastrado_por" READONLY tabindex=-1 class="PLLe" style="width:460px;margin-left:2pt;" 
				value="<%=r_OS.usuario%>"></td>
	</tr>
	<tr bgColor="#FFFFFF">
	<td class="MDBE" NOWRAP align="right"><span class="PLTe" style="vertical-align:middle;">Estoque origem</span></td>
	<td class="MDB" colspan="4">
		<input name="c_estoque_origem_aux" id="c_estoque_origem_aux" READONLY tabindex=-1 class="PLLe" style="width:460px;margin-left:2pt;" 
				value="<%=x_estoque(r_OS.cod_estoque_origem)%>"></td>
	</tr>
	<% if r_OS.loja_estoque_origem <> "" then %>
		<%	s = r_OS.loja_estoque_origem & " - " & x_loja(r_OS.loja_estoque_origem) %>
		<tr bgColor="#FFFFFF">
		<td class="MDBE" NOWRAP align="right"><span class="PLTe" style="vertical-align:middle;">Loja origem</span></td>
		<td class="MDB" colspan="4">
			<input name="c_loja_aux" id="c_loja_aux" READONLY tabindex=-1 class="PLLe" style="width:460px;margin-left:2pt;" 
					value="<%=s%>"></td>
		</tr>
	<% end if %>
	<tr bgColor="#FFFFFF">
	<td class="MDBE" NOWRAP align="right"><span class="PLTe" style="vertical-align:middle;">Empresa (CD)</span></td>
	<td class="MDB" colspan="4">
		<input name="c_id_nfe_emitente_apelido" id="c_id_nfe_emitente_apelido" READONLY tabindex=-1 class="PLLe" style="width:460px;margin-left:2pt;" 
				value="<%=obtem_apelido_empresa_NFe_emitente(r_OS.id_nfe_emitente)%>"></td>
	</tr>

<tr><td colspan="5">&nbsp;</td></tr>

<% if r_OS.pedido <> "" then %>

	<!--  TÍTULO  -->
	<tr bgColor="#FFFFFF">
	<td colspan="5" class="MT" valign="middle" align="center" NOWRAP style="background:azure;"><span class="PLTc" style="vertical-align:middle;"
		>DADOS DO PEDIDO</span></td>
	</tr>
<!--  PEDIDO  -->
	<tr bgColor="#FFFFFF">
	<td class="MDBE" NOWRAP align="right"><span class="PLTe" style="vertical-align:middle;">Pedido</span></td>
	<td class="MDB" colspan="4">
		<a href="Pedido.asp?pedido_selecionado=<%=r_OS.pedido%><%= "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="clique para consultar o pedido">
		<input name="c_pedido_aux" id="c_pedido_aux" READONLY tabindex=-1 class="PLLe" style="width:70px;margin-left:2pt;cursor:pointer;" 
				value="<%=r_OS.pedido%>">
		</a></td>
	</tr>
<!--  NF  -->
	<tr bgColor="#FFFFFF">
	<td class="MDBE" NOWRAP align="right"><span class="PLTe" style="vertical-align:middle;">NF</span></td>
	<td class="MDB" colspan="4">
		<input name="c_nf" id="c_nf" READONLY tabindex=-1 class="PLLe" style="width:460px;margin-left:2pt;" 
				value="<%=r_OS.nf%>"></td>
	</tr>
<!--  NOME DO CLIENTE  -->
	<tr bgColor="#FFFFFF">
	<td class="MDBE" NOWRAP align="right"><span class="PLTe" style="vertical-align:middle;">Cliente</span></td>
	<td class="MDB" colspan="4">
		<input name="c_nome_cliente" id="c_nome_cliente" READONLY tabindex=-1 class="PLLe" style="width:460px;margin-left:2pt;" 
				value="<%=r_OS.nome_cliente & s_nome_contato%>"></td>
	</tr>
<!--  ENDEREÇO  -->
	<%
	with r_OS
		s = formata_endereco(.endereco, .endereco_numero, .endereco_complemento, .bairro, .cidade, .uf, .cep)
		end with
	%>
	<tr bgColor="#FFFFFF">
	<td class="MDBE" NOWRAP align="right" valign="top"><span class="PLTe" style="vertical-align:middle;">Endereço</span></td>
	<td class="MDB" colspan="4">
		<textarea rows="<%=Cstr(MAX_LINHAS_OS_ENDERECO)%>" name="c_endereco" id="c_endereco" READONLY tabindex=-1 class="PLLe" style="width:100%;margin-left:2pt;"><%=s%></textarea></td>
	</tr>
<% if r_OS.tipo_cliente = ID_PF then %>
<!--  TELEFONE  -->
	<tr bgColor="#FFFFFF">
	<td class="MDBE" NOWRAP align="right"><span class="PLTe" style="vertical-align:middle;">Tel Res</span></td>
	<td class="MDB" colspan="4">
		<input name="c_tel_res" id="c_tel_res" READONLY tabindex=-1 class="PLLe" style="width:460px;margin-left:2pt;" 
				value="<%=formata_ddd_telefone_ramal(r_OS.ddd_res, r_OS.tel_res, "")%>"></td>
	</tr>
	<tr bgColor="#FFFFFF">
	<td class="MDBE" NOWRAP align="right"><span class="PLTe" style="vertical-align:middle;">Tel Com</span></td>
	<td class="MDB" colspan="4">
		<input name="c_tel_com" id="c_tel_com" READONLY tabindex=-1 class="PLLe" style="width:460px;margin-left:2pt;" 
				value="<%=formata_ddd_telefone_ramal(r_OS.ddd_com, r_OS.tel_com, r_OS.ramal_com)%>"></td>
	</tr>
<% else %>
<!--  TELEFONE  -->
	<tr bgColor="#FFFFFF">
	<td class="MDBE" NOWRAP align="right"><span class="PLTe" style="vertical-align:middle;">Telefone</span></td>
	<td class="MDB" colspan="4">
		<input name="c_telefone" id="c_telefone" READONLY tabindex=-1 class="PLLe" style="width:460px;margin-left:2pt;" 
				value="<%=formata_ddd_telefone_ramal(r_OS.ddd_com, r_OS.tel_com, r_OS.ramal_com)%>"></td>
	</tr>
<%end if%>
<!--  INDICADOR  -->
	<tr bgColor="#FFFFFF">
	<td class="MDBE" NOWRAP align="right"><span class="PLTe" style="vertical-align:middle;">Indicador</span></td>
	<td class="MDB" colspan="4">
		<input name="c_indicador" id="c_indicador" READONLY tabindex=-1 class="PLLe" style="width:460px;margin-left:2pt;" 
				value="<%=r_OS.indicador & s_telefone_indicador%>"></td>
	</tr>

<tr><td colspan="5">&nbsp;</td></tr>
<% end if %>

<!--  P R O D U T O  -->
	<!--  TÍTULO DA TABELA  -->
	<tr bgColor="#FFFFFF">
	<td colspan="5" class="MT" valign="middle" align="center" NOWRAP style="background:azure;"><span class="PLTc" style="vertical-align:middle;"
		>PRODUTO</span></td>
	</tr>
	<!--  TÍTULO DAS COLUNAS  -->
	<tr bgColor="#FFFFFF">
	<td class="MDBE"><p class="PLTe">Fabr</p></td>
	<td class="MDB"><p class="PLTe">Produto</p></td>
	<td class="MDB"><p class="PLTe">EAN</p></td>
	<td class="MDB"><p class="PLTe">Descrição</p></td>
	<td class="MDB"><p class="PLTd">Qtde</p></td>
	</tr>

<%	i=1 %>
	<tr>
	<td class="MDBE">
		<input name="c_fabricante" id="c_fabricante" READONLY tabindex=-1 class="PLLe" style="width:30px;"
			value="<%=r_OS.fabricante%>"></td>
	<td class="MDB">
		<input name="c_produto" id="c_produto" READONLY tabindex=-1 class="PLLe" style="width:55px;"
			value="<%=r_OS.produto%>"></td>
	<td class="MDB">
		<input name="c_ean" id="c_ean" READONLY tabindex=-1 class="PLLe" style="width:85px;"
			value="<%=r_OS.ean%>"></td>
	<td class="MDB" style="width:277px;">
		<span class="PLLe"><%=produto_formata_descricao_em_html(r_OS.descricao_html)%></span>
		<input type="hidden" name="c_descricao" id="c_descricao" value="<%=r_OS.descricao%>">
	</td>
	<td class="MDB" align="right">
		<input name="c_qtde" id="c_qtde" READONLY tabindex=-1 class="PLLd" style="width:35px;"
			value="<%=Cstr(r_OS.qtde)%>"></td>
	</tr>


<tr><td colspan="5">&nbsp;</td></tr>
	<!--  TÍTULO DA TABELA  -->
	<tr bgColor="#FFFFFF">
	<td colspan="5" class="MT" valign="middle" align="center" NOWRAP style="background:azure;"><span class="PLTc" style="vertical-align:middle;"
		>VOLUMES</span></td>
	</tr>

<!--  R E L A Ç Ã O   D E   V O L U M E S  -->
	<tr bgColor="#FFFFFF">
	<td class="MDBE"><p class="PLTe">Volume</p></td>
	<td class="MDB"><p class="PLTe">Tipo</p></td>
	<td class="MDB"><p class="PLTe">Nº Série</p></td>
	<td class="MDB" colspan="2"><p class="PLTe">Problema</p></td>
	</tr>
<%  n = Lbound(r_OS_item)-1
	for i=1 to max_qtde_itens
		n = n+1
		if n <= Ubound(r_OS_item) then
			with r_OS_item(n)
				s_descricao_volume = .descricao_volume
				s_tipo_vol = .tipo
				s_num_serie = .num_serie
				s_obs_problema = .obs_problema
				end with
		else
			s_descricao_volume = ""
			s_tipo_vol = ""
			s_num_serie = ""
			s_obs_problema = ""
			end if
%>
	<tr>
	<td class="MDBE" valign="top"><input name="c_descricao_volume" id="c_descricao_volume" 
		READONLY tabindex=-1 class="PLLe" maxlength="12" 
		style="width:100px;" onkeypress="if (digitou_enter(true)&&(tem_info(this.value)||(<%=Cstr(i)%>!=1))) if (trim(this.value)=='') bCONFIRMA.focus(); else fOP.c_tipo[<%=Cstr(i-1)%>].focus(); filtra_nome_identificador();" onblur="this.value=trim(this.value);"
		value='<%=s_descricao_volume%>'></td>
	<td class="MDB" valign="top"><input name="c_tipo" id="c_tipo" 
		READONLY tabindex=-1 class="PLLe" maxlength="12" 
		style="width:100px;" onkeypress="if (digitou_enter(true)) fOP.c_num_serie[<%=Cstr(i-1)%>].focus(); filtra_nome_identificador();" onblur="this.value=trim(this.value);"
		value='<%=s_tipo_vol%>'></td>
	<td class="MDB" valign="top"><input name="c_num_serie" id="c_num_serie" 
		READONLY tabindex=-1 class="PLLe" maxlength="20" 
		style="width:130px;" onkeypress="if (digitou_enter(true)) fOP.c_obs_problema[<%=Cstr(i-1)%>].focus(); filtra_nome_identificador();" onblur="this.value=trim(this.value);"
		value='<%=s_num_serie%>'></td>
	<td class="MDB" colspan="2" align="right" style="width:344px;"><textarea name="c_obs_problema" id="c_obs_problema" rows="<%=Cstr(MAX_LINHAS_OS_OBS_PROBLEMA)%>" 
		READONLY tabindex=-1 class="PLLe" onkeypress="return maxLength(this,MAX_TAM_OS_OBS_PROBLEMA);" onpaste="return maxLengthPaste(this,MAX_TAM_OS_OBS_PROBLEMA);" 
		style="width:340px;" onkeypress="if (digitou_enter(true)) {if (<%=Cstr(i)%>==fOP.c_obs_problema.length) bCONFIRMA.focus(); else fOP.c_descricao_volume[<%=Cstr(i)%>].focus();} filtra_nome_identificador();"
		><%=s_obs_problema%></textarea></td>
	</tr>
<% next %>


<tr><td colspan="5">&nbsp;</td></tr>

<!--  PEÇAS NECESSÁRIAS  -->
	<!--  TÍTULO  -->
	<tr bgColor="#FFFFFF">
	<td colspan="5" class="MT" valign="middle" align="center" NOWRAP style="background:azure;"><span class="PLTc" style="vertical-align:middle;"
		>PEÇAS NECESSÁRIAS</span></td>
	</tr>
	<tr>
	<td colspan="5" class="MDBE" align="right" style="width:685px;"><textarea name="c_obs_pecas_necessarias" id="c_obs_pecas_necessarias" rows="<%=Cstr(MAX_LINHAS_OS_OBS_PECAS_NECESSARIAS)%>" 
		READONLY tabindex=-1 class="PLLe" onkeypress="return maxLength(this,MAX_TAM_OS_OBS_PECAS_NECESSARIAS);" onpaste="return maxLengthPaste(this,MAX_TAM_OS_OBS_PECAS_NECESSARIAS);" 
		style="width:685px;" onkeypress="filtra_nome_identificador();"
		><%=r_OS.obs_pecas_necessarias%></textarea></td>
	</tr>
</table>



<!-- ************   SEPARADOR   ************ -->
<table width="749" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>
<br>


<table width="749" cellSpacing="0">
<tr>
	<td><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()" title="volta para página anterior">
		<img src="../botao/anterior.gif" width="176" height="55" border="0"></a></td>
	<td align="RIGHT"><div name="dCONFIRMA" id="dCONFIRMA"><a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fOPConfirma(fOP)" title="confirma a operação">
		<img src="../botao/confirmar.gif" width="176" height="55" border="0"></a></div>
	</td>
</tr>
</table>
</form>

</center>
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