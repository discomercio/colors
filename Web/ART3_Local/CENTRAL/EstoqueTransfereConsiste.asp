<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/estoque.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =============================================================
'	  E S T O Q U E T R A N S F E R E C O N S I S T E . A S P
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

	dim s, usuario, i, n, j, flag_ok, msg_erro
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

	dim s_id_nfe_emitente
	dim s_tipo, s_loja, s_nome_loja, s_pedido, s_fabricante, s_produto, s_ean
	dim s_descricao, s_descricao_html, s_qtde, s_op_descricao
	dim s_fluxo, s_cod_estoque, s_ckb_spe, s_ckb_spe_descricao
	dim v_aux, s_cod_estoque_origem, s_cod_estoque_destino
	dim v_item
	
'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	
'	OBTÉM DADOS DO FORMULÁRIO
	s_id_nfe_emitente = Trim(Request.Form("c_id_nfe_emitente"))
	s_tipo = Ucase(Trim(Request.Form("rb_tipo")))
	s_loja = Trim(Request.Form("c_loja"))
	s_pedido = Trim(Request.Form("c_pedido"))
	s_op_descricao = Trim(Request.Form("op_selecionada_descricao"))
	s_ckb_spe = Ucase(Trim(Request.Form("ckb_spe")))
	s_ckb_spe_descricao = Trim(Request.Form("ckb_spe_descricao"))
	
	s_cod_estoque = ""
	s_cod_estoque_origem = ""
	s_cod_estoque_destino = ""
	
	if InStr(s_tipo, "TRANSF_") > 0 then
		v_aux = Split(s_tipo, "_")
		s_cod_estoque_origem = v_aux(Ubound(v_aux)-1)
		s_cod_estoque_destino = v_aux(Ubound(v_aux))
		s_fluxo = "TRANSF"
		if (s_cod_estoque_origem<>ID_ESTOQUE_SHOW_ROOM)And(s_cod_estoque_origem<>ID_ESTOQUE_DEVOLUCAO)And(s_cod_estoque_destino<>ID_ESTOQUE_SHOW_ROOM)And(s_cod_estoque_destino<>ID_ESTOQUE_DEVOLUCAO) then s_loja=""
	else
		s_cod_estoque = Right(s_tipo, 3)
		s_fluxo = Left(s_tipo, 3)
		if (s_cod_estoque<>ID_ESTOQUE_SHOW_ROOM)And(s_cod_estoque<>ID_ESTOQUE_DEVOLUCAO) then s_loja=""
		end if
	
	redim v_item(0)
	set v_item(0) = New cl_ITEM_PEDIDO
	n = Request.Form("c_codigo").Count
	for i = 1 to n
		s=Trim(Request.Form("c_codigo")(i))
		if s <> "" then
			if Trim(v_item(ubound(v_item)).produto) <> "" then
				redim preserve v_item(ubound(v_item)+1)
				set v_item(ubound(v_item)) = New cl_ITEM_PEDIDO
				end if
			with v_item(ubound(v_item))
				s = retorna_so_digitos(Request.Form("c_fabricante")(i))
				s = normaliza_codigo(s, TAM_MIN_FABRICANTE)
				.fabricante = s
				.produto=Ucase(Trim(Request.Form("c_codigo")(i)))
				s = Trim(Request.Form("c_qtde")(i))
				if IsNumeric(s) then .qtde = CLng(s) else .qtde = 0
				end with
			end if
		next
	
'	CONSISTE DADOS DIGITADOS
	dim alerta
	alerta=""

	if s_id_nfe_emitente = "" then
		alerta=texto_add_br(alerta)
		alerta=alerta & "Não foi informada a empresa."
		end if

	if s_tipo = "" then
		alerta=texto_add_br(alerta)
		alerta=alerta & "Não foi indicado o tipo de transferência/movimentação do estoque a ser efetuado."
	elseif s_fluxo = "" then
		alerta=texto_add_br(alerta)
		alerta=alerta & "Operação selecionada é inválida."
		end if

	if s_fluxo = "TRANSF" then
		if (s_loja="") And ((s_cod_estoque_origem=ID_ESTOQUE_SHOW_ROOM)Or(s_cod_estoque_origem=ID_ESTOQUE_DEVOLUCAO)Or(s_cod_estoque_destino=ID_ESTOQUE_SHOW_ROOM)Or(s_cod_estoque_destino=ID_ESTOQUE_DEVOLUCAO)) then
			alerta=texto_add_br(alerta)
			alerta=alerta & "Número da loja não especificado."
		elseif (s_cod_estoque_origem="") Or (s_cod_estoque_destino="") then
			alerta=texto_add_br(alerta)
			alerta=alerta & "Operação selecionada é inválida."
			end if
	else
		if (s_loja="") And ((s_cod_estoque=ID_ESTOQUE_SHOW_ROOM)Or(s_cod_estoque=ID_ESTOQUE_DEVOLUCAO)) then
			alerta=texto_add_br(alerta)
			alerta=alerta & "Número da loja não especificado."
		elseif (s_fluxo<>"SAI") And (s_fluxo<>"ENT") then
			alerta=texto_add_br(alerta)
			alerta=alerta & "Operação selecionada é inválida."
		elseif (s_cod_estoque="") then
			alerta=texto_add_br(alerta)
			alerta=alerta & "Operação selecionada é inválida."
			end if
		end if

	if alerta = "" then
	'	VERIFICA A LOJA
		if s_loja <> "" then
			s = "SELECT * FROM t_LOJA WHERE" & _
				" (loja='" & s_loja & "')"
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
	
	if alerta = "" then
	'	VERIFICA CADA UM DOS PRODUTOS A SEREM TRANSFERIDOS
		for i=Lbound(v_item) to Ubound(v_item)
			with v_item(i)
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

				for j=Lbound(v_item) to (i-1)
					if (.fabricante=v_item(j).fabricante) And (.produto=v_item(j).produto) then
						alerta=texto_add_br(alerta)
						alerta=alerta & "Produto " & .produto & ": linha " & renumera_com_base1(Lbound(v_item),i) & " repete o mesmo produto da linha " & renumera_com_base1(Lbound(v_item),j) & "."
						exit for
						end if
					next
				end with
			next
		end if


'	ESTOQUE DE DEVOLUÇÃO
	Dim blnEstoqueDevolucao
	blnEstoqueDevolucao=False
	if alerta = "" then
		if (s_tipo="ENT_DEV") then blnEstoqueDevolucao=True
		if (s_tipo="TRANSF_DEV_DAN") then blnEstoqueDevolucao=True
		if (s_tipo="TRANSF_DEV_ROU") then blnEstoqueDevolucao=True
		if blnEstoqueDevolucao then
			if s_pedido = "" then
				alerta = "Número do pedido não informado."
			else
				s = "SELECT " & _
						"*" & _
					" FROM t_PEDIDO" & _
					" WHERE" & _
						" (pedido='" & s_pedido & "')"
				if rs.State <> 0 then rs.Close
				rs.open s, cn
				if rs.Eof then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Pedido " & s_pedido & " NÃO está cadastrado."
					end if
				end if
			
			if alerta = "" then
				if converte_numero(rs("loja")) <> converte_numero(s_loja) then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Pedido " & s_pedido & " NÃO pertence à loja " & s_loja & "."
					end if
				if converte_numero(rs("id_nfe_emitente")) <> converte_numero(s_id_nfe_emitente) then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Pedido " & s_pedido & " NÃO está vinculado ao CD '" & obtem_apelido_empresa_NFe_emitente(s_id_nfe_emitente) & "'"
					end if
				end if
			
			if alerta = "" then
				for i=Lbound(v_item) to Ubound(v_item)
					with v_item(i)
					'	QUANTIDADE DE PRODUTOS A SER TRANSFERIDA
						s = "SELECT" & _
								" SUM(qtde) AS total" & _
							" FROM t_ESTOQUE_MOVIMENTO" & _
							" WHERE" & _
								" (anulado_status=0)" & _
								" AND (fabricante='" & Trim(.fabricante) & "')" & _
								" AND (produto='" & Trim(.produto) & "')" & _
								" AND (estoque='" & ID_ESTOQUE_DEVOLUCAO & "')" & _
								" AND (loja='" & s_loja & "')" & _
								" AND (pedido='" & s_pedido & "')"
						if rs.State <> 0 then rs.Close
						rs.open s, cn
						j=0
						if Not rs.Eof then 
							if Not IsNull(rs("total")) then j = CLng(rs("total"))
							end if
						if .qtde > j then
							alerta=texto_add_br(alerta)
							alerta=alerta & "Pedido " & s_pedido & ": faltam " & CStr(.qtde-j) & " unidades do produto " & .produto & " do fabricante " & .fabricante & "."
							end if
						end with
					next
				end if
			end if
		end if


'	VERIFICA DISPONIBILIDADE NO ESTOQUE
	if alerta = "" then
		for i=Lbound(v_item) to Ubound(v_item)
			with v_item(i)
			'	QUANTIDADE DE PRODUTOS A SER TRANSFERIDA
			'	OBS: RESSALTANDO QUE, NA TABELA T_ESTOQUE_MOVIMENTO, SOMENTE O ESTOQUE LÓGICO 'SPE' (SEM PRESENÇA NO ESTOQUE) NÃO POSSUI CONTEÚDO NO CAMPO 'id_estoque'.
				s = ""
				if s_fluxo="SAI" then
					s = "SELECT" & _
							" SUM(qtde-qtde_utilizada) AS total" & _
						" FROM t_ESTOQUE tE" & _
							" INNER JOIN t_ESTOQUE_ITEM tEI ON (tE.id_estoque = tEI.id_estoque)" & _
						" WHERE" & _
							" (id_nfe_emitente = " & s_id_nfe_emitente & ")" & _
							" AND (tEI.fabricante = '" & Trim(.fabricante) & "')" & _
							" AND (tEI.produto = '" & Trim(.produto) & "')" & _
							" AND ((qtde-qtde_utilizada) > 0)"
				elseif s_fluxo="ENT" then
					s = "SELECT" & _
							" SUM(qtde) AS total" & _
						" FROM t_ESTOQUE_MOVIMENTO tEM" & _
							" INNER JOIN t_ESTOQUE tE ON (tEM.id_estoque = tE.id_estoque)" & _
						" WHERE" & _
							" (tEM.anulado_status=0)" & _
							" AND (tE.id_nfe_emitente = " & s_id_nfe_emitente & ")" & _
							" AND (tEM.fabricante='" & Trim(.fabricante) & "')" & _
							" AND (tEM.produto='" & Trim(.produto) & "')" & _
							" AND (tEM.estoque='" & s_cod_estoque & "')"
					if (s_cod_estoque=ID_ESTOQUE_SHOW_ROOM)Or(s_cod_estoque=ID_ESTOQUE_DEVOLUCAO) then
						s = s & " AND (tEM.loja='" & s_loja & "')"
						end if
				elseif s_fluxo="TRANSF" then
					s = "SELECT" & _
							" SUM(qtde) AS total" & _
						" FROM t_ESTOQUE_MOVIMENTO tEM" & _
							" INNER JOIN t_ESTOQUE tE ON (tEM.id_estoque = tE.id_estoque)" & _
						" WHERE" & _
							" (tEM.anulado_status=0)" & _
							" AND (tE.id_nfe_emitente = " & s_id_nfe_emitente & ")" & _
							" AND (tEM.fabricante='" & Trim(.fabricante) & "')" & _
							" AND (tEM.produto='" & Trim(.produto) & "')" & _
							" AND (tEM.estoque='" & s_cod_estoque_origem & "')"
					if (s_cod_estoque_origem=ID_ESTOQUE_SHOW_ROOM)Or(s_cod_estoque_origem=ID_ESTOQUE_DEVOLUCAO) then
						s = s & " AND (tEM.loja='" & s_loja & "')"
						end if
					end if
				
				if rs.State <> 0 then rs.Close
				rs.open s, cn
				j=0
				if Not rs.Eof then 
					if Not IsNull(rs("total")) then j = CLng(rs("total"))
					end if
				if .qtde > j then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Faltam " & CStr(.qtde-j) & " unidades do produto " & .produto & " do fabricante " & .fabricante & " (CD '" & obtem_apelido_empresa_NFe_emitente(s_id_nfe_emitente) & "')"
					end if
				end with
			next
		end if

	if alerta <> "" then 
		alerta = texto_add_br(s_op_descricao) & alerta
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



<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>

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

<link href="<%=URL_FILE__E_CSS%>" rel="stylesheet" type="text/css">


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
	<td align="center"><a name="bVOLTAR" href="javascript:history.back()"><img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
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

<form id="fOP" name="fOP" method="post" action="EstoqueTransfereConfirma.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type=hidden name="c_loja" id="c_loja" value="<%=s_loja%>">
<input type=hidden name="c_pedido" id="c_pedido" value="<%=s_pedido%>">
<input type=hidden name="rb_tipo" id="rb_tipo" value="<%=s_tipo%>">
<input type=hidden name="op_selecionada_descricao" id="op_selecionada_descricao" value="<%=s_op_descricao%>">
<input type=hidden name="ckb_spe" id="ckb_spe" value="<%=s_ckb_spe%>">
<input type=hidden name="ckb_spe_descricao" id="ckb_spe_descricao" value="<%=s_ckb_spe_descricao%>">
<input type="hidden" name="c_id_nfe_emitente" id="c_id_nfe_emitente" value="<%=s_id_nfe_emitente%>" />

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">

<tr>
	<td align="right" valign="bottom"><p class="PEDIDO">Transferência/Movimentação do Estoque<span class="C">&nbsp;</span></p></td>
</tr>
</table>
<br>

<!--  TRANSFERÊNCIA ENTRE ESTOQUES  -->
<table class="Qx" cellspacing="0">
	<!--  EMPRESA  -->
	<tr bgcolor="#FFFFFF">
	<td><span style="width:30px;">&nbsp;</span></td>
	<td class="MT" valign="middle" align="center" nowrap style="background:azure;"><span class="PLTc" style="vertical-align:middle;"
		>EMPRESA</span></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td>&nbsp;</td>
		<td class="MDBE" nowrap><input name="c_id_nfe_emitente_aux" id="c_id_nfe_emitente_aux" tabindex="-1" class="PLLe" style="width:460px;margin-left:2pt;" value="<%=obtem_apelido_empresa_NFe_emitente(s_id_nfe_emitente)%>" /></td>
	</tr>

	<!--  PULA LINHA  -->
	<tr bgcolor="#FFFFFF">
	<td><span style="width:30px;">&nbsp;</span></td>
	<td>&nbsp;</td>
	</tr>

	<!--  TÍTULO  -->
	<tr bgcolor="#FFFFFF">
	<td><span style="width:30px;">&nbsp;</span></td>
	<td class="MT" valign="middle" align="center" nowrap style="background:azure;"><span class="PLTc" style="vertical-align:middle;"
		>TRANSFERÊNCIA/MOVIMENTAÇÃO DO ESTOQUE</span></td>
	</tr>
	<!--  OPÇÕES  -->
	<tr bgcolor="#FFFFFF">
	<td>&nbsp;</td>
	<td class="MDBE" nowrap>
		<span class="PLTe">Operação</span>
		<br>
		<span class="C" style="cursor:default;margin-left:3pt;"><%=s_op_descricao%></span>
		<%if s_fluxo="ENT" then%>
			<br><input type="checkbox" tabindex="-1" id="ckb_spe_aux" name="ckb_spe_aux" value="SPE_ON" disabled
				<%if s_ckb_spe<>"" then Response.Write " checked"%>
			><span class="C" id="sckb_spe" name="sckb_spe" style="cursor:default;font-weight:normal;font-style:normal;" 
				><%=s_ckb_spe_descricao%></span>
		<%end if%>
		</td>
	</tr>
<!--  LOJA  -->
	<tr bgcolor="#FFFFFF">
	<td>&nbsp;</td>
	<td class="MDBE" nowrap><span class="PLTe">Loja</span>
		<%	s = s_loja
			if (s<>"") And (s_nome_loja<>"") then s = s & " - " & s_nome_loja %>
		<br><input name="c_loja_aux" id="c_loja_aux" readonly tabindex=-1 class="PLLe" style="width:460px;margin-left:2pt;" 
				value="<%=s%>"></td></tr>
<% if blnEstoqueDevolucao then %>
<!--  PEDIDO  -->
	<tr bgcolor="#FFFFFF">
	<td>&nbsp;</td>
	<td class="MDBE" nowrap><span class="PLTe">Pedido</span>
		<br><input name="c_pedido_aux" id="c_pedido_aux" readonly tabindex=-1 class="PLLe" style="width:460px;margin-left:2pt;" 
				value="<%=s_pedido%>"></td></tr>
<% end if %>
</table>
<br><br>

<!--  R E L A Ç Ã O   D E   P R O D U T O S  -->
<table class="Qx" cellspacing="0">
	<!--  TÍTULO DA TABELA  -->
	<tr bgcolor="#FFFFFF">
	<td>&nbsp;</td>
	<td colspan="5" class="MT" valign="middle" align="center" nowrap style="background:azure;"><span class="PLTc" style="vertical-align:middle;"
		>PRODUTOS A MOVIMENTAR</span></td>
	</tr>
	<!--  TÍTULO DAS COLUNAS  -->
	<tr bgcolor="#FFFFFF">
	<td>&nbsp;</td>
	<td class="MDBE"><p class="PLTe">Fabr</p></td>
	<td class="MDB"><p class="PLTe">Produto</p></td>
	<td class="MDB"><p class="PLTe">EAN</p></td>
	<td class="MDB"><p class="PLTe">Descrição</p></td>
	<td class="MDB"><p class="PLTd">Qtde</p></td>
	</tr>

<%	n = Lbound(v_item)-1
	for i=1 to MAX_PRODUTOS_TRANSFERENCIA_ESTOQUE
		n = n+1
		if n <= Ubound(v_item) then
			with v_item(n)
				s_fabricante = .fabricante
				s_produto = .produto
				s_ean = .ean
				s_descricao = .descricao
				s_descricao_html = produto_formata_descricao_em_html(.descricao_html)
				s_qtde = Cstr(.qtde)
				end with
		else
			exit for
			end if
%>
	<tr>
	<td><input name="c_linha" id="c_linha" readonly tabindex=-1 class="PLLe" maxlength="2" style="width:30px;text-align:right;color:#808080;" value="<%=Cstr(i) & ". " %>"></td>
	<td class="MDBE">
		<input name="c_fabricante" id="c_fabricante" readonly tabindex=-1 class="PLLe" style="width:30px;"
			value="<%=s_fabricante%>">
	</td>
	<td class="MDB">
		<input name="c_produto" id="c_produto" readonly tabindex=-1 class="PLLe" style="width:55px;"
			value="<%=s_produto%>">
	</td>
	<td class="MDB">
		<input name="c_ean" id="c_ean" readonly tabindex=-1 class="PLLe" style="width:85px;"
			value="<%=s_ean%>">
	</td>
	<td class="MDB" style="width:277px;">
		<span class="PLLe"><%=s_descricao_html%></span>
		<input type=hidden name="c_descricao" id="c_descricao" value="<%=s_descricao%>">
	</td>
	<td class="MDB" align="right">
		<input name="c_qtde" id="c_qtde" readonly tabindex=-1 class="PLLd" style="width:35px;"
			value="<%=s_qtde%>">
	</td>
	</tr>
<% next %>
</table>


<!--  ASSEGURA CRIAÇÃO DE UM ARRAY DE CAMPOS, MESMO QUANDO HOUVER SOMENTE 1 PRODUTO!! -->
<input type=hidden name="c_linha" id="c_linha" value="">
<input type=hidden name="c_fabricante" id="c_fabricante" value="">
<input type=hidden name="c_produto" id="c_produto" value="">
<input type=hidden name="c_ean" id="c_ean" value="">
<input type=hidden name="c_descricao" id="c_descricao" value="">
<input type=hidden name="c_qtde" id="c_qtde" value="">


<!-- ************   SEPARADOR   ************ -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>
<br>


<table width="649" cellspacing="0">
<tr>
	<td><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()" title="volta para página anterior">
		<img src="../botao/anterior.gif" width="176" height="55" border="0"></a></td>
	<td align="right"><div name="dCONFIRMA" id="dCONFIRMA"><a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fOPConfirma(fOP)" title="confirma a transferência entre estoques">
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