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
'	  E S T O Q U E C O N V E R S O R K I T S U M A R I O . A S P
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

	dim s, usuario, msg_erro
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	if Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

	dim alerta
	alerta = ""
	
	dim v, v_aux, i, j, n, s_clipboard, s_id_nfe_emitente, s_kit_fabricante, s_kit, n_kit_qtde, s_documento
	dim c_ncm, c_cst
	dim s_vetor_composicao, s_vetor_sumario, s_sumario, s_data_entrada, s_usuario
	dim v_item, v_kit
	dim id_nfe_emitente
	s_clipboard = Session(SESSION_CLIPBOARD)
	if Trim(s_clipboard) = "" then
		alerta = "Não há informações adicionais."
	else
	'	DECODIFICA AS INFORMAÇÕES P/ EXIBIR O RESULTADO DA CONVERSÃO
		v = split(s_clipboard,chr(0),-1)
		i = Lbound(v)
		s_id_nfe_emitente = v(i)
		id_nfe_emitente = converte_numero(s_id_nfe_emitente)
		i = i + 1
		s_kit_fabricante = v(i)
		i = i + 1
		s_kit = v(i)
		i = i + 1
		s = v(i)
		if IsNumeric(s) then n_kit_qtde = CLng(s) else n_kit_qtde = 0
		i = i + 1
		c_ncm = v(i)
		i = i + 1
		c_cst = v(i)
		i = i + 1
		s_documento = v(i)
		i = i + 1
		s_vetor_composicao = v(i)
		i = i + 1
		s_vetor_sumario = v(i)
		i = i + 1
		s_sumario = v(i)
		
		redim v_item(0)
		set v_item(0) = New cl_ITEM_PEDIDO
		if s_vetor_composicao <> "" then
			v = split(s_vetor_composicao,chr(1),-1)
			for i = Lbound(v) to Ubound(v)
				if Trim(v_item(ubound(v_item)).produto) <> "" then
					redim preserve v_item(ubound(v_item)+1)
					set v_item(ubound(v_item)) = New cl_ITEM_PEDIDO
					end if
				if Trim(v(i))<>"" then
					with v_item(ubound(v_item))
						v_aux = split(v(i),chr(2),-1)
						j = Lbound(v_aux)
						.fabricante = v_aux(j)
						j = j + 1
						.produto = v_aux(j)
						j = j + 1
						s = v_aux(j)
						if IsNumeric(s) then .qtde = CLng(s) else .qtde = 0
						end with
					end if
				next
			end if
		
		redim v_kit(0)
		set v_kit(Ubound(v_kit)) = New cl_AGRUPA_KIT_POR_PRECO
		if s_vetor_sumario <> "" then
			v = split(s_vetor_sumario,chr(1),-1)
			for i = Lbound(v) to Ubound(v)
				if Trim(v_kit(ubound(v_kit)).id_estoque) <> "" then
					redim preserve v_kit(ubound(v_kit)+1)
					set v_kit(ubound(v_kit)) = New cl_AGRUPA_KIT_POR_PRECO
					end if
				if Trim(v(i))<>"" then
					with v_kit(ubound(v_kit))
						v_aux = split(v(i),chr(2),-1)
						j = Lbound(v_aux)
						.id_estoque = v_aux(j)
						j = j + 1
						s = v_aux(j)
						if IsNumeric(s) then .qtde = CLng(s) else .qtde = 0
						j = j + 1
						.preco_fabricante = converte_numero(v_aux(j))
						j = j + 1
						.vl_custo2 = converte_numero(v_aux(j))
						j = j + 1
						.vl_BC_ICMS_ST = converte_numero(v_aux(j))
						j = j + 1
						.vl_ICMS_ST = converte_numero(v_aux(j))
						end with
					end if
				next
			end if
		end if

	dim s_kit_nome_fabricante, s_kit_descricao, s_kit_descricao_html
	dim s_fabricante, s_produto, s_ean, s_descricao, s_descricao_html, s_qtde
	
	if alerta = "" then
	'	DESCRIÇÃO DO KIT
		s = "SELECT" & _
				" fabricante, produto, descricao, descricao_html" & _
			" FROM t_PRODUTO" & _
			" WHERE" & _
				" (fabricante='" & s_kit_fabricante & "')" & _
				" AND (produto='" & s_kit & "')"
		if rs.State <> 0 then rs.Close
		rs.open s, cn
		if rs.Eof then
			alerta = "O kit " & s_kit & " do fabricante " & s_kit_fabricante & " não está cadastrado."
		else
			s_kit_descricao = Trim("" & rs("descricao"))
			s_kit_descricao_html = Trim("" & rs("descricao_html"))
			end if
		end if
	
	if alerta = "" then
	'	DESCRIÇÃO DO FABRICANTE
		s = "SELECT nome, razao_social FROM t_FABRICANTE WHERE (fabricante='" & s_kit_fabricante & "')"
		if rs.State <> 0 then rs.Close
		rs.open s, cn
		s_kit_nome_fabricante = ""
		if rs.Eof then
			alerta = "Fabricante " & s_kit_fabricante & " não está cadastrado."
		else
			s_kit_nome_fabricante = Trim("" & rs("razao_social"))
			if s_kit_nome_fabricante = "" then s_kit_nome_fabricante = Trim("" & rs("nome"))
			end if
		end if

	if alerta = "" then
	'	OBTÉM DESCRIÇÃO E EAN DOS PRODUTOS USADOS NA COMPOSIÇÃO
		for i = Lbound(v_item) to Ubound(v_item)
			with v_item(i)
				if .produto <> "" then
					s = "SELECT" & _
							" fabricante, produto, ean, descricao, descricao_html" & _
						" FROM t_PRODUTO" & _
						" WHERE" & _
							" (fabricante='" & .fabricante & "')" & _
							" AND (produto='" & .produto & "')"
					if rs.State <> 0 then rs.Close
					rs.open s, cn
					if rs.Eof then
						.descricao = "NÃO CADASTRADO"
						.descricao_html = "NÃO CADASTRADO"
					else
						.descricao = Trim("" & rs("descricao"))
						.descricao_html = Trim("" & rs("descricao_html"))
						.ean = Trim("" & rs("ean"))
						end if
					end if
				end with
			next
		end if
	
	if alerta = "" then
	'	OBTÉM DATA/OPERADOR
	'	MESMO QUE TENHAM SIDO GERADOS VÁRIOS REGISTROS NO ESTOQUE (DEVIDO AOS DIFERENTES
	'	"PREÇO FABRICANTE" RESULTANTES DA CONVERSÃO), O OPERADOR É SEMPRE O MESMO E A
	'	DATA/HORA DA OPERAÇÃO VARIA MUITO POUCO.
		s = "SELECT " & _
				"*" & _
			" FROM t_ESTOQUE" & _
			" WHERE" & _
				" (id_estoque='" & v_kit(Ubound(v_kit)).id_estoque & "')"
		if rs.State <> 0 then rs.Close
		rs.open s, cn
		if Not rs.Eof then
			s = Trim("" & rs("hora_entrada"))
			if s <> "" then 
				s = formata_hhnnss_para_hh_nn_ss(s)
				s = " - " & s
				end if
			s_data_entrada = formata_data(rs("data_entrada"))
			s_data_entrada = s_data_entrada & s
			s_usuario = Trim("" & rs("usuario"))
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



<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>




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
<link href="<%=URL_FILE__EPRINTER_CSS%>" rel="stylesheet" type="text/css" media="print">


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
	<td align="center"><a name="bVOLTAR" id="bVOLTAR" href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>"><img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
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

<form id="fKIT" name="fKIT" method="post" action="EstoqueConversorKitSumario.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type=hidden name="c_kit" id="c_kit" value="<%=s_kit%>">
<input type=hidden name="c_kit_fabricante" id="c_kit_fabricante" value="<%=s_kit_fabricante%>">
<input type=hidden name="c_nfe_emitente" id="c_nfe_emitente" value="<%=s_id_nfe_emitente%>">

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">

<tr>
	<td align="right" valign="bottom"><p class="PEDIDO">Sumário da Conversão de Kits<span class="C">&nbsp;</span></p></td>
</tr>
</table>
<br>

<!--  COMVERSOR DE KITS  -->
<table class="Qx" cellspacing="0">
	<!--  TÍTULO  -->
	<tr bgcolor="#FFFFFF">
	<td><span style="width:30px;">&nbsp;</span></td>
	<td colspan="2" class="MT" valign="middle" align="center" nowrap style="background:azure;"><span class="PLTc" style="vertical-align:middle;"
		>KIT GERADO</span></td>
	</tr>
<!--  EMPRESA  -->
	<tr bgcolor="#FFFFFF">
	<td>&nbsp;</td>
	<td colspan="2" class="MDBE" nowrap><span class="PLTe">Empresa</span>
		<br />
		<% s = obtem_apelido_empresa_NFe_emitente(id_nfe_emitente) %>
		<input name="c_nfe_emitente_aux" id="c_nfe_emitente_aux" readonly tabindex=-1 class="PLLe" style="width:460px;margin-left:2pt;" 
				value="<%=s%>">
	</td>
	</tr>
<!--  FABRICANTE  -->
	<tr bgcolor="#FFFFFF">
	<td>&nbsp;</td>
	<td colspan="2" class="MDBE" nowrap><span class="PLTe">Fabricante</span>
		<%	s = s_kit_fabricante
			if (s<>"") And (s_kit_nome_fabricante<>"") then s = s & " - " & s_kit_nome_fabricante %>
		<br><input name="c_kit_fabricante_aux" id="c_kit_fabricante_aux" readonly tabindex=-1 class="PLLe" style="width:460px;margin-left:2pt;" 
				value="<%=s%>"></td></tr>
<!--  KIT  -->
	<tr bgcolor="#FFFFFF">
	<td>&nbsp;</td>
	<td colspan="2" class="MDBE" nowrap><span class="PLTe">Kit</span>
		<%	s = s_kit
			if (s<>"") And (s_kit_descricao_html<>"") then s = s & " - " & produto_formata_descricao_em_html(s_kit_descricao_html) %>
		<br><span class="PLLe" style="width:460px;margin-left:2pt;"><%=s%></span>
		<%	s = s_kit
			if (s<>"") And (s_kit_descricao<>"") then s = s & " - " & s_kit_descricao %>
		<input type=hidden name="c_kit_aux" id="c_kit_aux" value="<%=s%>">
	</td>
	</tr>

<!--  QTDE/NCM/CST  -->
	<tr bgColor="#FFFFFF">
		<td>&nbsp;</td>
		<td colspan="2" class="MDBE">
			<table width="100%" cellpadding=0 cellspacing=0>
				<tr>
				<td width="30%" class="MD" nowrap><span class="PLTe" style="margin-right:2pt;">Qtde</span>
					<br><input name="c_kit_qtde" id="c_kit_qtde" readonly tabindex=-1 class="PLLe" style="width:35px;margin-left:2pt;"
					value="<%=Cstr(n_kit_qtde)%>">
				<td width="35%" class="MD" nowrap><span class="PLTe">NCM</span>
					<br><input name="c_ncm" id="c_ncm" readonly tabindex=-1 class="PLLe" style="margin-left:2pt;width:80px;"
					value="<%=c_ncm%>"></td>
				<td nowrap><span class="PLTe">CST (entrada)</span>
					<br><input name="c_cst" id="c_cst" readonly tabindex=-1 class="PLLe" style="margin-left:2pt;width:80px;"
					value="<%=c_cst%>"></td>
				</tr>
			</table>
		</td>
	</tr>

<!--  DOCUMENTO  -->
	<tr bgcolor="#FFFFFF">
		<td>&nbsp;</td>
		<td colspan="2" class="MDBE" nowrap><span class="PLTe">Documento</span>
			<br><input name="c_documento" id="c_documento" readonly tabindex=-1 class="PLLe" style="width:270px;margin-left:2pt;"
			value="<%=s_documento%>"></td></tr>

<!--  DATA/OPERADOR  -->
	<tr bgcolor="#FFFFFF">
		<td>&nbsp;</td>
		<td class="MDBE" nowrap><span class="PLTe" style="margin-right:2pt;">Data do Cadastramento</span>
			<br><input name="c_data_entrada" id="c_data_entrada" readonly tabindex=-1 class="PLLe" style="width:120px;margin-left:2pt;"
			value="<%=s_data_entrada%>">
		<td class="MDB" nowrap><span class="PLTe">Cadastrado por</span>
			<br><input name="c_usuario" id="c_usuario" readonly tabindex=-1 class="PLLe" style="width:120px;margin-left:2pt;"
			value="<%=s_usuario%>"></td></tr>
</table>
<br><br>

<!--  R E L A Ç Ã O   D E   P R O D U T O S  -->
<table class="Qx" cellspacing="0">
	<!--  TÍTULO DA TABELA  -->
	<tr bgcolor="#FFFFFF">
	<td>&nbsp;</td>
	<td colspan="5" class="MT" valign="middle" align="center" nowrap style="background:azure;"><span class="PLTc" style="vertical-align:middle;"
		>COMPOSIÇÃO DE 1 UNIDADE DO KIT</span></td>
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
	for i=1 to MAX_PRODUTOS_CONVERSOR_KIT
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
	<td><input name="c_linha" readonly tabindex=-1 class="PLLe" maxlength="2" style="width:30px;text-align:right;color:#808080;" value="<%=Cstr(i) & ". " %>"></td>
	<td class="MDBE">
		<input name="c_fabricante" readonly tabindex=-1 class="PLLe" style="width:30px;"
			value="<%=s_fabricante%>"></td>
	<td class="MDB">
		<input name="c_produto" readonly tabindex=-1 class="PLLe" style="width:55px;"
			value="<%=s_produto%>"></td>
	<td class="MDB">
		<input name="c_ean" readonly tabindex=-1 class="PLLe" style="width:85px;"
			value="<%=s_ean%>"></td>
	<td class="MDB" style="width:277px;">
		<span class="PLLe"><%=s_descricao_html%></span>
		<input type=hidden name="c_descricao" value="<%=s_descricao%>">
	</td>
	<td class="MDB" align="right">
		<input name="c_qtde" readonly tabindex=-1 class="PLLd" style="width:35px;"
			value="<%=s_qtde%>"></td>
	</tr>
<% next %>
</table>


<!--  ASSEGURA CRIAÇÃO DE UM ARRAY DE CAMPOS, MESMO QUANDO HOUVER SOMENTE 1 PRODUTO!! -->
<input type=HIDDEN name="c_linha" value="">
<input type=HIDDEN name="c_fabricante" value="">
<input type=HIDDEN name="c_produto" value="">
<input type=HIDDEN name="c_ean" value="">
<input type=HIDDEN name="c_descricao" value="">
<input type=HIDDEN name="c_qtde" value="">


<!-- ************   SUMÁRIO   ************ -->
<br><br>
<table class="Qx" cellspacing="0" cellpadding="0">
	<!--  TÍTULO DA TABELA  -->
	<tr bgcolor="#FFFFFF">
	<td class="MT" valign="middle" align="center" nowrap style="background:azure;"><span class="PLTc" style="vertical-align:middle;"
		>SUMÁRIO DOS KITS GERADOS</span></td>
	</tr>
	<!--  TÍTULO DAS COLUNAS  -->
	<tr bgcolor="#FFFFFF">
	<td class="MDBE">
		<textarea readonly name="c_sumario" id="c_sumario" class="TA" style="width:800px;height:400px;font-family:Courier New,Arial,Helvetica,sans-serif;font-size:8pt;font-weight:bold;margin:0px 0px 0px 4px;"><%=s_sumario%></textarea>
		</td>
	</tr>
</table>

<!-- ************   SEPARADOR   ************ -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>
<br>


<table class="notPrint" width="649" cellSpacing="0">
<tr>
	<td align="center"><a name="bVOLTAR" id="bVOLTAR" href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
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