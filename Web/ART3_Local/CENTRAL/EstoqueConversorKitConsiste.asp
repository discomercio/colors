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
'	  E S T O Q U E C O N V E R S O R K I T C O N S I S T E . A S P
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

	dim c_nfe_emitente
	dim s_kit_fabricante, s_kit, s_kit_descricao, s_kit_descricao_html, n_kit_qtde, s_documento, s_kit_nome_fabricante
	dim s_fabricante, s_produto, s_ean, s_descricao, s_descricao_html, s_qtde
	dim c_ncm, c_ncm_redigite, c_cst, c_cst_redigite
	dim v_item
	
'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	
'	OBTÉM DADOS DIGITADOS NO FORMULÁRIO
	c_nfe_emitente = Trim(Request.Form("c_nfe_emitente"))
	s_kit_fabricante = normaliza_codigo(retorna_so_digitos(Request.Form("c_kit_fabricante")), TAM_MIN_FABRICANTE)
	s_documento = Trim(Request.Form("c_documento"))
	s_kit = Ucase(Trim(Request.Form("c_kit")))
	s = Trim(Request.Form("c_kit_qtde"))
	if IsNumeric(s) then n_kit_qtde = CLng(s) else n_kit_qtde = 0
	c_ncm = retorna_so_digitos(Trim(Request.Form("c_ncm")))
	c_ncm_redigite = retorna_so_digitos(Trim(Request.Form("c_ncm_redigite")))
	c_cst = retorna_so_digitos(Trim(Request.Form("c_cst")))
	c_cst_redigite = retorna_so_digitos(Trim(Request.Form("c_cst_redigite")))
	
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

	if c_nfe_emitente = "" then
		alerta = "Não foi selecionada a empresa cujo estoque será processado."
	elseif converte_numero(c_nfe_emitente) = 0 then
		alerta = "É necessário informar a empresa cujo estoque será processado."
	elseif s_kit = "" then
		alerta = "O código de produto do kit a ser gerado não foi preenchido."
	elseif (Not IsEAN(s_kit)) And (s_kit_fabricante = "") then
		alerta = "O fabricante do kit a ser gerado não foi informado."
	elseif n_kit_qtde <= 0 then
		alerta = "Quantidade inválida de kits a serem gerados."
		end if

	if alerta = "" then
		if (Not IsEAN(s_kit)) And (s_kit_fabricante="") then
			alerta=texto_add_br(alerta)
			alerta=alerta & "Não foi especificado o fabricante do kit a ser gerado."
		else
			s = "SELECT * FROM t_PRODUTO WHERE"
			if IsEAN(s_kit) then
				s = s & " (ean='" & s_kit & "')"
			else
				s = s & " (fabricante='" & s_kit_fabricante & "') AND (produto='" & s_kit & "')"
				end if
			
			if rs.State <> 0 then rs.Close
			rs.open s, cn
			if rs.Eof then
				alerta=texto_add_br(alerta)
				alerta=alerta & "Código de produto do kit a ser gerado (produto " & s_kit & ", fabricante " & s_kit_fabricante & ") NÃO está cadastrado."
			else
				flag_ok = True
				if IsEAN(s_kit) And (s_kit_fabricante<>"") then
					if (s_kit_fabricante<>Trim("" & rs("fabricante"))) then
						flag_ok = False
						alerta=texto_add_br(alerta)
						alerta=alerta & "Kit a ser gerado " & s_kit & " NÃO pertence ao fabricante " & s_kit_fabricante & "."
						end if
					end if
				if flag_ok then
				'	CARREGA CÓDIGO INTERNO DO PRODUTO
					s_kit_fabricante = Trim("" & rs("fabricante"))
					s_kit = Trim("" & rs("produto"))
					s_kit_descricao = Trim("" & rs("descricao"))
					s_kit_descricao_html = Trim("" & rs("descricao_html"))
					end if
				end if
			end if
		end if
	
	if alerta = "" then
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
	'	VERIFICA CADA UM DOS PRODUTOS QUE COMPÕEM O KIT
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
					if (.produto = v_item(j).produto) then
						alerta=texto_add_br(alerta)
						alerta=alerta & "Produto " & .produto & ": linha " & renumera_com_base1(Lbound(v_item),i) & " repete o mesmo produto da linha " & renumera_com_base1(Lbound(v_item),j) & "."
						exit for
						end if
					next
				
				if IsEAN(s_kit) then
					if s_kit = .produto then
						alerta=texto_add_br(alerta)
						alerta=alerta & "O código de produto do kit a ser gerado não pode constar na relação de produtos usados em sua composição."
						end if
				else
					if (s_kit_fabricante=.fabricante) And (s_kit=.produto) then
						alerta=texto_add_br(alerta)
						alerta=alerta & "O código de produto do kit a ser gerado não pode constar na relação de produtos usados em sua composição."
						end if
					end if
				end with
			next
		end if

'	VERIFICA DISPONIBILIDADE NO ESTOQUE
	if alerta = "" then
		for i=Lbound(v_item) to Ubound(v_item)
			with v_item(i)
			'	QUANTIDADE DE PRODUTOS A SAIR DO ESTOQUE = QUANTIDADE DE KITS x QTDE DE PRODUTOS POR KIT
				n = n_kit_qtde * .qtde
				s = "SELECT" & _
						" SUM(qtde-qtde_utilizada) AS total" & _
					" FROM t_ESTOQUE tE INNER JOIN t_ESTOQUE_ITEM tEI ON (tE.id_estoque = tEI.id_estoque)" & _
					" WHERE" & _
						" (tE.id_nfe_emitente = " & c_nfe_emitente & ")" & _
						" AND (tEI.fabricante = '" & Trim(.fabricante) & "')" & _
						" AND (tEI.produto = '" & Trim(.produto) & "')" & _
						" AND ((qtde-qtde_utilizada) > 0)"
				if rs.State <> 0 then rs.Close
				rs.open s, cn
				j=0
				if Not rs.Eof then 
					if Not IsNull(rs("total")) then j = CLng(rs("total"))
					end if
				if n > j then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Faltam " & CStr(n-j) & " unidades no estoque do produto " & .produto & " do fabricante " & .fabricante & " (estoque: " & obtem_apelido_empresa_NFe_emitente(c_nfe_emitente) & ")."
					end if
				end with
			next
		end if

'	NCM / CST
	if alerta = "" then
		if c_ncm = "" then
			alerta=texto_add_br(alerta)
			alerta=alerta & "Informe o NCM do kit."
		elseif (Len(c_ncm) <> 2) And (Len(c_ncm) <> 8) then
			alerta=texto_add_br(alerta)
			alerta=alerta & "NCM possui tamanho inválido."
		elseif c_ncm <> c_ncm_redigite then
			alerta=texto_add_br(alerta)
			alerta=alerta & "Falha na conferência do NCM redigitado: '" & c_ncm & "' e '" & c_ncm_redigite & "'"
			end if
		
		if c_cst = "" then
			alerta=texto_add_br(alerta)
			alerta=alerta & "Informe o CST (entrada) do kit."
		elseif (Len(c_cst) <> 3) then
			alerta=texto_add_br(alerta)
			alerta=alerta & "CST possui tamanho inválido."
		elseif c_cst <> c_cst_redigite then
			alerta=texto_add_br(alerta)
			alerta=alerta & "Falha na conferência do CST redigitado: '" & c_cst & "' e '" & c_cst_redigite & "'"
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

<script language="JavaScript" type="text/javascript">
function fKITConfirma( f ) {
var b, s;
	s = "Executa a operação para gerar " + f.c_kit_qtde.value + " kits: " + f.c_kit_aux.value + "?";
	s = s + "\n\n";
	s = s + "Lembre-se: após os kits serem gerados, não será possível alterar os dados e nem reverter a operação!!";
	b=window.confirm(s);
	if (!b) return;
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

<form id="fKIT" name="fKIT" method="post" action="EstoqueConversorKitConfirma.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type=hidden name="c_kit" id="c_kit" value="<%=s_kit%>">
<input type=hidden name="c_kit_fabricante" id="c_kit_fabricante" value="<%=s_kit_fabricante%>">
<input type=hidden name="c_nfe_emitente" id="c_nfe_emitente" value="<%=c_nfe_emitente%>">

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">

<tr>
	<td align="right" valign="bottom"><p class="PEDIDO">Conversor para Cadastramento de Kits<span class="C">&nbsp;</span></p></td>
</tr>
</table>
<br>

<!--  COMVERSOR DE KITS  -->
<table class="Qx" cellspacing="0" cellpadding="0">
	<!--  TÍTULO  -->
	<tr bgcolor="#FFFFFF">
	<td><span style="width:30px;">&nbsp;</span></td>
	<td colspan="2" class="MT" valign="middle" align="center" NOWRAP style="background:azure;"><span class="PLTc" style="vertical-align:middle;"
		>KIT A SER GERADO</span></td>
	</tr>
<!--  EMPRESA  -->
	<tr bgcolor="#FFFFFF">
	<td>&nbsp;</td>
	<td colspan="2" class="MDBE" nowrap><span class="PLTe">Empresa</span>
		<br />
		<% s = obtem_apelido_empresa_NFe_emitente(c_nfe_emitente) %>
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
	<tr bgcolor="#FFFFFF">
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
	<tr bgColor="#FFFFFF">
		<td>&nbsp;</td>
		<td colspan="2" class="MDBE" nowrap><span class="PLTe">Documento</span>
			<br><input name="c_documento" id="c_documento" readonly tabindex=-1 class="PLLe" style="width:270px;margin-left:2pt;"
			value="<%=s_documento%>"></td>
	</tr>
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
<input type=hidden name="c_linha" value="">
<input type=hidden name="c_fabricante" value="">
<input type=hidden name="c_produto" value="">
<input type=hidden name="c_ean" value="">
<input type=hidden name="c_descricao" value="">
<input type=hidden name="c_qtde" value="">


<!-- ************   SEPARADOR   ************ -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>
<br>


<table width="649" cellspacing="0">
<tr>
	<td><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()" title="volta para página anterior">
		<img src="../botao/anterior.gif" width="176" height="55" border="0"></a></td>
	<td align="right"><div name="dCONFIRMA" id="dCONFIRMA"><a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fKITConfirma(fKIT)" title="confirma a conversão de kits">
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