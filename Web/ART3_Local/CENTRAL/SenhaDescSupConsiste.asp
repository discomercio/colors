<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =============================================================
'	  S E N H A D E S C S U P C O N S I S T E . A S P
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

	dim s, s_aux, usuario
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

	dim cliente_selecionado
	cliente_selecionado=Trim(request("cliente_selecionado"))
	if cliente_selecionado = "" then Response.Redirect("aviso.asp?id=" & ERR_CLIENTE_NAO_ESPECIFICADO)
	
'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs, msg_erro
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	
'	OBTÉM DADOS DO FORMULÁRIO
	dim c_supervisor_autorizador, c_loja
	dim s_nome_loja

	c_supervisor_autorizador = Trim(Request.Form("c_supervisor_autorizador"))
	c_loja = retorna_so_digitos(Trim(Request.Form("c_loja")))
	s = normaliza_codigo(c_loja, TAM_MIN_LOJA)
	if s <> "" then c_loja = s

	dim intCounter, intCounterAux, intQtdeItens
	dim v_item
	redim v_item(0)
	set v_item(0) = New cl_ITEM_SENHA_DESCONTO
	intQtdeItens = Request.Form("c_produto").Count
	for intCounter = 1 to intQtdeItens
		s = Trim(Request.Form("c_produto")(intCounter))
		if s <> "" then
			if Trim(v_item(Ubound(v_item)).produto) <> "" then
				redim preserve v_item(Ubound(v_item)+1)
				set v_item(Ubound(v_item)) = New cl_ITEM_SENHA_DESCONTO
				end if
			with v_item(Ubound(v_item))
				.produto = UCase(Trim(Request.Form("c_produto")(intCounter)))
				s = retorna_so_digitos(Request.Form("c_fabricante")(intCounter))
				.fabricante = normaliza_codigo(s, TAM_MIN_FABRICANTE)
				s = Trim(Request.Form("c_desc_max_senha")(intCounter))
				.perc_desconto = converte_numero(s)
				end with
			end if
		next
		
'	CONSISTE DADOS DIGITADOS
	dim blnTemItem
	dim alerta
	alerta=""

	if c_loja = "" then
		alerta = "Não foi especificada a loja."
	elseif c_supervisor_autorizador = "" then
		alerta = "Não foi informado quem está autorizando o desconto."
		end if

'	VERIFICA SE HÁ PRODUTOS REPETIDOS
	if alerta = "" then
		for intCounter = Lbound(v_item) to Ubound(v_item)
			with v_item(intCounter)
				for intCounterAux=Lbound(v_item) to (intCounter-1)
					if (.produto = v_item(intCounterAux).produto) And (.fabricante = v_item(intCounterAux).fabricante) then
						alerta=texto_add_br(alerta)
						alerta=alerta & "Produto " & .produto & " do fabricante " & .fabricante & ": linha " & renumera_com_base1(Lbound(v_item),intCounter) & " repete o mesmo produto da linha " & renumera_com_base1(Lbound(v_item),intCounterAux) & "."
						exit for
						end if
					next
				end with
			next
		end if
		
	if alerta = "" then
	'	VERIFICA A LOJA
		s = "SELECT " & _
				"*" & _
			" FROM t_LOJA" & _
			" WHERE" & _
				" (loja='" & c_loja & "')"
		if rs.State <> 0 then rs.Close
		rs.open s, cn
		if rs.Eof then
			alerta=texto_add_br(alerta)
			alerta=alerta & "Loja " & c_loja & " NÃO está cadastrada."
		else
			s_nome_loja = Trim("" & rs("nome"))
			if s_nome_loja = "" then s_nome_loja = Trim("" & rs("razao_social"))
			end if
		end if


	for intCounter = Lbound(v_item) to Ubound(v_item)
		if alerta = "" then
			with v_item(intCounter)
				blnTemItem = False
				if Trim(.fabricante) <> "" then 
					blnTemItem = True
				elseif Trim(.produto) <> "" then
					blnTemItem = True
				elseif .perc_desconto > 0 then
					blnTemItem = True
					end if

				if blnTemItem then
					if .fabricante = "" then
						alerta=texto_add_br(alerta)
						alerta=alerta & "Não foi especificado o código do fabricante."
					elseif .produto = "" then
						alerta=texto_add_br(alerta)
						alerta=alerta & "Não foi especificado o código do produto."
					elseif (.perc_desconto <= 0) Or (.perc_desconto > 100) then
						alerta=texto_add_br(alerta)
						alerta=alerta & "Percentual de desconto inválido."
						end if

					if alerta = "" then
					'	VERIFICA O FABRICANTE
						s = "SELECT " & _
								"*" & _
							" FROM t_FABRICANTE" & _
							" WHERE" & _
								" (fabricante='" & .fabricante & "')"
						if rs.State <> 0 then rs.Close
						rs.open s, cn
						if rs.Eof then
							alerta=texto_add_br(alerta)
							alerta=alerta & "Fabricante " & .fabricante & " NÃO está cadastrado."
						else
							.nome_fabricante = Trim("" & rs("razao_social"))
							if .nome_fabricante = "" then .nome_fabricante = Trim("" & rs("nome"))
							end if
						end if
						
					if alerta = "" then
					'	VERIFICA O PRODUTO (TABELA BÁSICA)
						s = "SELECT " & _
								"*" & _
							" FROM t_PRODUTO" & _
							" WHERE" & _
								" (fabricante='" & .fabricante & "')" & _
								" AND (produto='" & .produto & "')"
						if rs.State <> 0 then rs.Close
						rs.open s, cn
						if rs.Eof then
							alerta=texto_add_br(alerta)
							alerta=alerta & "Produto " & .produto & " do fabricante " & .fabricante & " NÃO está cadastrado."
						else
							.nome_produto = produto_formata_descricao_em_html(Trim("" & rs("descricao_html")))
							end if
						end if

					if alerta = "" then
					'	VERIFICA O PRODUTO (TABELA LOJA)
						s = "SELECT " & _
								"*" & _
							" FROM t_PRODUTO_LOJA" & _
							" WHERE" & _
								" (fabricante='" & .fabricante & "')" & _
								" AND (produto='" & .produto & "')" & _
								" AND (loja='" & c_loja & "')"
						if rs.State <> 0 then rs.Close
						rs.open s, cn
						if rs.Eof then
							alerta=texto_add_br(alerta)
							alerta=alerta & "Produto " & .produto & " do fabricante " & .fabricante & " NÃO está cadastrado para a loja " & c_loja & "."
						else
							if Ucase(Trim("" & rs("vendavel"))) <> "S" then
								alerta=texto_add_br(alerta)
								alerta=alerta & "Produto " & .produto & " do fabricante " & .fabricante & " NÃO está disponível para venda na loja " & c_loja & "."
								end if
							if alerta = "" then
								.preco_lista = rs("preco_lista")
								.desc_max_original = rs("desc_max")
								end if
							end if
						end if

					end if  'if (tem item)
				end with
			end if  'if (alerta)
		next


	if alerta = "" then
		dim r_cliente
		set r_cliente = New cl_CLIENTE
		if Not x_cliente_bd(cliente_selecionado, r_cliente) then Response.Redirect("aviso.asp?id=" & ERR_CLIENTE_NAO_CADASTRADO)
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
<link href="<%=URL_FILE__EPRINTER_CSS%>" Rel="stylesheet" Type="text/css" media="print">

<style type="text/css">
.CelFabr {
	width: 131px;
	vertical-align: top;
	}
.CelProd {
	width: 131px;
	vertical-align: top;
	}
.CelPerc {
	width: 70px;
	vertical-align: top;
	}
.CelValor {
	width: 75px;
	vertical-align: top;
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
<!-- *************************************************************** -->
<!-- **********  PÁGINA PARA EXIBIR DADOS DE CONFIRMAÇÃO  ********** -->
<!-- *************************************************************** -->
<body onload="focus();">
<center>

<form id="fOP" name="fOP" method="post" action="SenhaDescSupConfirma.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="cliente_selecionado" id="cliente_selecionado" value="<%=cliente_selecionado%>">
<input type="hidden" name="c_loja" id="c_loja" value="<%=c_loja%>">
<input type="hidden" name="c_supervisor_autorizador" id="c_supervisor_autorizador" value="<%=c_supervisor_autorizador%>">

<!-- FORÇA A CRIAÇÃO DE UM ARRAY MESMO QUANDO EXISTE SOMENTE 1 ITEM -->
<input type="hidden" name="c_fabricante" id="c_fabricante" value="">
<input type="hidden" name="c_produto" id="c_produto" value="">
<input type="hidden" name="c_desc_max_senha" id="c_desc_max_senha" value="">

<%
	for intCounter=Lbound(v_item) to Ubound(v_item)
		with v_item(intCounter)
%>
		<input type="hidden" name="c_fabricante" id="c_fabricante" value="<%=.fabricante%>">
		<input type="hidden" name="c_produto" id="c_produto" value="<%=.produto%>">
		<input type="hidden" name="c_desc_max_senha" id="c_desc_max_senha" value="<%=formata_perc_desc(.perc_desconto)%>">
<%		end with
	next
%>



<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">

<tr>
	<td align="right" valign="bottom"><p class="PEDIDO">Senha para Autorização de Desconto Superior<span class="C">&nbsp;</span></p></td>
</tr>
</table>
<br>

<!--  CLIENTE  -->
<table class="Qx" cellSpacing="0" style="width:450px;">
	<tr bgColor="#FFFFFF">
		<%	s = cnpj_cpf_formata(r_cliente.cnpj_cpf)
			if s="" then s="&nbsp;"
		%>
		<td class="MT"><span class="PLTe">CLIENTE</span>
		<br><p class="C"><%=s%></p>
		</td>
	</tr>
	<tr bgColor="#FFFFFF">
		<%	s = iniciais_em_maiusculas(r_cliente.nome)
			if s="" then s="&nbsp;"
		%>
		<td class="MDBE"><span class="PLTe">NOME</span>
		<br><p class="C"><%=s%></p>
		</td>
	</tr>
	<tr bgColor="#FFFFFF">
		<%	s = ""
			with r_cliente
				if .endereco <> "" then
					s = iniciais_em_maiusculas(.endereco)
					s_aux=Trim(.endereco_numero)
					if s_aux<>"" then s=s & ", " & s_aux
					s_aux=Trim(.endereco_complemento)
					if s_aux<>"" then s=s & " " & s_aux
					s_aux=iniciais_em_maiusculas(.bairro)
					if s_aux<>"" then s=s & " - " & s_aux
					s_aux=iniciais_em_maiusculas(.cidade)
					if s_aux<>"" then s=s & " - " & s_aux
					s_aux=.uf
					if s_aux<>"" then s=s & " - " & s_aux
					s_aux=cep_formata(.cep)
					if s_aux<>"" then s=s & " - " & s_aux
					end if
				end with
			if s="" then s="&nbsp;"
		%>
		<td class="MDBE"><span class="PLTe">ENDEREÇO</span>
		<p class="C"><%=s%></p>
		</td>
	</tr>
	<tr bgColor="#FFFFFF">
		<%	s = r_cliente.obs_crediticias
			if s="" then s="&nbsp;"
		%>
		<td class="MDBE"><span class="PLTe">OBSERVAÇÕES CREDITÍCIAS</span>
		<br><p class="C" style="color:red;"><%=s%></p>
		</td>
	</tr>
<!--  PULA LINHA  -->
	<tr bgColor="#FFFFFF">
		<td>&nbsp;</td>
	</tr>
<!--  AUTORIZADO POR  -->
	<tr bgColor="#FFFFFF">
		<td class="MT"><span class="PLTe">AUTORIZADO POR</span>
		<br><p class="C"><%=c_supervisor_autorizador & " - " & x_usuario(c_supervisor_autorizador)%></p>
		</td>
	</tr>
<!--  LOJA  -->
	<tr bgColor="#FFFFFF">
		<td class="MDBE"><span class="PLTe">LOJA</span>
		<br>
		<%	s = c_loja
			if (s<>"") And (s_nome_loja<>"") then s = s & " - "
			s = s & iniciais_em_maiusculas(s_nome_loja)
		%>
		<p class="C"><%=s%></p>
		</td>
	</tr>
</table>

<!--  PULA LINHA  -->
<br><br>

<table class="Qx" cellSpacing="0">
	<tr bgColor="#FFFFFF">
	<td class="MB CelFabr" style="vertical-align:bottom"><p class="PLTe">Fabricante</p></td>
	<td class="MB CelProd" style="vertical-align:bottom"><p class="PLTe">Produto</p></td>
	<td class="MB CelValor" style="vertical-align:bottom"><p class="PLTd">Preço de Lista</p></td>
	<td class="MB CelPerc" style="vertical-align:bottom"><p class="PLTd">Desc Máx Original (%)</p></td>
	<td class="MB CelValor" style="vertical-align:bottom"><p class="PLTd">Preço Mín Original</p></td>
	<td class="MB CelPerc" style="vertical-align:bottom"><p class="PLTd">Desc Máx Autorizado</p></td>
	<td class="MB CelValor" style="vertical-align:bottom"><p class="PLTd">Preço Mín Autorizado</p></td>
	</tr>

<%
	for intCounter=Lbound(v_item) to Ubound(v_item)
		with v_item(intCounter)
%>	
	<tr bgColor="#FFFFFF">
	<!--  FABRICANTE  -->
	<td class="MDBE CelFabr">
		<%	s = .fabricante
			if (s<>"") And (.nome_fabricante<>"") then s = s & " - "
			s = s & iniciais_em_maiusculas(.nome_fabricante)
		%>
		<p class="C"><%=s%></p>
		</td>

<!--  PRODUTO  -->
	<td class="MDB CelProd">
		<%	s = .produto
			if (s<>"") And (.nome_produto<>"") then s = s & " - "
			s = s & .nome_produto
		%>
		<p class="C"><%=s%></p>
		</td>

<!--  PREÇO DE LISTA  -->
	<td class="MDB CelValor">
		<p class="Cd"><%=formata_moeda(.preco_lista)%></p>
		</td>

<!--  DESCONTO MÁXIMO (ORIGINAL)  -->
	<td class="MDB CelPerc">
		<p class="Cd"><%=formata_perc_desc(.desc_max_original)%></p>
		</td>

<!--  PREÇO MÍNIMO ORIGINAL  -->
	<td class="MDB CelValor">
		<p class="Cd"><%=formata_moeda(.preco_lista-(.preco_lista*(.desc_max_original/100)))%></p>
		</td>

<!--  DESCONTO MÁXIMO (AUTORIZAÇÃO) -->
	<td class="MDB CelPerc">
		<p class="Cd" style="color:green;font-size:11pt;"><%=formata_perc_desc(.perc_desconto)%></p>
		</td>

<!--  PREÇO MÍNIMO AUTORIZADO -->
	<td class="MDB CelValor">
		<p class="Cd" style="color:green;font-size:11pt;"><%=formata_moeda(.preco_lista-(.preco_lista*(.perc_desconto/100)))%></p>
		</td>
	</tr>
	
<%		end with
	next
%>
</table>


<!-- ************   SEPARADOR   ************ -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>
<br>


<table class="notPrint" width="649" cellSpacing="0">
<tr>
	<td><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()" title="volta para página anterior">
		<img src="../botao/anterior.gif" width="176" height="55" border="0"></a></td>
	<td align="right"><div name="dCONFIRMA" id="dCONFIRMA"><a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fOPConfirma(fOP)" title="confirma o cadastramento da senha">
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