<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/estoque.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ===========================================
'	  P E D I D O S P L I T . A S P
'     ===========================================
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

	dim s, usuario, pedido_selecionado
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

	pedido_selecionado = ucase(Trim(request("pedido_selecionado")))
	if (pedido_selecionado = "") then Response.Redirect("aviso.asp?id=" & ERR_PEDIDO_NAO_ESPECIFICADO)
	s = normaliza_num_pedido(pedido_selecionado)
	if s <> "" then pedido_selecionado = s
	
	dim i, n
	dim s_fabricante, s_produto, s_descricao, s_descricao_html, s_qtde, s_vl_unitario, s_vl_total, m_vl_total, m_total_geral
	dim s_cor, s_falta, s_readonly
	dim total_produtos
	total_produtos=0
	
'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	dim r_pedido, v_item, v_disp, alerta, msg_erro
	alerta=""
	if Not le_pedido(pedido_selecionado, r_pedido, msg_erro) then 
		alerta = msg_erro
	else
		if Not le_pedido_item(pedido_selecionado, v_item, msg_erro) then alerta = msg_erro
		end if

	if alerta = "" then
		if Not IsPedidoSplitable(r_pedido.st_entrega) then
			alerta = "Pedido " & pedido_selecionado & " não pode gerar filhote de pedido."
			end if
		end if
	
	dim insert_request_guid
	insert_request_guid = Trim(Request.Form("request_guid"))

	dim r_cliente
	set r_cliente = New cl_CLIENTE
	if Not x_cliente_bd(r_pedido.id_cliente, r_cliente) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
	
	if alerta = "" then
		if Not IsPedidoEncerrado(r_pedido.st_entrega) then
			redim v_disp(Ubound(v_item))
			for i=Lbound(v_disp) to Ubound(v_disp)
				set v_disp(i) = New cl_ITEM_STATUS_ESTOQUE
				v_disp(i).pedido		= v_item(i).pedido
				v_disp(i).fabricante	= v_item(i).fabricante
				v_disp(i).produto		= v_item(i).produto
				v_disp(i).qtde			= v_item(i).qtde
				next
				
			if Not estoque_verifica_status_item(v_disp, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)

			for i=Lbound(v_item) to Ubound(v_item)
				with v_item(i)
					if Trim("" & .produto) <> "" then
						if .qtde > 0 then total_produtos = total_produtos + .qtde
						end if
					end with
				next

			if total_produtos <= 1 then
				alerta = "O pedido não possui a quantidade de produtos suficiente para permitir a operação de split."
				end if
			end if
		end if
	
	if alerta = "" then
		if Len(Trim(r_cliente.endereco)) > CLng(MAX_TAMANHO_CAMPO_ENDERECO) then
			alerta = "Endereço no cadastro do cliente excede o tamanho máximo permitido:<br>Tamanho atual: " & Cstr(Len(r_cliente.endereco)) & " caracteres<br>Tamanho máximo: " & Cstr(MAX_TAMANHO_CAMPO_ENDERECO) & " caracteres"
		elseif Trim(r_cliente.endereco_numero) = "" then
			alerta = "O endereço no cadastro do cliente deve ser corrigido, separando as informações do número e complemento nos campos adequados."
			end if
		end if
	
	if alerta = "" then
		if CLng(r_pedido.st_end_entrega) <> 0 then
			if Len(r_pedido.EndEtg_endereco) > CLng(MAX_TAMANHO_CAMPO_ENDERECO) then
				alerta = "Endereço de entrega excede o tamanho máximo permitido:<br>Tamanho atual: " & Cstr(Len(r_pedido.EndEtg_endereco)) & " caracteres<br>Tamanho máximo: " & Cstr(MAX_TAMANHO_CAMPO_ENDERECO) & " caracteres"
			elseif Trim(r_pedido.EndEtg_endereco_numero) = "" then
				alerta = "O endereço de entrega deve ser corrigido, separando as informações do número e complemento nos campos adequados."
				end if
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
function posiciona_proximo_foco( idx ) {
var f,i, i_base;
	f=fPED;
	i_base=idx;
	if (i_base<0) i_base=0;
	for (i=i_base; i<f.c_qtde_split.length; i++) {
		if (!f.c_qtde_split[i].readOnly) {
			f.c_qtde_split[i].focus();
			return true;
			}
		}
	return false;
}

function verifica_qtde_max( campo, idx ) {
var f,p,n;
	if (campo.readOnly) return;
	f=fPED;
	p=converte_numero(idx)-1;
	n=converte_numero(f.c_qtde[p].value)-converte_numero(f.c_qtde_falta[p].value);
	if (converte_numero(campo.value)>n) {
		alert("A quantidade máxima para o split é de " + n + "!!");
		campo.focus();
		return;
		}
}

function fPEDConfirma( f ) {
var b, i, n;
	b=false;
	for (i=0; i<f.c_qtde_split.length; i++) {
		if (!f.c_qtde_split[i].readOnly) {
			if (converte_numero(f.c_qtde_split[i].value)>0) b=true;
			n=converte_numero(f.c_qtde[i].value)-converte_numero(f.c_qtde_falta[i].value);
			if (converte_numero(f.c_qtde_split[i].value) > n) {
				alert("A quantidade máxima para o split é de " + n + "!!");
				f.c_qtde_split[i].focus();
				return;
				}
			}
		}
	if (!b) {
		alert("Não foi especificado nenhum produto para fazer o split do pedido!!");
		return;
		}
	
	f.action="pedidosplitconfirma.asp";
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
<!-- ********************************************************** -->
<!-- **********  PÁGINA PARA EDITAR ITENS DO PEDIDO  ********** -->
<!-- ********************************************************** -->
<body onload="posiciona_proximo_foco(-1);">
<center>

<form id="fPED" name="fPED" method="post">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="pedido_selecionado" id="pedido_selecionado" value='<%=pedido_selecionado%>'>
<input type="hidden" name="insert_request_guid" id="insert_request_guid" value="<%=insert_request_guid%>" />

<!--  I D E N T I F I C A Ç Ã O   D O   P E D I D O -->
<%=MontaHeaderIdentificacaoPedido(pedido_selecionado, r_pedido, 649)%>
<br>

<!--  DESCRIÇÃO DA OPERAÇÃO -->
<table width="649" cellPadding="0" CellSpacing="0">
<tr><td><p class="Expl">SPLIT</p></td></tr>
<tr><td>
	<p class="Expl">A operação de "split" permite desmembrar o pedido criando um filhote de pedido, que poderá conter somente os produtos disponíveis para entrega.</p>
	</td>
</tr>
</table>
<br>

<!--  R E L A Ç Ã O   D E   P R O D U T O S  -->
<table class="Qx" cellSpacing="0">
	<tr bgColor="#FFFFFF">
	<td class="MB"><p class="PLTe">Fabr</p></td>
	<td class="MB"><p class="PLTe">Produto</p></td>
	<td class="MB"><p class="PLTe">Descrição</p></td>
	<td class="MB"><p class="PLTd">Qtde</p></td>
	<td class="MB"><p class="PLTd">Faltam</p></td>
	<td class="MB"><p class="PLTd">Qtde Split</p></td>
	<td class="MB"><p class="PLTd">Valor Unit</p></td>
	<td class="MB"><p class="PLTd">Valor Total</p></td>
	</tr>

<% m_total_geral=0
   n = Lbound(v_item)-1
   for i=1 to MAX_ITENS 
	 n = n+1
	 s_cor = "black"
	 s_readonly = "readonly tabindex=-1"
	 if n <= Ubound(v_item) then
		with v_item(n)
			s_fabricante=.fabricante
			s_produto=.produto
			s_descricao=.descricao
			s_descricao_html=produto_formata_descricao_em_html(.descricao_html)
			s_qtde=.qtde
			s_vl_unitario=formata_moeda(.preco_NF)
			m_vl_total=.qtde * .preco_NF
			s_vl_total=formata_moeda(m_vl_total)
			m_total_geral=m_total_geral + m_vl_total
			end with
		s_falta=""
		if Not IsPedidoEncerrado(r_pedido.st_entrega) then
			with v_disp(n)
				if .qtde_estoque_sem_presenca<>0 then s_falta=Cstr(.qtde_estoque_sem_presenca)
				s_cor = x_cor_item(.qtde, .qtde_estoque_vendido, .qtde_estoque_sem_presenca)
				if .qtde_estoque_vendido > 0 then s_readonly = ""
				end with
			end if
		
	 else
		s_fabricante=""
		s_produto=""
		s_descricao=""
		s_descricao_html=""
		s_qtde=""
		s_falta=""
		s_vl_unitario=""
		s_vl_total=""
		s_readonly = "readonly tabindex=-1"
		end if
%>
	<tr>
	<td class="MDBE"><input name="c_fabricante" id="c_fabricante" class="PLLe" style="width:25px; color:<%=s_cor%>"
		value='<%=s_fabricante%>' readonly tabindex=-1></td>
	<td class="MDB"><input name="c_produto" id="c_produto" class="PLLe" style="width:54px; color:<%=s_cor%>"
		value='<%=s_produto%>' readonly tabindex=-1></td>
	<td class="MDB">
		<span class="PLLe" style="width:271px;color:<%=s_cor%>"><%=s_descricao_html%></span>
		<input type="hidden" name="c_descricao" id="c_descricao" value='<%=s_descricao%>'>
	</td>
	<td class="MDB" align="right"><input name="c_qtde" id="c_qtde" class="PLLd" style="width:40px; color:<%=s_cor%>"
		value='<%=s_qtde%>' readonly tabindex=-1></td>
	<td class="MDB" align="right"><input name="c_qtde_falta" id="c_qtde_falta" class="PLLd" style="width:40px; color:<%=s_cor%>"
		value='<%=s_falta%>' readonly tabindex=-1></td>
	<td class="MDB" align="right"><input name="c_qtde_split" id="c_qtde_split" class="PLLd" maxlength="4" style="width:40px; color:green" onkeypress="if (digitou_enter(true)) {if (!posiciona_proximo_foco(<%=Cstr(i)%>)) bCONFIRMA.focus();} filtra_numerico();" onblur="verifica_qtde_max(this,<%=Cstr(i)%>);"
		value='' <%=s_readonly%>></td>
	<td class="MDB" align="right"><input name="c_vl_unitario" id="c_vl_unitario" class="PLLd" style="width:62px; color:<%=s_cor%>"
		value='<%=s_vl_unitario%>' readonly tabindex=-1></td>
	<td class="MDB" align="right"><input name="c_vl_total" id="c_vl_total" class="PLLd" style="width:70px; color:<%=s_cor%>" 
		value='<%=s_vl_total%>' readonly tabindex=-1></td>
	</tr>
<% next %>
	<tr>
	<td colspan="7" class="MD">&nbsp;</td>
	<td class="MDB" align="right"><input name="c_total_geral" id="c_total_geral" class="PLLd" style="width:70px;color:blue;" 
		value='<%=formata_moeda(m_total_geral)%>' readonly tabindex=-1></td>
	</tr>
</table>

<!-- ************   SEPARADOR   ************ -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>
<br>

<!-- ************   BOTÕES   ************ -->
<table class="notPrint" width="649" cellSpacing="0">
<tr>
	<td align="left"><a name="bVOLTAR" id="bVOLTAR" href="pedido.asp?pedido_selecionado=<%=pedido_selecionado & "&url_back=X&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>"
		title="volta para página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a>
	</td>
	<td align="right"><div name="dCONFIRMA" id="dCONFIRMA">
		<a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fPEDConfirma(fPED)" title="confirma as alterações">
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
'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>
