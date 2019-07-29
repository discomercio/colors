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
'	  P E D I D O D E V O L U C A O . A S P
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

	dim s, usuario, loja, pedido_selecionado
	usuario = Trim(Session("usuario_atual"))
	loja = Trim(Session("loja_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 
	If (loja = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	if Not operacao_permitida(OP_LJA_CADASTRA_DEVOLUCAO_PRODUTO, s_lista_operacoes_permitidas) then 
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	pedido_selecionado = ucase(Trim(request("pedido_selecionado")))
	if (pedido_selecionado = "") then Response.Redirect("aviso.asp?id=" & ERR_PEDIDO_NAO_ESPECIFICADO)
	s = normaliza_num_pedido(pedido_selecionado)
	if s <> "" then pedido_selecionado = s
	
	dim i, n
	dim s_fabricante, s_produto, s_descricao, s_descricao_html, s_qtde, s_vl_unitario, s_vl_TotalItem, m_TotalItem, m_TotalDestePedido
	dim s_cor, s_devolucao_anterior, s_readonly
	
'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	dim r_pedido, v_item, v_devol, alerta, msg_erro
	alerta=""
	if Not le_pedido(pedido_selecionado, r_pedido, msg_erro) then 
		alerta = msg_erro
	else
		if Trim(r_pedido.loja) <> loja then Response.Redirect("aviso.asp?id=" & ERR_PEDIDO_INVALIDO)
		if Not le_pedido_item(pedido_selecionado, v_item, msg_erro) then alerta = msg_erro
		end if

	if alerta = "" then
		if r_pedido.st_entrega <> ST_ENTREGA_ENTREGUE then
			alerta = "Pedido " & pedido_selecionado & " não consta como entregue, portanto, não é possível processar a sua devolução."
			end if
		end if
	
	if alerta = "" then
		redim v_devol(Ubound(v_item))
		for i=Lbound(v_devol) to Ubound(v_devol)
			set v_devol(i) = New cl_ITEM_DEVOLUCAO_MERCADORIAS
			v_devol(i).pedido		= v_item(i).pedido
			v_devol(i).fabricante	= v_item(i).fabricante
			v_devol(i).produto		= v_item(i).produto
			v_devol(i).qtde			= v_item(i).qtde
			next
				
		if Not estoque_verifica_mercadorias_para_devolucao(v_devol, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
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
	<title>LOJA<%=MontaNumPedidoExibicaoTitleBrowser(pedido_selecionado)%></title>
	</head>



<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
function posiciona_foco_prox_linha( idx ) {
var f,i, i_base;
	f=fPED;
	i_base=idx;
	if (i_base<0) i_base=0;
	for (i=i_base; i<f.c_qtde_devolucao.length; i++) {
		if (!f.c_qtde_devolucao[i].readOnly) {
			f.c_qtde_devolucao[i].focus();
			return true;
			}
		}
	return false;
}

function verifica_qtde_max( campo, idx ) {
var f,p,n;
	if (campo.readOnly) return true;
	f=fPED;
	p=converte_numero(idx)-1;
	n=converte_numero(f.c_qtde[p].value)-converte_numero(f.c_devolucao_anterior[p].value);
	if (converte_numero(campo.value)>n) {
		alert("A quantidade máxima para devolução é de " + n + " unidades!!");
		campo.focus();
		return false;
		}
	return true;
}

function recalcula_valores() {
var f, i, n, m, t;
	f=fPED;
	t=0
	for (i=0; i<f.c_qtde_devolucao.length; i++) {
		if (!f.c_qtde_devolucao[i].readOnly) {
			n=converte_numero(f.c_qtde_devolucao[i].value);
			m=converte_numero(f.c_vl_unitario[i].value);
			m=n*m;
			t=t+m;
			f.c_vl_total[i].value=formata_moeda(m);
			}
		}
	f.c_total_geral.value=formata_moeda(t);
}

function fPEDConfirma( f ) {
var b, i, n;
	b=false;
	for (i=0; i<f.c_qtde_devolucao.length; i++) {
		if (!f.c_qtde_devolucao[i].readOnly) {
			if (converte_numero(f.c_qtde_devolucao[i].value)>0) {
				b=true;
				if (trim(f.c_motivo[i].value)=="") {
					alert("Informe o motivo da devolução!!");
					f.c_motivo[i].focus();
					return;
					}
				}
			n=converte_numero(f.c_qtde[i].value)-converte_numero(f.c_devolucao_anterior[i].value);
			if (converte_numero(f.c_qtde_devolucao[i].value) > n) {
				alert("A quantidade máxima para devolução é de " + n + " unidades!!");
				f.c_qtde_devolucao[i].focus();
				return;
				}
			}
		}
	if (!b) {
		alert("Não foi especificada nenhuma mercadoria para devolução!!");
		return;
		}
		
	b=window.confirm("Confirma a devolução de mercadorias?");
	if (!b) return;
	
	f.action="pedidodevolucaoconfirma.asp";
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
<div class="MtAlerta" style="width:600px;font-weight:bold;" align="center"><P style='margin:5px 2px 5px 2px;'><%=alerta%></p></div>
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
<!-- ************************************************* -->
<!-- **********  PÁGINA EDITAR QUANTIDADES  ********** -->
<!-- ************************************************* -->
<body onload="posiciona_foco_prox_linha(-1);">
<center>

<form id="fPED" name="fPED" method="post">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="pedido_selecionado" id="pedido_selecionado" value='<%=pedido_selecionado%>'>

<!--  I D E N T I F I C A Ç Ã O   D O   P E D I D O -->
<%=MontaHeaderIdentificacaoPedido(pedido_selecionado, r_pedido, 852)%>
<br>

<!--  DESCRIÇÃO DA OPERAÇÃO -->
<table width="649" cellPadding="0" CellSpacing="0">
<tr><td><p class="Expl">DEVOLUÇÃO</p></td></tr>
<tr><td>
	<p class="Expl">A devolução de mercadorias só é possível em pedido que já tenha sido entregue ao cliente.</p>
	</td>
</tr>
</table>
<br>
	
<!--  R E L A Ç Ã O   D E   P R O D U T O S  -->
<table class="Qx" cellSpacing="0">
	<tr bgColor="#FFFFFF">
	<td class="MB" valign="bottom"><p class="PLTe">Fabr</p></td>
	<td class="MB" valign="bottom"><p class="PLTe">Produto</p></td>
	<td class="MB" valign="bottom"><p class="PLTe">Descrição</p></td>
	<td class="MB" valign="bottom"><p class="PLTd">Qtde</p></td>
	<td class="MB" valign="bottom"><p class="PLTd">Devol<br>Anter</p></td>
	<td class="MB" valign="bottom"><p class="PLTd">Devolver</p></td>
	<td class="MB" valign="bottom"><p class="PLTe">Motivo</p></td>
	<td class="MB" valign="bottom"><p class="PLTd">Valor<br>Unitário</p></td>
	<td class="MB" valign="bottom"><p class="PLTd">Total<br>Devolução</p></td>
	</tr>

<% m_TotalDestePedido=0
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
			m_TotalItem=0
			s_vl_TotalItem=formata_moeda(m_TotalItem)
			m_TotalDestePedido=m_TotalDestePedido + m_TotalItem
			end with
		s_devolucao_anterior=""
		with v_devol(n)
			if .qtde_devolvida_anteriormente<>0 then 
				s_devolucao_anterior=Cstr(.qtde_devolvida_anteriormente)
				s_cor = "darkorange"
				end if
			if (.qtde - .qtde_devolvida_anteriormente) > 0 then s_readonly = ""
			end with
			
	 else
		s_fabricante=""
		s_produto=""
		s_descricao=""
		s_descricao_html=""
		s_qtde=""
		s_devolucao_anterior=""
		s_vl_unitario=""
		s_vl_TotalItem=""
		s_readonly = "readonly tabindex=-1"
		end if
%>
	<tr>
	<td class="MDBE"><input name="c_fabricante" id="c_fabricante" class="PLLe" style="width:25px; color:<%=s_cor%>"
		value='<%=s_fabricante%>' readonly tabindex=-1></td>
	<td class="MDB"><input name="c_produto" id="c_produto" class="PLLe" style="width:54px; color:<%=s_cor%>"
		value='<%=s_produto%>' readonly tabindex=-1></td>
	<td class="MDB" style="width:269px;">
		<span class="PLLe" style="color:<%=s_cor%>"><%=s_descricao_html%></span>
		<input type="hidden" name="c_descricao" id="c_descricao" value='<%=s_descricao%>'>
	</td>
	<td class="MDB" align="right"><input name="c_qtde" id="c_qtde" class="PLLd" style="width:38px; color:<%=s_cor%>"
		value='<%=s_qtde%>' readonly tabindex=-1></td>
	<td class="MDB" align="right"><input name="c_devolucao_anterior" id="c_devolucao_anterior" class="PLLd" style="width:40px; color:<%=s_cor%>"
		value='<%=s_devolucao_anterior%>' readonly tabindex=-1></td>
	<td class="MDB" align="right"><input name="c_qtde_devolucao" id="c_qtde_devolucao" class="PLLd" maxlength="4" style="width:40px;color:red" onkeypress="if (digitou_enter(true)) fPED.c_motivo[<%=cstr(i-1)%>].focus(); filtra_numerico();" onblur="if (verifica_qtde_max(this,<%=Cstr(i)%>)) recalcula_valores();"
		value='' <%=s_readonly%>></td>
	<td class="MDB"><input name="c_motivo" id="c_motivo" class="PLLe" maxlength="80" style="width:200px; color:red" onkeypress="if (digitou_enter(true)) {if (!posiciona_foco_prox_linha(<%=Cstr(i)%>)) bCONFIRMA.focus();} filtra_nome_identificador();" onblur="this.value=trim(this.value);"
		value='' <%=s_readonly%>></td>
	<td class="MDB" align="right"><input name="c_vl_unitario" id="c_vl_unitario" class="PLLd" style="width:62px; color:<%=s_cor%>"
		value='<%=s_vl_unitario%>' readonly tabindex=-1></td>
	<td class="MDB" align="right"><input name="c_vl_total" id="c_vl_total" class="PLLd" style="width:70px; color:<%=s_cor%>" 
		value='<%=s_vl_TotalItem%>' readonly tabindex=-1></td>
	</tr>
<% next %>
	<tr>
	<td colspan="8" class="MD">&nbsp;</td>
	<td class="MDB" align="right"><input name="c_total_geral" id="c_total_geral" class="PLLd" style="width:70px;color:red;" 
		value='<%=formata_moeda(m_TotalDestePedido)%>' readonly tabindex=-1></td>
	</tr>
</table>

<!-- ************   SEPARADOR   ************ -->
<table width="852" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>
<br>

<!-- ************   BOTÕES   ************ -->
<table width="649" cellSpacing="0">
<tr>
	<td align="left"><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()" title="volta para página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a>
	</td>
	<td align="right"><div name="dCONFIRMA" id="dCONFIRMA">
		<a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fPEDConfirma(fPED)" title="confirma a devolução">
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
