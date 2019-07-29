<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/estoque.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ====================================================
'	  E S T O Q U E C O N S U L T A E A N . A S P
'     ====================================================
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

	dim estoque_selecionado
	estoque_selecionado = Trim(request("estoque_selecionado"))
	if (estoque_selecionado = "") then Response.Redirect("aviso.asp?id=" & ERR_ESTOQUE_NAO_ESPECIFICADO)

	dim url_back
	url_back = Trim(request("url_back"))
	
    dim c_perc_agio
    dim s_entrada_tipo
	
	dim r_estoque, v_item_bd, v_item
	dim s, i, n
	dim s_nome_fabricante, s_produto, s_ean, s_descricao, s_descricao_html, s_qtde, s_vl_unitario, s_vl_total, m_vl_total, m_total_geral
	dim s_vl_custo2, s_vl_total_custo2, m_vl_total_custo2, m_total_geral_custo2
	dim s_vl_BC_ICMS_ST, s_vl_ICMS_ST, s_ncm, s_cst
    dim s_aliq_ipi, s_aliq_icms, s_vl_ipi
	dim s_nome_nfe_emitente
	dim s_vl_diferenca, s_total_diferenca, m_vl_diferenca, m_total_diferenca, m_total_geral_diferenca
	
'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

	dim alerta, msg_erro
	alerta=""
	
	if Not le_estoque_agio(estoque_selecionado, r_estoque, msg_erro) then
		alerta = msg_erro
	else
		if Not le_estoque_item_xml(estoque_selecionado, v_item_bd, msg_erro) then alerta = msg_erro
		end if

	if alerta = "" then
		s_nome_fabricante = fabricante_descricao(r_estoque.fabricante)
        c_perc_agio = r_estoque.perc_agio
        if r_estoque.entrada_tipo = 1 then
            s_entrada_tipo = "Via XML"
        else
            s_entrada_tipo = "Manual"
            end if
		redim v_item(Ubound(v_item_bd))
		for i = Lbound(v_item_bd) to Ubound(v_item_bd)
			set v_item(i) = New cl_ITEM_ESTOQUE_TELA_AGIO
			v_item(i).id_estoque			= v_item_bd(i).id_estoque
			v_item(i).fabricante			= v_item_bd(i).fabricante
			v_item(i).produto				= v_item_bd(i).produto
			v_item(i).qtde					= v_item_bd(i).qtde
			v_item(i).qtde_utilizada		= v_item_bd(i).qtde_utilizada
			v_item(i).preco_fabricante		= v_item_bd(i).preco_fabricante
			v_item(i).vl_custo2				= v_item_bd(i).vl_custo2
			v_item(i).vl_BC_ICMS_ST			= v_item_bd(i).vl_BC_ICMS_ST
			v_item(i).vl_ICMS_ST			= v_item_bd(i).vl_ICMS_ST
			v_item(i).data_ult_movimento	= v_item_bd(i).data_ult_movimento
			v_item(i).sequencia				= v_item_bd(i).sequencia
			v_item(i).ncm					= v_item_bd(i).ncm
			v_item(i).cst					= v_item_bd(i).cst
			v_item(i).descricao = ""
			v_item(i).descricao_html = ""
			v_item(i).ean = ""
            v_item(i).aliq_ipi				= v_item_bd(i).aliq_ipi
            v_item(i).aliq_icms				= v_item_bd(i).aliq_icms
            v_item(i).vl_ipi				= v_item_bd(i).vl_ipi
			next
		
		for i = Lbound(v_item) to Ubound(v_item)
			with v_item(i)
				if .produto <> "" then
					s = "SELECT fabricante, produto, descricao, descricao_html, ean FROM t_PRODUTO WHERE" & _
						" (fabricante='" & .fabricante & "')" & _
						" AND (produto='" & .produto & "')"
					set rs = cn.execute(s)
					if Not rs.Eof then
						.descricao = Trim("" & rs("descricao"))
						.descricao_html = Trim("" & rs("descricao_html"))
						.ean = Trim("" & rs("ean"))
					else
						.descricao = "NÃO CADASTRADO"
						.descricao_html = "NÃO CADASTRADO"
						end if
					if rs.State <> 0 then rs.Close
					end if
				end with
			next
		end if
	
	s_nome_nfe_emitente = ""
	if alerta = "" then
		s = "SELECT id, razao_social FROM t_NFe_EMITENTE WHERE (id = " & r_estoque.id_nfe_emitente & ")"
		if rs.State <> 0 then rs.Close
		set rs = cn.execute(s)
		if Not rs.Eof then
			s_nome_nfe_emitente = Trim("" & rs("razao_social"))
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
	$(document).ready(function () {
		<% if Not CADASTRAR_WMS_CD_ENTRADA_ESTOQUE then %>
		$(".trWmsCd").hide();
		<% end if %>
		// Observação: Unlike JavaScript indices, the CSS-based :nth-child(n) pseudo-class begins numbering at 1, not 0.
		// 1 - Numeração da linha
		// 2 - Produto
		// 3 - EAN
		// 4 - Descrição
		// 5 - NCM
		// 6 - CST
		// 7 - Qtde
		// 8 - Base Cálculo ICMS ST
		// 9 - Valor ICMS ST
		// 10 - Preço Fabricante
		// 11 - Total Preço Fabricante
        // 12 - Valor Referência
		// 13 - Alíquota IPI
		// 14 - Valor IPI
		// 15 - Alíquota ICMS
		// 16 - Total Valor Referência
		//$("#tableProduto thead th:nth-child(3), #tableProduto tbody td:nth-child(3)").hide();
		$("#tableProduto thead th:nth-child(8), #tableProduto tbody td:nth-child(8)").hide();
		$("#tableProduto thead th:nth-child(9), #tableProduto tbody td:nth-child(9)").hide();
		$("#tableProduto thead th:nth-child(11), #tableProduto tbody td:nth-child(11)").hide();
		//$("#tdTotalGeralFabricante").hide();
		//$("#tdPreTotalGeralFabricante").removeClass("MD").attr("colSpan", 7);
		$("#tdPreTotalGeralFabricante").removeClass("MD").attr("colSpan", 6);
		$("#tdPreChecagem").attr("colSpan", 6);
		$("input:text:enabled:visible:not([readonly])").focus(function() {
			$(this).select();
		});
	});

function fESTOQModifica( f ) {
	f.action="estoqueeditaean.asp";
	dMODIFICA.style.visibility="hidden";
	window.status = "Aguarde ...";
	f.submit();
}

function fESTOQRemove( f ) {
var b;
	b=window.confirm('Confirma a exclusão deste registro de entrada de mercadorias no estoque?');
	if (b){
		f.action="estoqueremove.asp";
		dREMOVE.style.visibility="hidden";
		window.status = "Aguarde ...";
		f.submit();
		}
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
<link href="<%=URL_FILE__EPRINTER_CSS%>" rel="stylesheet" type="text/css" media="print">

<style type="text/css">
#ckb_especial_aux {
	margin: 0pt 2pt 1pt 15pt;
	vertical-align: top;
	}
#ckb_kit {
	margin: 0pt 2pt 1pt 15pt;
	vertical-align: top;
	}
#ckb_devolucao {
	margin: 0pt 2pt 1pt 15pt;
	vertical-align: top;
	}
.tdValor
{
	width:62px;
}
.tdVlTotal
{
	width:70px;
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
<table cellspacing="0">
<tr>
	<td align="center"><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()"><img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
</tr>
</table>
</center>
</body>




<% else %>
<!-- ************************************************************* -->
<!-- **********  PÁGINA PARA EXIBIR DADOS DOS PRODUTOS  ********** -->
<!-- ************************************************************* -->
<body onload="focus();">
<center>

<form id="fESTOQ" name="fESTOQ" method="post">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="c_fabricante" id="c_fabricante" value="<%=r_estoque.fabricante%>">
<input type="hidden" name="c_id_nfe_emitente" id="c_id_nfe_emitente" value="<%=r_estoque.id_nfe_emitente%>">
<input type="hidden" name='estoque_selecionado' id="estoque_selecionado" value='<%=estoque_selecionado%>'>
<!-- É NECESSÁRIO CRIAR UM CAMPO DO TIPO HIDDEN PARA QUE A PÁGINA SEGUINTE CONSIGA
	 RECUPERAR A INFORMAÇÃO REFERENTE A ESTE CAMPO, JÁ QUE REQUEST.FORM() EM UM
	 CAMPO DO TIPO CHECKBOX QUE ESTÁ DISABLED RETORNA VAZIO.
-->
<% if r_estoque.entrada_especial <> 0 then s="ESPECIAL_ON" else s="" %>
<input type="hidden" name="ckb_especial" id="ckb_especial" value="<%=s%>">

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="780" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Estoque (consulta)</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>
<br>

<!--  CADASTRO DA ENTRADA DE MERCADORIAS NO ESTOQUE  -->
<table class="Qx" width="650" cellspacing="0" cellpadding="0">
<!--  EMPRESA  -->
	<tr class="trWmsCd">
		<td colspan="2">
			<table width="100%" cellpadding="0" cellspacing="0">
				<tr>
					<td class="MT" align="left">
						<span class="PLTe">Empresa</span>
						<br />
						<span class="PLLe"><%=obtem_apelido_empresa_NFe_emitente(r_estoque.id_nfe_emitente)%></span>
					</td>
				</tr>
			</table>
		</td>
	</tr>
<!--  FABRICANTE  -->
	<tr bgcolor="#FFFFFF"><td colspan="2" class="MDBE" align="left" nowrap><span class="PLTe">Fabricante</span>
		<%	s = r_estoque.fabricante
			if (s<>"") And (s_nome_fabricante<>"") then s = s & " - " & s_nome_fabricante %>
		<br><input name="c_fabricante_aux" id="c_fabricante_aux" readonly tabindex=-1 class="PLLe" style="width:460px;margin-left:2pt;"
				value="<%=s%>"></td></tr>
<!--  DOCUMENTO  -->
	<tr bgcolor="#FFFFFF"><td colspan="2" class="MDBE" align="left" nowrap><span class="PLTe">Documento</span>
		<br><input name="c_documento" id="c_documento" readonly tabindex=-1 class="PLLe" style="width:270px;margin-left:2pt;"
			value="<%=r_estoque.documento%>"></td></tr>
<!-- ÁGIO / TIPO ENTRADA  -->
	<tr bgcolor="#FFFFFF">
        <td class="MDBE" width="50%" align="left" nowrap><span class="PLTe">% Ágio</span>
		<br><input name="c_perc_agio" id="c_perc_agio" readonly tabindex=-1 class="PLLe" style="width:120px;margin-left:2pt;" 
			value="<%=c_perc_agio%>">
        </td>
        <td class="MDB" width="50%" align="left" nowrap>
            <span class="PLTe">Tipo de Entrada</span>
		<br><input name="s_entrada_tipo" id="s_entrada_tipo" readonly tabindex=-1 class="PLLe" style="margin-left:2pt;" 
			value="<%=s_entrada_tipo%>">
        </td>
	</tr>
<!--  DATA DE ENTRADA / USUÁRIO  -->
	<tr bgcolor="#FFFFFF"><td class="MDBE" width="50%" align="left" nowrap><span class="PLTe">Data de Entrada no Estoque</span>
		<%	s = formata_hhnnss_para_hh_nn_ss(r_estoque.hora_entrada)
			if s <> "" then s = " - " & s
			s = formata_data(r_estoque.data_entrada) & s %>
		<br><input name="c_data_entrada" id="c_data_entrada" readonly tabindex=-1 class="PLLe" style="width:120px;margin-left:2pt;"
			value="<%=s%>"></td>
		<td class="MDB" width="50%" align="left" nowrap><span class="PLTe">Cadastrado por</span>
		<br><input name="c_usuario" id="c_usuario" readonly tabindex=-1 class="PLLe" style="width:120px;margin-left:2pt;"
			value="<%=r_estoque.usuario%>"></td>
		</tr>
<!--  TIPO DE CADASTRAMENTO  -->
	<tr bgcolor="#FFFFFF"><td colspan="2" class="MDBE" align="left" nowrap><span class="PLTe">Tipo de Cadastramento</span>
		<!--  ENTRADA ESPECIAL  -->
		<br><input type="checkbox" disabled tabindex="-1" id="ckb_especial_aux" name="ckb_especial_aux" value=""
		<%if r_estoque.entrada_especial <> 0 then Response.Write " checked" %>
		><span class="C" style="cursor:default;vertical-align:bottom;">Entrada Especial</span>
		<!--  KIT  -->
		<br><input type="checkbox" disabled tabindex="-1" id="ckb_kit" name="ckb_kit" value=""
		<%if r_estoque.kit <> 0 then Response.Write " checked" %>
		><span class="C" style="cursor:default;vertical-align:bottom;">Kit</span>
		<!--  DEVOLUÇÃO  -->
		<br><input type="checkbox" disabled tabindex="-1" id="ckb_devolucao" name="ckb_devolucao" value=""
		<%if r_estoque.devolucao_status <> 0 then Response.Write " checked" %>
		><span class="C" style="cursor:default;vertical-align:bottom;">Devolução</span>
	</td></tr>
<!--  OBS  -->
	<tr bgcolor="#FFFFFF">
	<td colspan="2" class="MDBE" align="left" nowrap><span class="PLTe">Observações</span>
		<br><textarea name="c_obs" id="c_obs" class="PLLe" rows="<%=Cstr(MAX_LINHAS_ESTOQUE_OBS)%>"
				style="width:642px;margin-left:2pt;" onkeypress="limita_tamanho(this,MAX_TAM_T_ESTOQUE_CAMPO_OBS);" onblur="this.value=trim(this.value);"
				readonly tabindex=-1><%=r_estoque.obs%></textarea>
	</td>
	</tr>
</table>
<br>

<!--  R E L A Ç Ã O   D E   P R O D U T O S  -->
<table id="tableProduto" class="Qx" cellspacing="0" border="0">
	<thead valign="bottom">
	<tr bgcolor="#FFFFFF">
	<th>&nbsp;</th>
	<th class="MB" align="left" valign="bottom"><span class="PLTe">Produto</span></th>
	<th class="MB" align="left" valign="bottom"><span class="PLTe">EAN</span></th>
	<th class="MB" align="left" valign="bottom"><span class="PLTe">Descrição</span></th>
	<th class="MB" align="center" valign="bottom" align="center"><span class="PLTe">NCM</span></th>
	<th class="MB" align="center" valign="bottom" align="center"><span class="PLTe">CST</span></th>
	<th class="MB" align="right" valign="bottom"><span class="PLTd">Qtde</span></th>
	<th class="MB tdValor" align="right" valign="bottom"><span class="PLTd">Base Cálc ICMS ST (Unit)</span></th>
	<th class="MB tdValor" align="right" valign="bottom"><span class="PLTd">Valor ICMS ST (Unit)</span></th>
	<th class="MB tdValor" align="right" valign="bottom"><span class="PLTd">Valor Unit</span></th>
	<th class="MB tdVlTotal" align="right" valign="bottom"><span class="PLTd">Valor Total</span></th>
    <th class="MB tdValor" align="right" valign="bottom"><span class="PLTd">Valor Referência</span></th>
    <th class="MB" align="right" valign="bottom" style="width:62px;"><span class="PLTd">A. IPI</span></th>
    <th class="MB" align="right" valign="bottom" style="width:62px;"><span class="PLTd">V. IPI</span></th>
    <th class="MB" align="right" valign="bottom" style="width:62px;"><span class="PLTd">A. ICMS</span></th>
	<th class="MB tdValor" align="right" valign="bottom"><span class="PLTd">Total Referência</span></th>																										  
	<th class="MB tdVlTotal" align="right" valign="bottom" style="width:62px;"><span class="PLTd">Total Diferença</span></th>
	</tr>
	</thead>
	<tbody>
<%	m_total_geral=0
	m_total_geral_custo2=0
	m_total_diferenca=0
	m_total_geral_diferenca=0
	n = Lbound(v_item)-1
	for i=1 to MAX_PRODUTOS_ENTRADA_ESTOQUE
		n = n+1
		if n <= Ubound(v_item) then
			with v_item(n)
				s_produto = .produto
				s_ean = .ean
				s_descricao = .descricao
				s_descricao_html = produto_formata_descricao_em_html(.descricao_html)
				s_qtde = .qtde
				s_vl_BC_ICMS_ST = formata_moeda(.vl_BC_ICMS_ST)
				s_vl_ICMS_ST = formata_moeda(.vl_ICMS_ST)
				s_vl_unitario = formata_moeda(.preco_fabricante)
				m_vl_total = .qtde * .preco_fabricante
				s_vl_total=formata_moeda(m_vl_total)
				m_total_geral=m_total_geral + m_vl_total
				s_vl_custo2 = formata_moeda(.vl_custo2)
				m_vl_total_custo2 = .qtde * .vl_custo2
				s_vl_total_custo2 = formata_moeda(m_vl_total_custo2)
				m_total_geral_custo2=m_total_geral_custo2 + m_vl_total_custo2
				s_ncm = .ncm
				s_cst = .cst
                s_aliq_ipi = .aliq_ipi
                s_aliq_icms = .aliq_icms
                s_vl_ipi = formata_moeda(.vl_ipi)
				m_vl_diferenca = .vl_custo2 - .preco_fabricante
				s_vl_diferenca = formata_moeda(m_vl_diferenca)
				m_total_diferenca = .qtde * m_vl_diferenca
				s_total_diferenca = formata_moeda(m_total_diferenca)
				m_total_geral_diferenca = m_total_geral_diferenca + m_total_diferenca
				end with
		else
			exit for
			end if
%>
	<tr>
	<td><input name="c_linha" id="c_linha" readonly tabindex=-1 class="PLLe" maxlength="2" style="width:24px;text-align:right;color:#808080;" value="<%=Cstr(i) & ". " %>"></td>
	<td class="MDBE">
		<input name="c_produto" id="c_produto" readonly tabindex=-1 class="PLLe" style="width:50px;"
			value="<%=s_produto%>"></td>
	<td class="MDB">
		<input name="c_ean" id="c_ean" readonly tabindex=-1 class="PLLe" style="width:80px;"
			value="<%=s_ean%>"></td>
	<td class="MDB">
		<span class="PLLe" style="width:240px;"><%=s_descricao_html%></span>
		<input type="hidden" name="c_descricao" id="c_descricao" value="<%=s_descricao%>">
	</td>
	<td class="MDB">
		<input name="c_ncm" id="c_ncm" readonly tabindex=-1 class="PLLc" style="width:56px;"
			value="<%=s_ncm%>">
	</td>
	<td class="MDB">
		<input name="c_cst" id="c_cst" readonly tabindex=-1 class="PLLc" style="width:40px;"
			value="<%=s_cst%>">
	</td>
	<td class="MDB" align="right">
		<input name="c_qtde" id="c_qtde" readonly tabindex=-1 class="PLLd" style="width:30px;"
			value="<%=s_qtde%>"></td>
	<td class="MDB" align="right">
		<input name="c_vl_BC_ICMS_ST" id="c_vl_BC_ICMS_ST" readonly tabindex=-1 class="PLLd" style="width:62px;"
			value="<%=s_vl_BC_ICMS_ST%>"></td>
	<td class="MDB" align="right">
		<input name="c_vl_ICMS_ST" id="c_vl_ICMS_ST" readonly tabindex=-1 class="PLLd" style="width:62px;"
			value="<%=s_vl_ICMS_ST%>"></td>
	<td class="MDB" align="right">
		<input name="c_vl_unitario" id="c_vl_unitario" readonly tabindex=-1 class="PLLd" style="width:62px;"
			value="<%=s_vl_unitario%>"></td>
	<td class="MDB" align="right">
		<input name="c_vl_total" id="c_vl_total" readonly tabindex=-1 class="PLLd" style="width:70px;"
			value="<%=s_vl_total%>"></td>

	<td class="MDB" align="right">
		<input name="c_vl_custo2" id="c_vl_custo2" readonly tabindex=-1 class="PLLd" style="width:62px;"
			value="<%=s_vl_custo2%>"></td>

	<td class="MDB" align="right">
		<input name="c_aliq_ipi" class="PLLd" maxlength="12" style="width:62px;"
            value="<%=s_aliq_ipi%>">
		</td>

	<td class="MDB" align="right">
		<input name="c_vl_ipi" class="PLLd" maxlength="12" style="width:62px;"
            value="<%=s_vl_ipi%>">
		</td>

	<td class="MDB" align="right">
		<input name="c_aliq_icms" class="PLLd" maxlength="12" style="width:62px;"
            value="<%=s_aliq_icms%>">
		</td>

	<td class="MDB" align="right">
		<input name="c_vl_total_custo2" id="c_vl_total_custo2" readonly tabindex=-1 class="PLLd" style="width:70px;"
			value="<%=s_vl_total_custo2%>"></td>
	<td class="MDB" align="right">
		<input name="c_vl_total_diferenca" id="c_vl_total_diferenca" readonly tabindex=-1 class="PLLd" style="width:70px;"
			value="<%=s_total_diferenca%>"></td>
	</tr>
<% next %>
	</tbody>
	
	<tfoot>
	<tr>
	<td colspan="13" id="tdPreTotalGeralFabricante">&nbsp;</td>
	<td class="MDBE" align="left"><p class="Cd">Total NF</p></td>
	<td class="MDB" align="right" id="tdTotalGeralFabricante"><input name="c_total_geral" id="c_total_geral" class="PLLd" style="width:70px;color:blue;" 
		value='<%=formata_moeda(m_total_geral)%>' readonly tabindex=-1></td>
	<td>&nbsp;</td>
    <td class="MD">&nbsp;</td>
    <td class="MD">&nbsp;</td>
    <td class="MD">&nbsp;</td>
	<td class="MDBE" align="right"><input name="c_total_geral_custo2" id="c_total_geral_custo2" class="PLLd" style="width:70px;color:blue;" 
		value='<%=formata_moeda(m_total_geral_custo2)%>' readonly tabindex=-1></td>
	<td class="MDB" align="right"><input name="c_total_geral_diferenca" id="c_total_geral_diferenca" class="PLLd" style="width:70px;color:blue;" 
		value='<%=formata_moeda(m_total_geral_diferenca)%>' readonly tabindex=-1></td>
	</tr>
	</tfoot>
</table>


<!--  ASSEGURA CRIAÇÃO DE UM ARRAY DE CAMPOS, MESMO QUANDO HOUVER SOMENTE 1 PRODUTO!! -->
<input type="hidden" name="c_linha" id="c_linha" value="">
<input type="hidden" name="c_produto" id="c_produto" value="">
<input type="hidden" name="c_ean" id="c_ean" value="">
<input type="hidden" name="c_descricao" id="c_descricao" value="">
<input type="hidden" name="c_ncm" id="c_ncm" value="">
<input type="hidden" name="c_cst" id="c_cst" value="">
<input type="hidden" name="c_qtde" id="c_qtde" value="">
<input type="hidden" name="c_vl_BC_ICMS_ST" id="c_vl_BC_ICMS_ST" value="">
<input type="hidden" name="c_vl_ICMS_ST" id="c_vl_ICMS_ST" value="">
<input type="hidden" name="c_vl_unitario" id="c_vl_unitario" value="">
<input type="hidden" name="c_vl_total" id="c_vl_total" value="">
<input type="hidden" name="c_vl_custo2" id="c_vl_custo2" value="">
<input type="hidden" name="c_vl_total_custo2" id="c_vl_total_custo2" value="">
<input type="hidden" name="c_aliq_ipi" value="">
<input type="hidden" name="c_aliq_icms" value="">
<input type="hidden" name="c_vl_ipi" value="">


<!-- ************   SEPARADOR   ************ -->
<table width="780" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>


<table class="notPrint" width="780" cellspacing="0" border="0">
<% if operacao_permitida(OP_CEN_ENTRADA_MERCADORIAS_ESTOQUE, s_lista_operacoes_permitidas) then %>
<tr>
	<td width="100%" align="left">
		<span class="Rc"><a href="EstoqueEntrada.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="Cadastrar nova entrada" class="LPagInicial" style="font-size:9pt;">Cadastrar Nova Entrada</a></span>
<!--	</td>
    <td width="50%" align="left">-->
		<span class="Rc">&nbsp;</span>
        <span class="Rc"><a href="EstoqueEntradaViaXmlUpload.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="Cadastrar nova entrada XML" class="LPagInicial" style="font-size:9pt;">Cadastrar Nova Entrada XML</a></span>
	</td>
</tr>
<% end if %>
<tr>
	<td align="center">
		<table width="100%">
			<tr>
			<% if operacao_permitida(OP_CEN_EDITA_ENTRADA_ESTOQUE, s_lista_operacoes_permitidas) And (r_estoque.kit = 0) And (r_estoque.devolucao_status=0) then s="left" else s="center" %>
			<td align="<%=s%>">
				<%	if url_back <> "" then 
						s="resumo.asp" & "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")) 
					else 
						s="javascript:history.back()" 
						end if
				%>
				<a name="bVOLTAR" id="bVOLTAR" href="<%=s%>" title="volta para página anterior">
					<img src="../botao/voltar.gif" width="176" height="55" border="0"></a>
				</td>

			<% if operacao_permitida(OP_CEN_EDITA_ENTRADA_ESTOQUE, s_lista_operacoes_permitidas) And (r_estoque.kit = 0) And (r_estoque.devolucao_status=0) then %>
				<td align="center"><div name="dREMOVE" id="dREMOVE">
					<a name="bREMOVE" id="bREMOVE" href="javascript:fESTOQRemove(fESTOQ)" title="remove este registro de entrada de mercadorias no estoque">
						<img src="../botao/remover.gif" width="176" height="55" border="0"></a></div>
					</td>
				<td align="right"><div name="dMODIFICA" id="dMODIFICA">
					<a name="bMODIFICA" id="bMODIFICA" href="javascript:fESTOQModifica(fESTOQ)" title="edita os dados deste registro de entrada de mercadorias no estoque">
						<img src="../botao/modificar.gif" width="176" height="55" border="0"></a></div>
					</td>
			<% end if %>
			</tr>
		</table>
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