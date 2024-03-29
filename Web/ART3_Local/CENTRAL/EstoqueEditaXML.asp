<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/estoque.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ==================================================
'	  E S T O Q U E E D I T A E A N X M L . A S P
'     ==================================================
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

    class cl_ITEM_ESTOQUE_TELA_XML
		dim id_estoque
		dim fabricante
		dim produto
		dim qtde
		dim qtde_utilizada
		dim preco_fabricante
		dim data_ult_movimento
		dim sequencia
		dim descricao
		dim descricao_html
		dim ean
		dim vl_custo2
		dim vl_BC_ICMS_ST
		dim vl_ICMS_ST
		dim ncm
		dim cst
        dim aliq_ipi
        dim vl_ipi
        dim aliq_icms
        dim vl_frete
		end class


	dim usuario
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

	dim estoque_selecionado
	estoque_selecionado = Trim(request("estoque_selecionado"))
	if (estoque_selecionado = "") then Response.Redirect("aviso.asp?id=" & ERR_ESTOQUE_NAO_ESPECIFICADO)

	dim s, i, n
	dim r_estoque, v_item_bd, v_item
	dim s_nome_fabricante, s_produto, s_ean, s_descricao, s_descricao_html, s_qtde, s_vl_unitario, s_vl_total, m_vl_total, m_total_geral
	dim s_vl_BC_ICMS_ST, s_vl_ICMS_ST, s_vl_custo2, s_vl_total_custo2, m_vl_total_custo2, m_total_geral_custo2
	dim s_ncm, s_cst
	dim s_nome_nfe_emitente
    dim s_vl_diferenca, s_total_diferenca, m_vl_diferenca, m_total_diferenca, m_total_geral_diferenca
    dim s_aliq_ipi, s_vl_ipi, s_aliq_icms, s_vl_frete
    dim c_nfe_dt_hr_emissao
	
'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

  	dim c_perc_agio
    dim s_entrada_tipo

'	VERIFICA PERMISS�O DE ACESSO DO USU�RIO
	if Not operacao_permitida(OP_CEN_EDITA_ENTRADA_ESTOQUE, s_lista_operacoes_permitidas) then 
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	dim alerta, msg_erro
	alerta=""
	
	if Not le_estoque_agio(estoque_selecionado, r_estoque, msg_erro) then
		alerta = msg_erro
	else
		if Not le_estoque_item_xml(estoque_selecionado, v_item_bd, msg_erro) then alerta = msg_erro
		end if

	if alerta = "" then
		s_nome_fabricante = fabricante_descricao(r_estoque.fabricante)
        c_perc_agio = formata_numero(r_estoque.perc_agio, 4)
        if r_estoque.entrada_tipo = 1 then
            s_entrada_tipo = "Via XML"
        else
            s_entrada_tipo = "Manual"
            end if
		redim v_item(Ubound(v_item_bd))
		for i = Lbound(v_item_bd) to Ubound(v_item_bd)
			set v_item(i) = New cl_ITEM_ESTOQUE_TELA_XML
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
            v_item(i).aliq_ipi              = v_item_bd(i).aliq_ipi
            v_item(i).vl_ipi                = v_item_bd(i).vl_ipi
            v_item(i).aliq_icms             = v_item_bd(i).aliq_icms
            v_item(i).vl_frete              = v_item_bd(i).vl_frete
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
						.descricao = "N�O CADASTRADO"
						.descricao_html = "N�O CADASTRADO"
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

    c_nfe_dt_hr_emissao = Trim(Request("c_nfe_dt_hr_emissao"))

	dim blnValorEditavel, sAnoMesEstoque, sAnoMesHoje, s_valor_readonly
	blnValorEditavel = False
	sAnoMesEstoque = Left(formata_data_yyyymmdd(r_estoque.data_entrada), 6)
	sAnoMesHoje = Left(formata_data_yyyymmdd(Now), 6)
	if sAnoMesEstoque = sAnoMesHoje then blnValorEditavel = True
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
<script src="<%=URL_FILE__JQUERY_MY_GLOBAL%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
	$(document).ready(function () {
		<% if Not CADASTRAR_WMS_CD_ENTRADA_ESTOQUE then %>
		$(".trWmsCd").hide();
		<% end if %>
		// Observa��o: Unlike JavaScript indices, the CSS-based :nth-child(n) pseudo-class begins numbering at 1, not 0.
		// 1 - Numera��o da linha
		// 2 - Produto
		// 3 - EAN
		// 4 - Descri��o
		// 5 - NCM
		// 6 - CST
		// 7 - Qtde
		// 8 - Base C�lculo ICMS ST
		// 9 - Valor ICMS ST
		// 10 - Pre�o Fabricante
		// 11 - Total Pre�o Fabricante
		// 12 - Valor Refer�ncia
		// 13 - Total Valor Refer�ncia
		//$("#tableProduto thead th:nth-child(3), #tableProduto tbody td:nth-child(3)").hide();
		//$("#tableProduto thead th:nth-child(8), #tableProduto tbody td:nth-child(8)").hide();
		//$("#tableProduto thead th:nth-child(9), #tableProduto tbody td:nth-child(9)").hide();
		//$("#tableProduto thead th:nth-child(11), #tableProduto tbody td:nth-child(11)").hide();
		//$("#tdTotalGeralFabricante").hide();
	    //$("#tdPreTotalGeralFabricante").removeClass("MD").attr("colSpan", 7);
		$("#tdPreTotalGeralFabricante").removeClass("MD").attr("colSpan", 6);
		//$("#tdPreChecagem").attr("colSpan", 6);
		$("input:text:enabled:visible:not([readonly])").focus(function() {
			$(this).select();
		});
	});

	function recalcula_itens() {
	    var v_agio;
        var v_calculo = 0;
        var v_ipi;
        var v_aliq_ipi;
        var v_frete;
	    var s;
	    var iQtdeItens = '<%=UBound(v_item) + 1%>';
	    var f = fESTOQ;

	    s = $("#c_perc_agio").val();
	    if (s == "") {
	        v_agio = 0;
	    }
    
	    v_agio = converte_numero(s) / 100;
	    for (var i = 0; i <= iQtdeItens-1; i++) {
	        //v_calculo = converte_numero(f.c_vl_unitario[i].value)  * (1 + v_agio);
            //f.c_vl_custo2[i].value = formata_moeda(v_calculo);

            //calculo do valor do produto com IPI, frete e �gio
            v_calculo = converte_numero(f.c_vl_unitario[i].value);
            v_frete = converte_numero(f.c_vl_frete[i].value);
            v_frete = v_frete / converte_numero(f.c_qtde[i].value);
            v_calculo = v_calculo + v_frete;
            v_aliq_ipi = converte_numero(f.c_vl_aliq_ipi[i].value) / 100;
            if (v_aliq_ipi > 0) {
                v_ipi = converte_numero(formata_moeda(v_calculo * v_aliq_ipi));
            }
            else {
                v_ipi = converte_numero(f.c_vl_ipi.value);
                v_ipi = v_ipi / converte_numero(f.c_qtde.value);
            }
            v_calculo = v_calculo + v_ipi;
            v_agio = converte_numero(formata_moeda(v_calculo * v_perc_agio));
            v_calculo = converte_numero(formata_moeda(v_calculo + v_agio));
            f.c_vl_total_custo2.value = formata_moeda(v_calculo)

	        recalcula_total(i + 1);
	    }
	    //recalcula_total_nf();
	    return;

	}


	function recalcula_total_nf() {
	    var v_calculo;
	    var v_total;
	    var f;
	    var i;

	    f=fESTOQ;
	    v_calculo = 0;
	    v_total = 0;
	    for (i = 0; i < f.c_vl_total_custo2.length; i++)
	    {
	        v_calculo = converte_numero(f.c_vl_unitario[i].value); 
	        v_calculo = v_calculo * converte_numero(f.c_qtde[i].value);
	        v_total = v_total + v_calculo;
	    }
        // O CAMPO ABAIXO DEIXOU DE APRESENTAR O TOTAL DOS PRODUTOS E PASSOU A APRESENTAR O TOTAL DA NOTA FISCAL, POR SOLICITA��O, PARA FACILITAR A VISUALIZA��O NA CONSULTA
        //f.c_total_nf.value = formata_moeda(v_total);
		
	    return;

	}


function recalcula_total ( id ) {
var idx, m, f, i;
	f=fESTOQ;
	idx=parseInt(id)-1;
	if (f.c_produto[idx].value=="") return;
	
	m = converte_numero(f.c_vl_custo2[idx].value);
	f.c_vl_total_custo2[idx].value = formata_moeda(parseInt(f.c_qtde[idx].value) * m);
	m = 0;
	for (i = 0; i < f.c_vl_total_custo2.length; i++) m = m + converte_numero(f.c_vl_total_custo2[i].value);
    f.c_total_geral_custo2.value = formata_moeda(m);

    //O CAMPO ABAIXO DEIXOU DE APRESENTAR O TOTAL DOS PRODUTOS E PASSOU A APRESENTAR O TOTAL DA NOTA FISCAL, POR SOLICITA��O, PARA FACILITAR A VISUALIZA��O NA CONSULTA
    f.c_total_nf.value = formata_moeda(m);
}

function fESTOQConfirma( f ) {
    f.action="estoqueatualizaxml.asp";
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
<link href="<%=URL_FILE__EPRINTER_CSS%>" rel="stylesheet" type="text/css" media="print">

<style type="text/css">
#ckb_especial {
	margin: 0pt 2pt 1pt 15pt;
	vertical-align: top;
	}
.rbOpt
{
	vertical-align:bottom;
}
.lblOpt
{
	vertical-align:bottom;
}
.TxtEditavel{
	color: blue;
}
.TxtNfeEmitNome{
	width:640px;
}
.TxtErpFabr{
	width:100px;
	text-align:left;
}
.TxtErpDocumento{
	width:270px;
	margin-left:2pt;
}
.TxtNfeDtHrEmissao{
	width:80px;
	margin-left:2pt;
}
.TxtErpObs{
	width:642px;
	margin-left:2pt;
}
.TxtErpCodigo{
	width: 50px;
	padding-left:4px;
}
.TxtSubtitulo{
	width: 30px;
	padding-left:4px;
    text-align:left;
    color: grey;
}
.TxtErpCst{
	width: 30px;
	text-align:center;
}
.TdErpCodigo{
	width:50px;
	vertical-align: middle;
}
.TdNfeCodigo{
	width: 150px;
	vertical-align: middle;
}
.TdNfeDescricao{
	width: 320px;
	vertical-align: middle;
}
.TdNfeNcm{
	width: 60px;
	vertical-align: middle;
	text-align:center;
}
.TdErpCst{
	width: 30px;
	vertical-align: middle;
	text-align:center;
}
.TdNfeCst{
	width: 30px;
	vertical-align: middle;
	text-align:center;
}
.TdNfeCfop{
	width: 40px;
	vertical-align: middle;
	text-align:center;
}
.TdNfeUnid{
	width: 30px;
	vertical-align: middle;
	text-align:center;
}
.TdNfeQtde{
	width: 50px;
	vertical-align: middle;
	text-align:right;
}
.TdNfeVlUnit{
	width: 70px;
	vertical-align: middle;
	text-align:right;
}
.TdNfeVlTotal{
	width: 70px;
	vertical-align: middle;
	text-align:right;
}
.TdNfeVlBcIcms{
	width: 70px;
	vertical-align: middle;
	text-align:right;
}
.TdNfeVlIcms{
	width: 70px;
	vertical-align: middle;
	text-align:right;
}
.TdNfeVlIpi{
	width: 70px;
	vertical-align: middle;
	text-align:right;
}
.TdNfeAliqIcms{
	width: 40px;
	vertical-align: middle;
	text-align:right;
}
.TdNfeAliqIpi{
	width: 40px;
	vertical-align: middle;
	text-align:right;
}

</style>

<% if alerta <> "" then %>
<!-- ************************************************************ -->
<!-- **********  P�GINA PARA EXIBIR MENSAGENS DE ERRO  ********** -->
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
	<td align="center"><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()"><img src="..\botao\voltar.gif" width="176" height="55" border="0"></a></td>
</tr>
</table>
</center>
</body>




<% else %>
<!-- ************************************************************* -->
<!-- **********  P�GINA PARA EXIBIR DADOS DOS PRODUTOS  ********** -->
<!-- ************************************************************* -->
<body onload="focus();">
<center>

<form id="fESTOQ" name="fESTOQ" method="post">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="c_fabricante" id="c_fabricante" value="<%=r_estoque.fabricante%>" />
<input type="hidden" name="c_id_nfe_emitente" id="c_id_nfe_emitente" value="<%=r_estoque.id_nfe_emitente%>" />
<input type="hidden" name="estoque_selecionado" id="estoque_selecionado" value="<%=estoque_selecionado%>" />

<% if blnValorEditavel then %>
<input type="hidden" name="c_flag_valor_editavel" id="c_flag_valor_editavel" value="S" />
<% else %>
<input type="hidden" name="c_flag_valor_editavel" id="c_flag_valor_editavel" value="N" />
<% end if %>



<!--  I D E N T I F I C A � � O   D A   T E L A  -->
<table width="780" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><p class="PEDIDO">Estoque XML (altera��o)<span class="C">&nbsp;</span></p></td>
</tr>
</table>
<br>

<!--  EDI��O DAS MERCADORIAS NO ESTOQUE  -->
<table class="Qx" width="650" cellspacing="0" cellpadding="0">
<!--  EMPRESA / CD  -->
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
	<tr bgcolor="#FFFFFF">
        <td class="MDBE" width="50%" align="left" nowrap><span class="PLTe">Documento</span>
		<br><input name="c_documento" id="c_documento" maxlength="30" class="PLLe" style="width:270px;margin-left:2pt;color:darkblue;"
			value="<%=r_estoque.documento%>">
        </td>
        <td class="MDB" width="50%" align="left" nowrap><span class="PLTe">Emiss�o</span>
		<br /><input name="c_nfe_dt_hr_emissao" id="c_nfe_dt_hr_emissao" readonly tabindex=-1  class="PLLe" style="margin-left:2pt;"
            value="<%=c_nfe_dt_hr_emissao%>">
	    </td>
	</tr>
<!-- �GIO / TIPO ENTRADA  -->
	<tr bgcolor="#FFFFFF">
        <td class="MDBE" width="50%" align="left" nowrap><span class="PLTe">% �gio</span>
		<br><input name="c_perc_agio" id="c_perc_agio" <%=s_valor_readonly%>  class="PLLe" style="width:120px;margin-left:2pt;" 
			value="<%=c_perc_agio%>"
            onblur="this.value=formata_numero(this.value, 4); recalcula_itens();"/>
        </td>
        <td class="MDB" width="50%" align="left" nowrap>
            <span class="PLTe">Tipo de Entrada</span>
		<br><input name="s_entrada_tipo" id="s_entrada_tipo" readonly tabindex=-1 class="PLLe" style="margin-left:2pt;" 
			value="<%=s_entrada_tipo%>">
        </td>
	</tr>
<!--  DATA DE ENTRADA / USU�RIO  -->
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
<!--  ENTRADA ESPECIAL  -->
	<tr bgcolor="#FFFFFF"><td colspan="2" class="MDBE" align="left" nowrap><span class="PLTe">Tipo de Cadastramento</span>
		<br><input type="checkbox" class="rbOpt" tabindex="-1" id="ckb_especial" name="ckb_especial" value="ESPECIAL_ON"
		<%if r_estoque.entrada_especial <> 0 then Response.Write " checked" %>
		><span class="C lblOpt" style="cursor:default;color:darkblue;" onclick="fESTOQ.ckb_especial.click();">Entrada Especial</span>
	</td></tr>
<!--  OBS  -->
	<tr bgcolor="#FFFFFF">
	<td colspan="2" class="MDBE" align="left" nowrap><span class="PLTe">Observa��es</span>
		<br><textarea name="c_obs" id="c_obs" class="PLLe" rows="<%=Cstr(MAX_LINHAS_ESTOQUE_OBS)%>"
				style="width:642px;margin-left:2pt;"
				onkeypress="limita_tamanho(this,MAX_TAM_T_ESTOQUE_CAMPO_OBS);"
				onblur="this.value=trim(this.value);"
				><%=r_estoque.obs%></textarea>
	</td>
	</tr>
</table>
<br>

<!--  R E L A � � O   D E   P R O D U T O S  -->
<table id="tableProduto" class="Qx" cellspacing="0">
	<thead>
	<tr bgcolor="#FFFFFF">
	<th>&nbsp;</th>
	<td class="MB TdErpCodigo" align="left" style="vertical-align:bottom;"><span class="PLTe">C�D</span></td>
	<td class="MB TdNfeCodigo" align="left" style="vertical-align:bottom;"><span class="PLTe">EAN</span></td>
	<td class="MB TdNfeDescricao" align="left" style="vertical-align:bottom;"><span class="PLTe">DESCRI��O DO PROD/SERV</span></td>
	<td class="MB TdNfeNcm" align="left" style="vertical-align:bottom;"><span class="PLTe">NCM/SH</span></td>
	<td class="MB TdNfeCst" align="left" style="vertical-align:bottom;"><span class="PLTe">CST</span></td>
	<td class="MB TdNfeQtde" align="left" style="vertical-align:bottom;"><span class="PLTe">QUANT</span></td>
	<td class="MB TdNfeVlUnit" align="left" style="vertical-align:bottom;"><span class="PLTe">VL Nota</span></td>
	<td class="MB TdNfeVlUnit" align="left" style="vertical-align:bottom;"><span class="PLTe">VL Refer�ncia</span></td>
	<td class="MB TdNfeAliqIpi" align="left" style="vertical-align:bottom;"><span class="PLTe">A.IPI</span></td>
	<td class="MB TdNfeVlIpi" align="left" style="vertical-align:bottom;"><span class="PLTe">VL IPI</span></td>
	<td class="MB TdNfeAliqIcms" align="left" style="vertical-align:bottom;"><span class="PLTe">A.ICMS</span></td>
	<td class="MB TdNfeVlIpi" align="left" style="vertical-align:bottom;"><span class="PLTe">VL Frete</span></td>
	<td class="MB TdNfeVlTotal" align="left" style="vertical-align:bottom;"><span class="PLTe">VL TOTAL</span></td>
	</tr>
	</thead>

	<tbody>
<%	s_valor_readonly = " readonly tabindex=-1"
	if blnValorEditavel then s_valor_readonly = ""
	m_total_geral=0
	m_total_geral_custo2=0
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
				s_vl_total_custo2=formata_moeda(m_vl_total_custo2)
				m_total_geral_custo2=m_total_geral_custo2 + m_vl_total_custo2
				s_ncm = .ncm
				s_cst = .cst
                's_aliq_ipi = formata_moeda(.aliq_ipi)
                s_aliq_ipi = formata_numero(.aliq_ipi, 2)
                s_vl_ipi = formata_moeda(.vl_ipi)
                's_aliq_icms = formata_moeda(.aliq_icms)
                s_aliq_icms = formata_numero(.aliq_icms, 0)
      			s_vl_frete = formata_moeda(.vl_frete)
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
	<td align="left"><input name="c_linha" id="c_linha" readonly tabindex=-1 class="PLLe" maxlength="2" style="width:24px;text-align:right;color:#808080;" value="<%=Cstr(i) & ". " %>"></td>
	<td class="MDBE" align="left">
		<input name="c_produto" id="c_produto" readonly tabindex=-1 class="PLLe" style="width:50px;"
			value="<%=s_produto%>"></td>
	<td class="MDB" align="left">
		<input name="c_ean" id="c_ean" readonly tabindex=-1 class="PLLe" style="width:80px;"
			value="<%=s_ean%>"></td>
	<td class="MDB" align="left">
		<span class="PLLe" style="width:240px;"><%=s_descricao_html%></span>
		<input type=hidden name="c_descricao" id="c_descricao" value="<%=s_descricao%>">
	</td>
	<td class="MDB" align="left">
		<input name="c_ncm" id="c_ncm" class="PLLc TxtEditavel" maxlength="8" style="width:56px;"
			onkeypress="if (digitou_enter(true)) $(this).hUtil('focusNext'); filtra_numerico();"
			value="<%=s_ncm%>">
	</td>
	<td class="MDB" align="left">
		<input name="c_cst" id="c_cst" class="PLLc TxtEditavel" maxlength="3" style="width:40px;"
			onkeypress="if (digitou_enter(true)) $(this).hUtil('focusNext'); filtra_numerico();"
			value="<%=s_cst%>">
	</td>
	<td class="MDB" align="right">
		<input name="c_qtde" id="c_qtde" maxlength="4" class="PLLd" style="width:30px;color:darkblue;"
			value="<%=s_qtde%>"
			onkeypress="if (digitou_enter(true)) $(this).hUtil('focusNext'); filtra_numerico();"
			onblur="recalcula_itens();"
			></td>
	<td class="MDB" align="right">
		<input name="c_vl_unitario" id="c_vl_unitario" <%=s_valor_readonly%> class="PLLd" style="width:62px;"
			value="<%=s_vl_unitario%>"
			onkeypress="if (digitou_enter(true)) $(this).hUtil('focusNext'); filtra_moeda();"
			onblur="this.value=formata_moeda(this.value); if (converte_numero(this.value)<0) {alert('Valor inv�lido!!');this.focus();} else {recalcula_itens();}"
			/></td>
	<td class="MDB" align="right">
		<input name="c_vl_custo2" id="c_vl_custo2" <%=s_valor_readonly%> class="PLLd" style="width:62px;"
			value="<%=s_vl_custo2%>"
			onkeypress="if (digitou_enter(true)) $(this).hUtil('focusNext'); filtra_moeda();"
			onblur="this.value=formata_moeda(this.value); if (converte_numero(this.value)<0) {alert('Valor inv�lido!!');this.focus();} else {recalcula_total(<%=Cstr(i)%>);}"
			/>
		</td>
	<td class="MDB" align="right">
		<input name="c_aliq_ipi" id="c_aliq_ipi" readonly tabindex=-1 class="PLLd" style="width:70px;"
			value="<%=s_aliq_ipi%>">
	    </td>
	<td class="MDB" align="right">
		<input name="c_vl_ipi" id="c_vl_ipi" readonly tabindex=-1 class="PLLd" style="width:70px;"
			value="<%=s_vl_ipi%>">
	    </td>
	<td class="MDB" align="right">
		<input name="c_aliq_icms" id="c_aliq_icms" readonly tabindex=-1 class="PLLd" style="width:70px;"
			value="<%=s_aliq_icms%>">
	    </td>
	<td class="MDB" align="right">
		<input name="c_vl_frete" id="c_vl_ipi" readonly tabindex=-1 class="PLLd" style="width:70px;"
			value="<%=s_vl_frete%>">
	    </td>
	<td class="MDB" align="right">
		<input name="c_vl_total_custo2" id="c_vl_total_custo2" readonly tabindex=-1 class="PLLd" style="width:70px;"
			value="<%=s_vl_total_custo2%>" />
		</td>
	</tr>
<% next %>
	</tbody>
	
	<tfoot>
	<tr>
	<td colspan="11" class="MD" id="tdPreTotalGeralFabricante">&nbsp;</td>

	<td>&nbsp;</td>
    <td class="MDBE" align="left"><p class="Cd">Total NF</p></td>
    <!--O CAMPO ABAIXO DEIXOU DE APRESENTAR O TOTAL DOS PRODUTOS E PASSOU A APRESENTAR O TOTAL DA NOTA FISCAL, POR SOLICITA��O, PARA FACILITAR A VISUALIZA��O NA CONSULTA-->
	<!--<td class="MDB" align="right"><input name="c_total_nf" id="c_total_nf" class="PLLd" style="width:62px;color:black;" 
		value='<%=formata_moeda(m_total_geral)%>'></td>-->
    <td class="MDB" align="right"><input name="c_total_nf" id="c_total_nf" class="PLLd" style="width:62px;color:black;" 
		value='<%=formata_moeda(m_total_geral_custo2)%>'></td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<td>&nbsp;</td>
	<td class="MD">&nbsp;</td>
	<td class="MDB" align="right"><input name="c_total_geral_custo2" id="c_total_geral_custo2" class="PLLd" style="width:70px;color:Blue;"
		value='<%=formata_moeda(m_total_geral_custo2)%>' readonly tabindex=-1 /></td>
	</tr>
	</tfoot>
</table>


<!--  ASSEGURA CRIA��O DE UM ARRAY DE CAMPOS, MESMO QUANDO HOUVER SOMENTE 1 PRODUTO!! -->
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
<input type="hidden" name="c_vl_total_custo2" id="c_vl_total_custo2" value="" />
<input type="hidden" name="c_vl_diferenca" id="c_vl_diferenca" value="">
<input type="hidden" name="c_aliq_ipi" id="c_aliq_ipi" value="">
<input type="hidden" name="c_vl_ipi" id="c_vl_ipi" value="">
<input type="hidden" name="c_aliq_icms" id="c_aliq_icms" value="">

<!-- ************   SEPARADOR   ************ -->
<table width="780" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>
<br>

<table class="notPrint" width="780" cellspacing="0">
<tr>
	<td align="left">
		<a name="bCANCELAR" id="bCANCELAR" href="javascript:history.back()" title="cancela a opera��o">
			<img src="../botao/cancelar.gif" width="176" height="55" border="0"></a>
		</td>
	<td align="right"><div name="dCONFIRMA" id="dCONFIRMA">
		<a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fESTOQConfirma(fESTOQ)" title="confirma as altera��es">
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