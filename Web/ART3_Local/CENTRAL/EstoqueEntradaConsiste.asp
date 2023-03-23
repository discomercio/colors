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
'	  E S T O Q U E E N T R A D A C O N S I S T E . A S P
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

	dim s, i, j, n, flag_ok, usuario
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

	dim v_item, s_fabricante, s_nome_fabricante, s_documento, ckb_especial, s_obs
	dim s_produto, s_ean, s_descricao, s_descricao_html, s_qtde, s_vl_unitario, s_vl_total, m_vl_total, m_total_geral
	dim s_vl_custo2, s_vl_total_custo2, m_vl_total_custo2, m_total_geral_custo2
	dim s_id_nfe_emitente
	dim s_nome_nfe_emitente
	
'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	s_fabricante = normaliza_codigo(retorna_so_digitos(Request.Form("c_fabricante")), TAM_MIN_FABRICANTE)
	s_documento = Trim(Request.Form("c_documento"))
	ckb_especial = Trim(Request.Form("ckb_especial"))
	s_obs = Trim(Request.Form("c_obs"))
	s_id_nfe_emitente = Trim(Request.Form("c_id_nfe_emitente"))

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
				.fabricante=s_fabricante
				.produto=Ucase(Trim(Request.Form("c_codigo")(i)))
				s = Trim(Request.Form("c_qtde")(i))
				if IsNumeric(s) then .qtde = CLng(s) else .qtde = 0
				end with
			end if
		next
	
	dim alerta
	alerta=""
	
	if s_fabricante = "" then
		alerta = "O código do fabricante não foi preenchido."
	elseif s_documento = "" then
		alerta = "O campo documento não foi preenchido."
		end if
	
	if alerta = "" then
		s = "SELECT nome, razao_social FROM t_FABRICANTE WHERE (fabricante='" & s_fabricante & "')"
		set rs = cn.execute(s)
		s_nome_fabricante = ""
		if rs.Eof then
			alerta = "Fabricante " & s_fabricante & " não está cadastrado."
		else
			s_nome_fabricante = Trim("" & rs("razao_social"))
			if s_nome_fabricante = "" then s_nome_fabricante = Trim("" & rs("nome"))
			end if
		end if
	
	if alerta = "" then
		if CADASTRAR_WMS_CD_ENTRADA_ESTOQUE then
			if s_id_nfe_emitente = "" then
				alerta=texto_add_br(alerta)
				alerta=alerta & "É necessário informar uma empresa."
				end if
			end if
		end if
	
	if alerta = "" then
		if CADASTRAR_WMS_CD_ENTRADA_ESTOQUE then
			s = "SELECT id, razao_social FROM t_NFe_EMITENTE WHERE (id = " & s_id_nfe_emitente & ")"
			if rs.State <> 0 then rs.Close
			set rs = cn.execute(s)
			if rs.Eof then
				alerta=texto_add_br(alerta)
				alerta=alerta & "Empresa " & s_id_nfe_emitente & " NÃO está cadastrada."
			else
				s_nome_nfe_emitente = Trim("" & rs("razao_social"))
				end if
			end if
		end if
	
	if alerta = "" then
	'	VERIFICA CADA UM DOS PRODUTOS
		for i=Lbound(v_item) to Ubound(v_item)
			with v_item(i)
				if .qtde <= 0 then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Produto " & .produto & ": quantidade " & cstr(.qtde) & " é inválida."
					end if
	
				s = "SELECT * FROM t_PRODUTO WHERE"
				if IsEAN(.produto) then
					s = s & " (ean='" & .produto & "')"
				else
					s = s & " (fabricante='" & s_fabricante & "') AND (produto='" & .produto & "')"
					end if
				
				s = s & " AND (excluido_status=0)"
				
				if rs.State <> 0 then rs.Close
				set rs = cn.execute(s)
				if rs.Eof then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Produto " & .produto & " NÃO está cadastrado."
				else
					flag_ok = True
					if IsEAN(.produto) And (s_fabricante<>Trim("" & rs("fabricante"))) then
						flag_ok = False
						alerta=texto_add_br(alerta)
						alerta=alerta & "Produto " & .produto & " NÃO pertence ao fabricante " & s_fabricante & "."
						end if
					if flag_ok then
					'	CARREGA CÓDIGO INTERNO DO PRODUTO
						.produto = Trim("" & rs("produto"))
						.ean = Trim("" & rs("ean"))
						.descricao = Trim("" & rs("descricao"))
						.descricao_html = Trim("" & rs("descricao_html"))
						.preco_fabricante = rs("preco_fabricante")
						.vl_custo2 = rs("vl_custo2")
						end if
					end if

				for j=Lbound(v_item) to (i-1)
					if (.produto = v_item(j).produto) then
						alerta=texto_add_br(alerta)
						alerta=alerta & "Produto " & .produto & ": linha " & renumera_com_base1(Lbound(v_item),i) & " repete o mesmo produto da linha " & renumera_com_base1(Lbound(v_item),j) & "."
						exit for
						end if
					next
				end with
			next
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
<script src="<%=URL_FILE__JQUERY_MY_GLOBAL%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_I18N%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">

	$(document).ready(function() {
		$("#divMsgAlerta").hide();
		$("#divAjaxProgress").hide();
		$("#divDialogBox").hide();
		$("#divDialogBox").hUtilUI('dialog_modal');

		$("#btnDivMsgCancelar").button().click(function(event) {
			event.preventDefault();
			$("#divMsgAlerta").hide();
			$("#dCONFIRMA").show();
		});

		$("#btnDivMsgConfirmar").button().click(function(event) {
			event.preventDefault();
			fESTOQ.submit();
			$(this).hide();
		});

		$("input:text:enabled:visible:not([readonly])").attr("autocomplete", "off");

		<% if Not CADASTRAR_WMS_CD_ENTRADA_ESTOQUE then %>
		$(".trWmsCd").hide();
		<% end if %>

		// Observação: Unlike JavaScript indices, the CSS-based :nth-child(n) pseudo-class begins numbering at 1, not 0.
		// 1 - Numeração da linha
		// 2 - Produto
		// 3 - EAN
		// 4 - Descrição
		// 5 - NCM
		// 6 - NCM (redigite)
		// 7 - CST
		// 8 - CST (redigite)
		// 9 - Qtde
		// 10 - Base Cálculo ICMS ST
		// 11 - Valor ICMS ST
		// 12 - Preço Fabricante
		// 13 - Total Preço Fabricante
		// 14 - Valor Referência
		// 15 - Alíquota IPI
        // 16 - Valor IPI
        // 17 - Alíquota ICMS
		// 18 - Total Valor Referência
		//$("#tableProduto thead th:nth-child(3), #tableProduto tbody td:nth-child(3)").hide();
		$("#tableProduto thead th:nth-child(10), #tableProduto tbody td:nth-child(10)").hide();
		$("#tableProduto thead th:nth-child(11), #tableProduto tbody td:nth-child(11)").hide();
		$("#tableProduto thead th:nth-child(13), #tableProduto tbody td:nth-child(13)").hide();
		$("#tdTotalGeralFabricante").hide();
	    $("#tdPreTotalGeralFabricante").removeClass("MD").attr("colSpan", 8);
		//$("#tdPreTotalGeralFabricante").removeClass("MD").attr("colSpan", 11);
	    //$("#tdPreChecagem").attr("colSpan", 8);
	    $("#tdPreChecagem").attr("colSpan", 13);
		$("input:text:enabled:visible:not([readonly])").focus(function() {
			$(this).select();
		});
		$("input:text:enabled:visible:not([readonly]):first").focus();

		// Tratamento p/ bug do jQuery-ui Dialog: ao tentar mover o dialog em uma tela que está c/ scroll
		// vertical, o dialog é "redesenhado" mais abaixo da posição do cursor na mesma medida do deslocamento do
		// scroll vertical. A movimentação do dialog ocorre c/ esse espaço em branco entre o cursor e o dialog.
		$(document).scroll(function(e) {
			if ($(".ui-widget-overlay")) //the dialog has popped up in modal view
			{
				//fix the overlay so it scrolls down with the page
				$(".ui-widget-overlay").css({
					position: 'fixed',
					top: '0'
				});
				//get the current popup position of the dialog box
				pos = $(".ui-dialog").position();
				//adjust the dialog box so that it scrolls as you scroll the page
				$(".ui-dialog").css({
					position: 'fixed',
					top: pos.y
				});
			}
		});
	});

function cancela_onpaste() {
	return false;
}


function recalcula_itens() {
    var v_agio;
    var v_calculo = 0;
    var s;
    var iQtdeItens = '<%=UBound(v_item) + 1%>';
    var f = fESTOQ;
    var v_ipi;
    var v_aliq_ipi;

    s = $("#c_perc_agio").val();
    if (s == "") {
        v_agio = 0;
    }
    
    v_agio = converte_numero(s) / 100;
    for (var i = 0; i <= iQtdeItens-1; i++) {
        v_calculo = converte_numero(f.c_vl_unitario[i].value);
        //ACRÉSCIMO DO IPI
        s = f.c_aliq_ipi[i].value
        v_aliq_ipi = converte_numero(s) / 100;
        if (v_aliq_ipi > 0) {
            v_ipi = converte_numero(formata_moeda(v_calculo * v_aliq_ipi));
            f.c_vl_ipi[i].value = formata_moeda(v_ipi);
        }
        else {
            v_ipi = converte_numero(f.c_vl_ipi[i].value);
        }
        v_calculo = v_calculo + v_ipi;	
        //APLICAÇÃO DO ÁGIO
        v_calculo = v_calculo   * (1 + v_agio);
        f.c_vl_custo2[i].value = formata_moeda(v_calculo);
        recalcula_total(i + 1);
        recalcula_diferenca(i + 1);
    }
    	recalcula_total_nf();
    return;

}

function recalcula_diferenca( id ) {
    var v_calculo;
    var f;
    var idx;
    var i;
    var v_total_dif;
    var v_ipi;

	f=fESTOQ;
	idx=parseInt(id)-1;
	v_calculo = 0;
	v_total_dif = 0;
	v_ipi = 0;
	s = $("#c_perc_agio").val();
    if (s == "") {
        v_agio = 0;
    }
    
	v_agio = converte_numero(s) / 100;
	//DIFERENÇA UNITÁRIA
	v_calculo = (converte_numero(f.c_vl_custo2[idx].value) - converte_numero(f.c_vl_unitario[idx].value));
    //f.c_vl_diferenca[idx].value = formata_moeda(v_calculo);
	//DIFERENÇA TOTAL
	v_calculo = v_calculo * converte_numero(f.c_qtde[idx].value);
	f.c_vl_total_diferenca[idx].value = formata_moeda(v_calculo);
	for (i = 0; i < f.c_vl_total_diferenca.length; i++)
		{
		v_total_dif = v_total_dif + converte_numero(f.c_vl_total_diferenca[i].value);
		}
	f.c_total_geral_diferenca.value = formata_moeda(v_total_dif);
		
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
    for (i = 0; i < f.c_vl_total_diferenca.length; i++)
    {
        v_calculo = converte_numero(f.c_vl_unitario[i].value); 
        v_calculo = v_calculo * converte_numero(f.c_qtde[i].value);
        v_total = v_total + v_calculo;
    }
    f.c_total_nf.value = formata_moeda(v_total);
		
    return;

}

function recalcula_total( id ) {
var idx, m, m2, f, i;
	f=fESTOQ;
	idx=parseInt(id)-1;
	if (f.c_produto[idx].value=="") return;
	m=converte_numero(f.c_vl_unitario[idx].value);
	if (f.c_vl_unitario[idx].value != formata_moeda(m)) f.c_vl_unitario[idx].value = formata_moeda(m);
	if (trim(f.c_vl_custo2[idx].value)!="") {
		m2=converte_numero(f.c_vl_custo2[idx].value);
		if (f.c_vl_custo2[idx].value != formata_moeda(m2)) f.c_vl_custo2[idx].value = formata_moeda(m2);
		}
//  DEVIDO A ARRENDODAMENTOS
	m=converte_numero(f.c_vl_unitario[idx].value);
	if (f.c_vl_total[idx].value != formata_moeda(parseInt(f.c_qtde[idx].value) * m)) f.c_vl_total[idx].value = formata_moeda(parseInt(f.c_qtde[idx].value) * m);
	m2=converte_numero(f.c_vl_custo2[idx].value);
	if (f.c_vl_total_custo2[idx].value != formata_moeda(parseInt(f.c_qtde[idx].value) * m2)) f.c_vl_total_custo2[idx].value = formata_moeda(parseInt(f.c_qtde[idx].value) * m2);
	m=0;
	m2=0;
	for (i=0; i<f.c_vl_total.length; i++) 
	{
		m=m+converte_numero(f.c_vl_total[i].value);
		m2=m2+converte_numero(f.c_vl_total_custo2[i].value);
	}
	if (f.c_total_geral.value != formata_moeda(m)) f.c_total_geral.value = formata_moeda(m);
	if (f.c_total_geral_custo2.value != formata_moeda(m2)) f.c_total_geral_custo2.value = formata_moeda(m2);
}

function trataRetornoConsultaIbpt(oResp) {
var f, i, j, strMsg, blnAchou, blnCadastrado, s_ncm_ja_testado, s_ncm_aux;
var blnExecutarSubmit = true;
	f = fESTOQ;
	strMsg = "";
	s_ncm_ja_testado = "";
	for (i = 0; i < f.c_ncm.length; i++) {
		s_ncm_aux = f.c_ncm[i].value;
		if (s_ncm_aux != "") {
			if (s_ncm_ja_testado.indexOf("|" + s_ncm_aux + "|") == -1) {
				s_ncm_ja_testado += "|" + s_ncm_aux + "|";
				blnAchou = false;
				blnCadastrado = false;
				for (j = 0; j < oResp.resposta.length; j++) {
					if (s_ncm_aux == oResp.resposta[j].ncm) {
						blnAchou = true;
						blnCadastrado = oResp.resposta[j].cadastrado;
						break;
					}
				}
				if ((!blnAchou) || (!blnCadastrado)) {
					blnExecutarSubmit = false;
					if (strMsg.length > 0) strMsg += "<br />";
					strMsg += "NCM '" + s_ncm_aux + "' NÃO está cadastrado na tabela do IBPT!!";
				}
			}
		}
	}

	if (blnExecutarSubmit) {
		f.submit();
	}
	else {
		if (strMsg.length > 0) {
			$("#divMsgAlerta div").html(strMsg);
			$("#btnDivMsgConfirmar").show();
			$("#divMsgAlerta").show();
		}
	}
}

function fESTOQConfirma( f ) {
var f, i, s, s_aux, s_produtos_preco_fabricante, s_produtos_vl_custo2, vl_total_custo2, vl_aux, intQtde;
var s_ncm_aux, s_ibpt_ncm, s_ncm_ja_listado;
var s_produtos_ean;
	f=fESTOQ;
	
	vl_total_custo2=0;
	for (i=0; i<f.c_vl_custo2.length; i++) {
		if ((trim(f.c_produto[i].value)!="")&&(trim(f.c_vl_custo2[i].value)=="")) {
			alert("Informe o valor de Referência para o produto " + f.c_produto[i].value);
			f.c_vl_custo2[i].focus();
			return;
			}
		intQtde=converte_numero(f.c_qtde[i].value);
		vl_aux=converte_numero(f.c_vl_custo2[i].value);
		vl_total_custo2=vl_total_custo2+(intQtde*vl_aux);
		}
		
//	DEVIDO A FALHAS DE PRECISÃO DO JAVASCRIPT (EX: 827,85 FICA 827,8499999999999)
	vl_total_custo2=converte_numero(formata_moeda(vl_total_custo2));
	
	f.c_log_edicao.value="";
	s_produtos_preco_fabricante="";
	s_produtos_vl_custo2="";
	for (i=0; i<f.c_vl_unitario.length; i++) {
		if (f.c_vl_unitario[i].value!=f.c_vl_unitario_original[i].value) {
			if (s_produtos_preco_fabricante!="") s_produtos_preco_fabricante=s_produtos_preco_fabricante + ", ";
			s_produtos_preco_fabricante = s_produtos_preco_fabricante + f.c_produto[i].value;
		 // INFORMAÇÕES PARA O LOG
			if (f.c_log_edicao.value!="") f.c_log_edicao.value=f.c_log_edicao.value + "; ";
			f.c_log_edicao.value=f.c_log_edicao.value + f.c_produto[i].value + ": preco_fabricante " + f.c_vl_unitario_original[i].value + "=>" + f.c_vl_unitario[i].value;
			}
		if (f.c_vl_custo2[i].value!=f.c_vl_custo2_original[i].value) {
			if (f.c_vl_custo2_original[i].value!="") {
				if (s_produtos_vl_custo2!="") s_produtos_vl_custo2=s_produtos_vl_custo2 + ", ";
				s_produtos_vl_custo2 = s_produtos_vl_custo2 + f.c_produto[i].value;
				}
		 // INFORMAÇÕES PARA O LOG
			if (f.c_log_edicao.value!="") f.c_log_edicao.value=f.c_log_edicao.value + "; ";
			s_aux = f.c_vl_custo2_original[i].value;
			if (s_aux == "") s_aux = String.fromCharCode(34) + String.fromCharCode(34);
			f.c_log_edicao.value=f.c_log_edicao.value + f.c_produto[i].value + ": vl_custo2 " + s_aux + "=>" + f.c_vl_custo2[i].value;
			}
		}
	s="";
	if (s_produtos_preco_fabricante!="") {
		if (s!="") s=s+"\n\n";
		s = s+"Houve edição no preço de fabricante do(s) seguinte(s) produto(s): " + s_produtos_preco_fabricante;
		}
	if (s_produtos_vl_custo2!="") {
		if (s!="") s=s+"\n\n";
		s = s+"Houve edição no valor de Referência do(s) seguinte(s) produto(s): " + s_produtos_vl_custo2;
		}
	if (s!="") {
		s = s + "\n\n" + "Confirma o cadastramento?";
		if (!confirm(s)) return;
		}
	
	if (f.c_log_edicao.value!="") {
		f.c_log_edicao.value="Cadastramento realizado com edição de valores: " + f.c_log_edicao.value;
		}
		
//  CHECAGEM DO TOTAL DO CUSTO2
	if (trim(f.c_total_custo2_checagem.value)=="") {
		alert("Informe o valor total de Referência para checagem!!");
		f.c_total_custo2_checagem.focus();
		return;
		}

	if (converte_numero(f.c_total_custo2_checagem.value)!=vl_total_custo2) {
		alert("O valor total de Referência não coincide com o valor informado para checagem!!\nVerifique se houve erro de digitação!!");
		return;
		}

	for (i = 0; i < f.c_produto.length; i++) {
		if (trim(f.c_produto[i].value) != "") {
			if (trim(f.c_ncm[i].value) == "") {
				alert("Informe o NCM do produto " + f.c_produto[i].value + "!!");
				f.c_ncm[i].focus();
				return;
			}
			if ((f.c_ncm[i].value.length != 2) && (f.c_ncm[i].value.length != 8)) {
				alert("Tamanho inválido de NCM no produto " + f.c_produto[i].value + "!!");
				f.c_ncm[i].focus();
				return;
			}
			if (trim(f.c_cst[i].value) == "") {
				alert("Informe o CST do produto " + f.c_produto[i].value + "!!");
				f.c_cst[i].focus();
				return;
			}
			if (f.c_cst[i].value.length != 3) {
				alert("Tamanho inválido de CST no produto " + f.c_produto[i].value + "!!");
				f.c_cst[i].focus();
				return;
			}
			if (trim(f.c_ncm[i].value) != trim(f.c_ncm_redigite[i].value)) {
				alert("Falha na conferência do NCM redigitado do produto " + f.c_produto[i].value + "!!");
				f.c_ncm_redigite[i].focus();
				return;
			}
			if (trim(f.c_cst[i].value) != trim(f.c_cst_redigite[i].value)) {
				alert("Falha na conferência do CST redigitado do produto " + f.c_produto[i].value + "!!");
				f.c_cst_redigite[i].focus();
				return;
			}
		}
	}

    // VERIFICAÇÃO DO EAN
	s_produtos_ean="";
	for (i=0; i<f.c_ean.length; i++) {
	    if (f.c_ean[i].value!=f.c_ean_original[i].value) {
	        if (s_produtos_ean!="") s_produtos_ean=s_produtos_ean + ", ";
	        s_produtos_ean = s_produtos_ean + f.c_produto[i].value;
	        // INFORMAÇÕES PARA O LOG
	        if (f.c_log_edicao.value!="") f.c_log_edicao.value=f.c_log_edicao.value + "; ";
	        f.c_log_edicao.value=f.c_log_edicao.value + f.c_produto[i].value + ": ean " + f.c_ean_original[i].value + "=>" + f.c_ean[i].value;
	    }
	}
	s="";
	if (s_produtos_ean!="") {
	    if (s!="") s=s+"\n\n";
	    s = s+"Houve edição no EAN do(s) seguinte(s) produto(s): " + s_produtos_ean;
	}
	if (s!="") {
	    s = s + "\n\n" + "Confirma o cadastramento?";
	    if (!confirm(s)) return;
	}

	$("#dCONFIRMA").hide();
	window.status = "Aguarde ...";

//  VERIFICA SE OS CÓDIGOS DE NCM ESTÃO CADASTRADOS NA TABELA DO IBPT
//  A FUNÇÃO CHAMADA NO CALLBACK IRÁ EXIBIR UMA MENSAGEM NO CASO DE ENCONTRAR
//  CÓDIGOS NÃO CADASTRADOS OU FAZER O SUBMIT() SE ESTIVER TUDO OK.
	s_ibpt_ncm = "";
	s_ncm_ja_listado = "";
	for (i = 0; i < f.c_ncm.length; i++) {
		s_ncm_aux = f.c_ncm[i].value;
		if (s_ncm_aux != "") {
			if (s_ncm_ja_listado.indexOf("|" + s_ncm_aux + "|") == -1) {
				s_ncm_ja_listado += "|" + s_ncm_aux + "|";
				if (s_ibpt_ncm.length > 0) s_ibpt_ncm += ",";
				s_ibpt_ncm += s_ncm_aux;
			}
		}
	}

	$("#divAjaxProgress").show();
	$.getJSON(
		"../Global/IbptNcmConsultaBD.asp",
		{ ncm: s_ibpt_ncm },
		trataRetornoConsultaIbpt)
	.fail(function(jqXHR, textStatus, errorThrown) {
		$("#divAjaxProgress").hide();
		$("#dCONFIRMA").show();
		$("#divDialogBox div").html("<b>Ocorreu um erro ao tentar consultar os dados no servidor!!<br />Por favor, tente novamente.</b>" + "<br /><br /><b><i>Status:</i></b><br />" + textStatus + "<br /><br /><b><i>Erro ocorrido:</i></b><br />" + errorThrown + "<br /><br /><b><i>Descrição do erro:</i></b><br />" + jqXHR.responseText);
		$("#divDialogBox").dialog("option", "title", "Erro!");
		$("#divDialogBox").dialog("open");
	})
	.always(function() { $("#divAjaxProgress").hide() });
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
<link href="<%=URL_FILE__JQUERY_UI_CSS%>" rel="stylesheet" type="text/css">

<style type="text/css">
#ckb_especial_aux {
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
.divMsgAlerta
{
	margin-top:30px;
	margin-bottom:10px;
	border: solid 2px #000000;
	font-weight: bold;
	text-align: center;
	padding: 10px;
	width: 760px;
	color: #EF0000;
	background-color: #FFFFC4;
}
.divAjaxProgress
{
	margin-top:15px;
	text-align: center;
	vertical-align: middle;
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
<!-- ************************************************************* -->
<!-- **********  PÁGINA PARA EXIBIR DADOS DOS PRODUTOS  ********** -->
<!-- ************************************************************* -->
<body>
<center>

<form id="fESTOQ" name="fESTOQ" method="post" action="EstoqueEntradaConfirmaEAN.asp" autocomplete="off">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="c_log_edicao" id="c_log_edicao" value="">
<input type="hidden" name="c_fabricante" id="c_fabricante" value="<%=s_fabricante%>">
<input type="hidden" name="c_id_nfe_emitente" id="c_id_nfe_emitente" value="<%=s_id_nfe_emitente%>">
<!-- É NECESSÁRIO CRIAR UM CAMPO DO TIPO HIDDEN PARA QUE A PÁGINA SEGUINTE CONSIGA
	 RECUPERAR A INFORMAÇÃO REFERENTE A ESTE CAMPO, JÁ QUE REQUEST.FORM() EM UM
	 CAMPO DO TIPO CHECKBOX QUE ESTÁ DISABLED RETORNA VAZIO.
-->
<input type="hidden" name="ckb_especial" id="ckb_especial" value="<%=ckb_especial%>">

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="780" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">

<tr>
	<td align="right" valign="bottom"><p class="PEDIDO">Entrada de Mercadorias no Estoque<span class="C">&nbsp;</span></p></td>
</tr>
</table>
<br>

<!--  CADASTRO DA ENTRADA DE MERCADORIAS NO ESTOQUE  -->
<table class="Qx" cellspacing="0" cellpadding="0">
<!--  EMPRESA  -->
	<tr class="trWmsCd">
		<td>
			<table width="100%" cellpadding="0" cellspacing="0">
				<tr>
					<td class="MT" align="left">
						<span class="PLTe">Empresa</span>
						<br />
						<span class="PLLe" style="margin-left:2pt;"><%=obtem_apelido_empresa_NFe_emitente(s_id_nfe_emitente)%></span>
					</td>
				</tr>
			</table>
		</td>
	</tr>
<!--  FABRICANTE -->
	<tr bgcolor="#FFFFFF"><td class="MDBE" align="left" nowrap><span class="PLTe">Fabricante</span>
		<%	s = s_fabricante
			if (s<>"") And (s_nome_fabricante<>"") then s = s & " - " & s_nome_fabricante %>
		<br><input name="c_fabricante_aux" id="c_fabricante_aux" readonly tabindex=-1 class="PLLe" style="width:460px;margin-left:2pt;" 
				value="<%=s%>"></td></tr>
<!-- ÁGIO  -->
	<tr bgcolor="#FFFFFF"><td class="MDBE" align="left" nowrap><span class="PLTe">% Ágio</span>
		<br><input name="c_perc_agio" id="c_perc_agio" class="PLLe TxtEditavel" maxlength="8" value="" 
            onkeypress="if (digitou_enter(true)) $('#c_fabricante').focus();" 
            onblur="this.value=formata_numero(this.value, 4); recalcula_itens();"></td></tr>
<!--  DOCUMENTO  -->
	<tr bgcolor="#FFFFFF"><td class="MDBE" align="left" nowrap><span class="PLTe">Documento</span>
		<br><input name="c_documento" id="c_documento" readonly tabindex=-1 class="PLLe" style="width:270px;margin-left:2pt;"
			value="<%=s_documento%>"></td></tr>
<!--  ENTRADA ESPECIAL  -->
	<tr bgcolor="#FFFFFF">
	<td class="MDBE" align="left" nowrap><span class="PLTe">Tipo de Cadastramento</span>
		<br><input type="checkbox" class="rbOpt" disabled tabindex="-1" id="ckb_especial_aux" name="ckb_especial_aux" value=""
		<%if ckb_especial <> "" then Response.Write " checked" %>
		><span class="C lblOpt" style="cursor:default">Entrada Especial</span>
	</td></tr>
<!--  OBS  -->
	<tr bgcolor="#FFFFFF">
	<td class="MDBE" align="left" nowrap><span class="PLTe">Observações</span>
		<br><textarea name="c_obs" id="c_obs" class="PLLe" rows="<%=Cstr(MAX_LINHAS_ESTOQUE_OBS)%>"
				style="width:642px;margin-left:2pt;" onkeypress="limita_tamanho(this,MAX_TAM_T_ESTOQUE_CAMPO_OBS);" onblur="this.value=trim(this.value);"
				readonly tabindex=-1><%=s_obs%></textarea>
	</td>
	</tr>
</table>
<br>

<!--  R E L A Ç Ã O   D E   P R O D U T O S  -->
<table id="tableProduto" class="Qx" cellspacing="0" border="0">
	<thead>
	<tr bgcolor="#FFFFFF">
	<th>&nbsp;</th>
	<th class="MB" align="left" valign="bottom"><span class="PLTe">Produto</span></th>
	<th class="MB" align="left" valign="bottom"><span class="PLTe">EAN</span></th>
	<th class="MB" align="left" valign="bottom"><span class="PLTe">Descrição</span></th>
	<th class="MB" align="center" valign="bottom" align="center"><span class="PLTe">NCM</span></th>
	<th class="MB" align="center" valign="bottom" align="center"><span class="PLTe">NCM<br /><span style="font-size:7pt;">(redigite)</span></span></th>
	<th class="MB" align="center" valign="bottom" align="center"><span class="PLTe">CST</span></th>
	<th class="MB" align="center" valign="bottom" align="center"><span class="PLTe">CST<br /><span style="font-size:7pt;">(redigite)</span></span></th>
	<th class="MB" align="right" valign="bottom"><span class="PLTd">Qtde</span></th>
	<th class="MB" align="right" valign="bottom" style="width:62px;"><span class="PLTd">Base Cálc ICMS ST (Unit)</span></th>
	<th class="MB" align="right" valign="bottom" style="width:62px;"><span class="PLTd">Valor ICMS ST (Unit)</span></th>
	<th class="MB" align="right" valign="bottom" style="width:62px;"><span class="PLTd">Preço Fabric</span></th>
	<th class="MB" align="right" valign="bottom" style="width:62px;"><span class="PLTd">Total Fabric</span></th>
	<th class="MB" align="right" valign="bottom" style="width:62px;"><span class="PLTd">Valor Referência</span></th>
    <th class="MB" align="right" valign="bottom" style="width:62px;"><span class="PLTd">A. IPI</span></th>
    <th class="MB" align="right" valign="bottom" style="width:62px;"><span class="PLTd">V. IPI</span></th>
    <th class="MB" align="right" valign="bottom" style="width:62px;"><span class="PLTd">A. ICMS</span></th>
	<th class="MB" align="right" valign="bottom" style="width:62px;"><span class="PLTd">Total Valor Referência</span></th>
    <th class="MB" align="right" valign="bottom" style="width:62px;"><span class="PLTd">Total Diferença</span></th>

	</tr>
	</thead>

	<tbody>
<%	m_total_geral=0
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
			'	PREÇO FABRICANTE
				s_vl_unitario = formata_moeda(.preco_fabricante)
				m_vl_total = .qtde * .preco_fabricante
				s_vl_total=formata_moeda(m_vl_total)
				m_total_geral=m_total_geral + m_vl_total
			'	CUSTO 2
				if .vl_custo2 = 0 then s_vl_custo2 = "" else s_vl_custo2 = formata_moeda(.vl_custo2)
				m_vl_total_custo2 = .qtde * .vl_custo2
				s_vl_total_custo2=formata_moeda(m_vl_total_custo2)
				m_total_geral_custo2=m_total_geral_custo2 + m_vl_total_custo2
				end with
		else
			exit for
			end if
%>
	<tr>
	<td align="left"><input name="c_linha" readonly tabindex=-1 class="PLLe" maxlength="2" style="width:24px;text-align:right;color:#808080;" value="<%=Cstr(i) & ". " %>"></td>
	<td class="MDBE" align="left">
		<input name="c_produto" readonly tabindex=-1 class="PLLe" style="width:50px;"
			value="<%=s_produto%>"></td>
	<td class="MDB" align="left">
		<input name="c_ean" class="PLLe" style="width:80px;"
            onkeypress="if (digitou_enter(true)) $(this).hUtil('focusNext'); filtra_numerico();"
            onblur="this.value=retorna_so_digitos(this.value); if (retorna_so_digitos(this.value)=="") this.value='<%=s_ean%>';"
            value="<%=s_ean%>">
                <!--  CONTROLA SE HOUVE EDIÇÃO -->
		<input type=hidden name="c_ean_original" value="<%=s_ean%>">
    	</td>

	<td class="MDB" align="left">
		<span class="PLLe" style="width:240px;"><%=s_descricao_html%></span>
		<input type=hidden name="c_descricao" value="<%=s_descricao%>">
	</td>
	<td class="MDB" align="left">
		<input name="c_ncm" class="PLLc" maxlength="8" style="width:56px;"
			onkeypress="if (digitou_enter(true)) $(this).hUtil('focusNext'); filtra_numerico();">
	</td>
	<td class="MDB" align="left">
		<input name="c_ncm_redigite" class="PLLc" maxlength="8" style="width:56px;"
			onkeypress="if (digitou_enter(true)) $(this).hUtil('focusNext'); filtra_numerico();"
			onpaste="return cancela_onpaste()">
	</td>
	<td class="MDB" align="left">
		<input name="c_cst" class="PLLc" maxlength="3" style="width:40px;"
			onkeypress="if (digitou_enter(true)) $(this).hUtil('focusNext'); filtra_numerico();">
	</td>
	<td class="MDB" align="left">
		<input name="c_cst_redigite" class="PLLc" maxlength="3" style="width:40px;"
			onkeypress="if (digitou_enter(true)) $(this).hUtil('focusNext'); filtra_numerico();"
			onpaste="return cancela_onpaste()">
	</td>
	<td class="MDB" align="right">
		<input name="c_qtde" readonly tabindex=-1 class="PLLd" style="width:30px;"
			value="<%=s_qtde%>"></td>
	<td class="MDB" align="right">
		<input name="c_vl_BC_ICMS_ST" class="PLLd" maxlength="12" style="width:62px;"
			onkeypress="if (digitou_enter(true)) $(this).hUtil('focusNext'); filtra_moeda();"
			onblur="this.value=formata_moeda(this.value); if (converte_numero(this.value)<0) {alert('Valor inválido!!');this.focus();}"
			>
		</td>
	<td class="MDB" align="right">
		<input name="c_vl_ICMS_ST" class="PLLd" maxlength="12" style="width:62px;"
			onkeypress="if (digitou_enter(true)) $(this).hUtil('focusNext'); filtra_moeda();"
			onblur="this.value=formata_moeda(this.value); if (converte_numero(this.value)<0) {alert('Valor inválido!!');this.focus();}"
			>
		</td>
	<td class="MDB" align="right">
		<input name="c_vl_unitario" class="PLLd" maxlength="12" style="width:62px;"
			onkeypress="if (digitou_enter(true)) $(this).hUtil('focusNext'); filtra_moeda();"
			onblur="this.value=formata_moeda(this.value); if (converte_numero(this.value)<0) {alert('Valor inválido!!');this.focus();} else {recalcula_itens(); recalcula_total_nf(); recalcula_total(<%=Cstr(i)%>);}"
			value="<%=s_vl_unitario%>">
		<!--  CONTROLA SE HOUVE EDIÇÃO -->
		<input type=hidden name="c_vl_unitario_original" value="<%=s_vl_unitario%>">
		</td>

	<td class="MDB" align="right">
		<input name="c_vl_total" readonly tabindex=-1 class="PLLd" style="width:62px;"
			value="<%=s_vl_total%>"></td>

	<td class="MDB" align="right">
		<input name="c_vl_custo2" class="PLLd" maxlength="12" style="width:62px;"
			onkeypress="if (digitou_enter(true)) $(this).hUtil('focusNext'); filtra_moeda();"
			onblur="if (trim(this.value)!='') this.value=formata_moeda(this.value); if (converte_numero(this.value)<0) {alert('Valor inválido!!');this.focus();} else {recalcula_total(<%=Cstr(i)%>); recalcula_diferenca(<%=Cstr(i)%>);}"
			value="<%=s_vl_custo2%>">
		<!--  CONTROLA SE HOUVE EDIÇÃO -->
		<input type=hidden name="c_vl_custo2_original" value="<%=s_vl_custo2%>">
		</td>

	<td class="MDB" align="right">
		<input name="c_aliq_ipi" class="PLLd" maxlength="12" style="width:62px;"
			onkeypress="if (digitou_enter(true)) $(this).hUtil('focusNext'); filtra_numerico();"
			onblur="this.value=formata_numero(this.value, 0); if (converte_numero(this.value)<0) {alert('Valor inválido!!');this.focus();}"
			>
		</td>

	<td class="MDB" align="right">
		<input name="c_vl_ipi" class="PLLd" maxlength="12" style="width:62px;"
			onkeypress="if (digitou_enter(true)) $(this).hUtil('focusNext'); filtra_moeda();"
			onblur="this.value=formata_moeda(this.value); if (converte_numero(this.value)<0) {alert('Valor inválido!!');this.focus();}"
			>
		</td>

	<td class="MDB" align="right">
		<input name="c_aliq_icms" class="PLLd" maxlength="12" style="width:62px;"
			onkeypress="if (digitou_enter(true)) $(this).hUtil('focusNext'); filtra_numerico();"
			onblur="this.value=formata_numero(this.value, 0); if (converte_numero(this.value)<0) {alert('Valor inválido!!');this.focus();}"
			>
		</td>


	<td class="MDB" align="right">
		<input name="c_vl_total_custo2" readonly tabindex=-1 class="PLLd" style="width:62px;"
			value="<%=s_vl_total_custo2%>"></td>

    <td class="MDB" align="right">
		<input name="c_vl_total_diferenca" readonly tabindex=-1 class="PLLd" style="width:62px;"
			value=""></td>

	</tr>
<% next %>
	</tbody>

	<tfoot>
	<tr>
    <td colspan="18" class="MD" id="tdPreTotalGeralFabricante">&nbsp;</td>
    
	<td class="MD">&nbsp;</td>
    <td class="MD" align="left"><p class="Cd">Total NF</p></td>
	<td class="MDB" align="right"><input name="c_total_nf" id="c_total_nf" class="PLLd" style="width:62px;color:blue;" 
		value='' readonly tabindex=-1></td>
    <!--O CAMPO ABAIXO DEIXOU DE APRESENTAR O TOTAL DOS PRODUTOS E PASSOU A APRESENTAR O TOTAL DA NOTA FISCAL, POR SOLICITAÇÃO, PARA FACILITAR A VISUALIZAÇÃO NA CONSULTA-->
    <!--<td class="MDB" align="right" id="tdTotalGeralFabricante"><input name="c_total_geral" id="c_total_geral" class="PLLd" style="width:62px;color:blue;" 
		value='<%=formata_moeda(m_total_geral)%>' readonly tabindex=-1></td>-->
	<td class="MDB" align="right" id="tdTotalGeralFabricante"><input name="c_total_geral" id="c_total_geral" class="PLLd" style="width:62px;color:blue;" 
		value='<%=formata_moeda(m_total_geral_custo2)%>' readonly tabindex=-1></td>
    <td class="MD">&nbsp;</td>
    <td class="MD">&nbsp;</td>
    <td class="MD">&nbsp;</td>
	<td class="MDB" align="right"><input name="c_total_geral_custo2" id="c_total_geral_custo2" class="PLLd" style="width:62px;color:blue;" 
		value='<%=formata_moeda(m_total_geral_custo2)%>' readonly tabindex=-1></td>
	<td class="MDB" align="right" id="tdTotalGeralDiferenca"><input name="c_total_geral_diferenca" id="c_total_geral_diferenca" class="PLLd" style="width:62px;color:blue;" 
		value='' readonly tabindex=-1></td>
	</tr>
	<tr>
		<td colspan="18">&nbsp;</td>
	</tr>
	<tr>
	<td colspan="18" id="tdPreChecagem">&nbsp;</td>
	<td class="MD" align="left"><p class="Cd">Checagem</p></td>
	<td class="MDB MC" align="right"><input name="c_total_custo2_checagem" id="c_total_custo2_checagem" class="PLLd" style="width:62px;color:black;" 
		onkeypress="if (digitou_enter(true)&&tem_info(this.value)) bCONFIRMA.focus(); filtra_moeda();" 
		onblur="if (trim(this.value)!='') this.value=formata_moeda(this.value); if (converte_numero(this.value)<0) {alert('Valor inválido!!');this.focus();}"
		onpaste="return cancela_onpaste()"
		value=''></td>
	</tr>
	</tfoot>
</table>


<!--  ASSEGURA CRIAÇÃO DE UM ARRAY DE CAMPOS, MESMO QUANDO HOUVER SOMENTE 1 PRODUTO!! -->
<input type=HIDDEN name="c_linha" value="">
<input type=HIDDEN name="c_produto" value="">
<input type=HIDDEN name="c_ean" value="">
<input type=HIDDEN name="c_ean_original" value="">
<input type=HIDDEN name="c_descricao" value="">
<input type=HIDDEN name="c_ncm" value="">
<input type=HIDDEN name="c_ncm_redigite" value="">
<input type=HIDDEN name="c_cst" value="">
<input type=HIDDEN name="c_cst_redigite" value="">
<input type=HIDDEN name="c_qtde" value="">
<input type=HIDDEN name="c_vl_BC_ICMS_ST" value="">
<input type=HIDDEN name="c_vl_ICMS_ST" value="">
<input type=HIDDEN name="c_vl_unitario" value="">
<input type=HIDDEN name="c_vl_unitario_original" value="">
<input type=HIDDEN name="c_vl_total" value="">
<input type=HIDDEN name="c_vl_custo2" value="">
<input type=HIDDEN name="c_vl_custo2_original" value="">
<input type=HIDDEN name="c_vl_total_custo2" value="">
<input type=HIDDEN name="c_vl_diferenca" value="">
<input type=HIDDEN name="c_vl_total_diferenca" value="">
<input type=HIDDEN name="c_aliq_ipi" value="">
<input type=HIDDEN name="c_aliq_icms" value="">
<input type=HIDDEN name="c_vl_ipi" value="">

<!--  AJAX PROGRESS GIF -->
<div id="divAjaxProgress" class="divAjaxProgress">
<img src="../imagem/ajax_loader_gray_128.gif" alt="Requisição ajax em andamento" />
</div>

<!--  MENSAGEM DE ALERTA SOBRE POSSÍVEIS PROBLEMAS DE CONSISTÊNCIA -->
<div id="divMsgAlerta" class="divMsgAlerta">
<div></div>
<br />
<center>
<table style="width:600px;">
<tr>
	<td align="left">
		<button id="btnDivMsgCancelar"> &nbsp;&nbsp; CANCELAR &nbsp;&nbsp; </button>
	</td>
	<td align="right">
		<button id="btnDivMsgConfirmar"> &nbsp; Confirmar cadastramento &nbsp; </button>
	</td>
</tr>
</table>
</center>
</div>

<!--  DIV P/ DIALOG BOX -->
<div id="divDialogBox">
<div></div>
</div>


<!-- ************   SEPARADOR   ************ -->
<table width="780" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>
<br>


<table width="780" cellSpacing="0">
<tr>
	<td align="left"><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()" title="volta para página anterior">
		<img src="../botao/anterior.gif" width="176" height="55" border="0"></a></td>
	<td align="right"><div name="dCONFIRMA" id="dCONFIRMA">
	<a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fESTOQConfirma(fESTOQ)" title="confirma a entrada das mercadorias no estoque">
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