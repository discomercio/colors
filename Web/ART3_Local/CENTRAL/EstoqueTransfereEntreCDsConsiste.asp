<%@  language="VBScript" %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/estoque.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ====================================================
'	  EstoqueTransfereEntreCDsConsiste.asp
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

'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs, msg_erro
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

	dim alerta
	alerta = ""

	dim s, i, j, n
	dim c_nfe_emitente_origem, c_nfe_emitente_destino ', c_obs
	dim c_fabricante, c_produto, c_qtde
	dim v_item
    dim v_item_transf
	dim cod_fabricante_aux

	c_nfe_emitente_origem = Trim(Request.Form("c_nfe_emitente_origem"))
	c_nfe_emitente_destino = Trim(Request.Form("c_nfe_emitente_destino"))
	'c_obs = Trim(Request.Form("c_obs"))

	redim v_item(0)
	set v_item(UBound(v_item)) = New cl_ITEM_TRANSF_ENTRE_CD
	v_item(UBound(v_item)).produto = ""

    redim v_item_transf(0)
	set v_item_transf(UBound(v_item_transf)) = New cl_ESTOQUE_TRANSFERENCIA_ITEM_SUB
	v_item_transf(UBound(v_item_transf)).produto = ""

	n = Request.Form("c_produto").Count
	for i = 1 to n
		c_fabricante = retorna_so_digitos(Trim(Request.Form("c_fabricante")(i)))
		if c_fabricante <> "" then c_fabricante = normaliza_codigo(c_fabricante, TAM_MIN_FABRICANTE)
		c_produto = retorna_so_digitos(Trim(Request.Form("c_produto")(i)))
		if c_produto <> "" then c_produto = normaliza_codigo(c_produto, TAM_MIN_PRODUTO)
		c_qtde = retorna_so_digitos(Trim(Request.Form("c_qtde")(i)))
		if c_produto <> "" then
			if Trim("" & v_item(Ubound(v_item)).produto) <> "" then
				redim preserve v_item(ubound(v_item)+1)
				set v_item(UBound(v_item)) = New cl_ITEM_TRANSF_ENTRE_CD
				end if
			
			with v_item(UBound(v_item))
				.fabricante = c_fabricante
				.produto = c_produto
				.qtde = converte_numero(c_qtde)
				end with
			end if
		next

	if alerta = "" then
		if converte_numero(c_nfe_emitente_origem) = 0 then
			alerta=texto_add_br(alerta)
			alerta=alerta & "É necessário informar o CD de origem."
			end if

		if converte_numero(c_nfe_emitente_destino) = 0 then
			alerta=texto_add_br(alerta)
			alerta=alerta & "É necessário informar o CD de destino."
			end if
		end if

	if alerta = "" then
		if converte_numero(c_nfe_emitente_origem) = converte_numero(c_nfe_emitente_destino) then
			alerta=texto_add_br(alerta)
			alerta=alerta & "O CD de destino da transferência não pode ser o mesmo da origem."
			end if
		end if

'	if alerta = "" then
'		if Len(c_obs) > MAX_TAM_T_ESTOQUE_CAMPO_OBS then
'			alerta=texto_add_br(alerta)
'			alerta=alerta & "O conteúdo do campo 'Observações' excede o tamanho máximo de " & Cstr(MAX_TAM_T_ESTOQUE_CAMPO_OBS) & " caracteres."
'			end if
'		end if

'	VERIFICA LISTA DE PRODUTOS
	if alerta = "" then
		for i=LBound(v_item) to UBound(v_item)
			if Trim(v_item(i).produto) <> "" then
				if Trim(v_item(i).fabricante) = "" then
				'	SE INFORMOU APENAS O CÓDIGO DO PRODUTO, TENTA LOCALIZAR O CÓDIGO DO FABRICANTE
					s = "SELECT fabricante, produto FROM t_PRODUTO WHERE (produto = '" & v_item(i).produto & "') AND (excluido_status = 0)"
					if rs.State <> 0 then rs.Close
					rs.open s, cn
					if rs.Eof then
						alerta=texto_add_br(alerta)
						alerta=alerta & "Produto " & v_item(i).produto & " não está cadastrado."
					else
						cod_fabricante_aux = Trim("" & rs("fabricante"))
						rs.MoveNext
						if rs.Eof then
							v_item(i).fabricante = cod_fabricante_aux
						else
							alerta=texto_add_br(alerta)
							alerta=alerta & "Produto " & v_item(i).produto & ": não foi possível determinar o código do fabricante porque há mais de um produto cadastrado com este mesmo código."
							end if
						end if
				else
				'	VERIFICA SE O PRODUTO ESTÁ CADASTRADO
					s = "SELECT fabricante, produto FROM t_PRODUTO WHERE (fabricante = '" & v_item(i).fabricante & "') AND (produto = '" & v_item(i).produto & "') AND (excluido_status = 0)"
					if rs.State <> 0 then rs.Close
                    rs.open s, cn
					if rs.Eof then
						alerta=texto_add_br(alerta)
						alerta=alerta & "Produto " & v_item(i).produto & " do fabricante " & v_item(i).fabricante & " não está cadastrado."
						end if
					end if

			'	VERIFICA SE A QUANTIDADE ESPECIFICADA É VÁLIDA
				if v_item(i).qtde <= 0 then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Produto " & v_item(i).produto & " do fabricante " & v_item(i).fabricante & ": quantidade " & CStr(v_item(i).qtde) & " é inválida."
					end if
				end if
			next
		end if

	if alerta = "" then
		for i=LBound(v_item) to UBound(v_item)
			if Trim(v_item(i).produto) <> "" then
			'	VERIFICA SE O PRODUTO ESTÁ REPETIDO NA LISTA (ESTA VERIFICAÇÃO DEVE OCORRER APÓS A ROTINA QUE RECUPERA O CÓDIGO DO FABRICANTE P/ OS ITENS EM QUE FOI INFORMADO SOMENTE O CÓDIGO DO PRODUTO)
				for j=Lbound(v_item) to (i-1)
					if (v_item(i).produto = v_item(j).produto) And (v_item(i).fabricante = v_item(j).fabricante) then
						alerta=texto_add_br(alerta)
						alerta=alerta & "Produto " & v_item(i).produto & " do fabricante " & v_item(i).fabricante & ": linha " & renumera_com_base1(Lbound(v_item),i) & " repete o mesmo produto da linha " & renumera_com_base1(Lbound(v_item),j) & "."
						exit for
						end if
					next
				end if
			next
		end if




    dim s_fabricante, s_nome_fabricante, s_documento, ckb_especial, s_obs
	dim s_produto, s_ean, s_descricao_html, s_qtde, s_vl_unitario, s_vl_total, m_vl_total, m_total_geral
	'dim s_vl_custo2, s_vl_total_custo2, m_vl_total_custo2, m_total_geral_custo2
	dim s_nome_nfe_emitente
    dim s_ncm, s_cst, s_preco_fabricante, s_vl_custo2, s_vl_BC_ICMS_ST, s_vl_ICMS_ST, s_st_ncm_cst_herdado_tabela_produto
    dim s_aliq_ipi, s_aliq_icms, s_vl_ipi
    dim s_preco_origem, s_produto_xml, s_entrada_tipo, s_id_estoque_origem
    dim s_nfe_entrada_numero

'	VERIFICA SE HÁ DISPONIBILIDADE NO ESTOQUE DE VENDA NO CD DE ORIGEM
	if alerta = "" then    
		for i=LBound(v_item) to UBound(v_item)
			if Trim(v_item(i).produto) <> "" then
                if not estoque_produto_transf_cd_monta(c_nfe_emitente_origem, _
                                                        v_item(i).fabricante, _
                                                        v_item(i).produto, _
                                                        v_item(i).qtde, _
                                                        v_item_transf, _
                                                        msg_erro) then
                    alerta = "Problemas na montagem da transferência: " & msg_erro
                    exit for
                    end if
                end if
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

    $(document).ready(function () {
        $("#divMsgAlerta").hide();
        $("#divAjaxProgress").hide();
        $("#divDialogBox").hide();
        $("#divDialogBox").hUtilUI('dialog_modal');

        $("#btnDivMsgCancelar").button().click(function (event) {
            event.preventDefault();
            $("#divMsgAlerta").hide();
            $("#dCONFIRMA").show();
        });

        $("#btnDivMsgConfirmar").button().click(function (event) {
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
            // 15 - Total Valor Referência
            //$("#tableProduto thead th:nth-child(3), #tableProduto tbody td:nth-child(3)").hide();
            //$("#tableProduto thead th:nth-child(10), #tableProduto tbody td:nth-child(10)").hide();
            //$("#tableProduto thead th:nth-child(11), #tableProduto tbody td:nth-child(11)").hide();
            //$("#tableProduto thead th:nth-child(13), #tableProduto tbody td:nth-child(13)").hide();
            //$("#tdTotalGeralFabricante").hide();
            //$("#tdPreTotalGeralFabricante").removeClass("MD").attr("colSpan", 9);
            $("#tdPreChecagem").attr("colSpan", 9);
        $("input:text:enabled:visible:not([readonly])").focus(function () {
            $(this).select();
        });
        $("input:text:enabled:visible:not([readonly]):first").focus();

        // Tratamento p/ bug do jQuery-ui Dialog: ao tentar mover o dialog em uma tela que está c/ scroll
        // vertical, o dialog é "redesenhado" mais abaixo da posição do cursor na mesma medida do deslocamento do
        // scroll vertical. A movimentação do dialog ocorre c/ esse espaço em branco entre o cursor e o dialog.
        $(document).scroll(function (e) {
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

    function calcula_tamanho_restante() {
        var f, s;
        f = fESTOQ;
        s = "" + f.c_obs.value;
        //f.c_tamanho_restante.value = "<%=MAX_TAM_T_ESTOQUE_CAMPO_OBS%>" - s.length;
        f.c_tamanho_restante.value = 500 - s.length;
    }

    function cancela_onpaste() {
        return false;
    }

    function recalcula_total(id) {
        var idx, m, m2, f, i;
        f = fESTOQ;
        idx = parseInt(id) - 1;
        if (f.c_produto[idx].value == "") return;
        m = converte_numero(f.c_vl_unitario[idx].value);
        if (f.c_vl_unitario[idx].value != formata_moeda(m)) f.c_vl_unitario[idx].value = formata_moeda(m);
        if (trim(f.c_vl_custo2[idx].value) != "") {
            m2 = converte_numero(f.c_vl_custo2[idx].value);
            if (f.c_vl_custo2[idx].value != formata_moeda(m2)) f.c_vl_custo2[idx].value = formata_moeda(m2);
        }
        //  DEVIDO A ARRENDODAMENTOS
        m = converte_numero(f.c_vl_unitario[idx].value);
        if (f.c_vl_total[idx].value != formata_moeda(parseInt(f.c_qtde[idx].value) * m)) f.c_vl_total[idx].value = formata_moeda(parseInt(f.c_qtde[idx].value) * m);
        m2 = converte_numero(f.c_vl_custo2[idx].value);
        if (f.c_vl_total_custo2[idx].value != formata_moeda(parseInt(f.c_qtde[idx].value) * m2)) f.c_vl_total_custo2[idx].value = formata_moeda(parseInt(f.c_qtde[idx].value) * m2);
        m = 0;
        m2 = 0;
        for (i = 0; i < f.c_vl_total.length; i++) {
            m = m + converte_numero(f.c_vl_total[i].value);
            m2 = m2 + converte_numero(f.c_vl_total_custo2[i].value);
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

    function fESTOQConfirma(f) {
        var f, i, s, s_aux, s_produtos_preco_fabricante, s_produtos_vl_custo2, vl_total_custo2, vl_aux, intQtde;
        var s_ncm_aux, s_ibpt_ncm, s_ncm_ja_listado;
        f = fESTOQ;

        //alert("Em desenvolvimento!")
        //return;

        //vl_total_custo2=0;
        //for (i=0; i<f.c_vl_custo2.length; i++) {
        //    if ((trim(f.c_produto[i].value)!="")&&(trim(f.c_vl_custo2[i].value)=="")) {
        //        alert("Informe o valor de Referência para o produto " + f.c_produto[i].value);
        //        f.c_vl_custo2[i].focus();
        //        return;
        //    }
        //    intQtde=converte_numero(f.c_qtde[i].value);
        //    vl_aux=converte_numero(f.c_vl_custo2[i].value);
        //    vl_total_custo2=vl_total_custo2+(intQtde*vl_aux);
        //}

        ////	DEVIDO A FALHAS DE PRECISÃO DO JAVASCRIPT (EX: 827,85 FICA 827,8499999999999)
        //vl_total_custo2=converte_numero(formata_moeda(vl_total_custo2));

        //f.c_log_edicao.value="";
        //s_produtos_preco_fabricante="";
        //s_produtos_vl_custo2="";
        //for (i=0; i<f.c_vl_unitario.length; i++) {
        //    if (f.c_vl_unitario[i].value!=f.c_vl_unitario_original[i].value) {
        //        if (s_produtos_preco_fabricante!="") s_produtos_preco_fabricante=s_produtos_preco_fabricante + ", ";
        //        s_produtos_preco_fabricante = s_produtos_preco_fabricante + f.c_produto[i].value;
        //        // INFORMAÇÕES PARA O LOG
        //        if (f.c_log_edicao.value!="") f.c_log_edicao.value=f.c_log_edicao.value + "; ";
        //        f.c_log_edicao.value=f.c_log_edicao.value + f.c_produto[i].value + ": preco_fabricante " + f.c_vl_unitario_original[i].value + "=>" + f.c_vl_unitario[i].value;
        //    }
        //    if (f.c_vl_custo2[i].value!=f.c_vl_custo2_original[i].value) {
        //        if (f.c_vl_custo2_original[i].value!="") {
        //            if (s_produtos_vl_custo2!="") s_produtos_vl_custo2=s_produtos_vl_custo2 + ", ";
        //            s_produtos_vl_custo2 = s_produtos_vl_custo2 + f.c_produto[i].value;
        //        }
        //        // INFORMAÇÕES PARA O LOG
        //        if (f.c_log_edicao.value!="") f.c_log_edicao.value=f.c_log_edicao.value + "; ";
        //        s_aux = f.c_vl_custo2_original[i].value;
        //        if (s_aux == "") s_aux = String.fromCharCode(34) + String.fromCharCode(34);
        //        f.c_log_edicao.value=f.c_log_edicao.value + f.c_produto[i].value + ": vl_custo2 " + s_aux + "=>" + f.c_vl_custo2[i].value;
        //    }
        //}
        //s="";
        //if (s_produtos_preco_fabricante!="") {
        //    if (s!="") s=s+"\n\n";
        //    s = s+"Houve edição no preço de fabricante do(s) seguinte(s) produto(s): " + s_produtos_preco_fabricante;
        //}
        //if (s_produtos_vl_custo2!="") {
        //    if (s!="") s=s+"\n\n";
        //    s = s+"Houve edição no valor de Referência do(s) seguinte(s) produto(s): " + s_produtos_vl_custo2;
        //}
        //if (s!="") {
        //    s = s + "\n\n" + "Confirma o cadastramento?";
        //    if (!confirm(s)) return;
        //}

        //if (f.c_log_edicao.value!="") {
        //    f.c_log_edicao.value="Cadastramento realizado com edição de valores: " + f.c_log_edicao.value;
        //}

        ////  CHECAGEM DO TOTAL DO CUSTO2
        //if (trim(f.c_total_custo2_checagem.value)=="") {
        //    alert("Informe o valor total de Referência para checagem!!");
        //    f.c_total_custo2_checagem.focus();
        //    return;
        //}

        //if (converte_numero(f.c_total_custo2_checagem.value)!=vl_total_custo2) {
        //    alert("O valor total de Referência não coincide com o valor informado para checagem!!\nVerifique se houve erro de digitação!!");
        //    return;
        //}

        //for (i = 0; i < f.c_produto.length; i++) {
        //    if (trim(f.c_produto[i].value) != "") {
        //        if (trim(f.c_ncm[i].value) == "") {
        //            alert("Informe o NCM do produto " + f.c_produto[i].value + "!!");
        //            f.c_ncm[i].focus();
        //            return;
        //        }
        //        if ((f.c_ncm[i].value.length != 2) && (f.c_ncm[i].value.length != 8)) {
        //            alert("Tamanho inválido de NCM no produto " + f.c_produto[i].value + "!!");
        //            f.c_ncm[i].focus();
        //            return;
        //        }
        //        if (trim(f.c_cst[i].value) == "") {
        //            alert("Informe o CST do produto " + f.c_produto[i].value + "!!");
        //            f.c_cst[i].focus();
        //            return;
        //        }
        //        if (f.c_cst[i].value.length != 3) {
        //            alert("Tamanho inválido de CST no produto " + f.c_produto[i].value + "!!");
        //            f.c_cst[i].focus();
        //            return;
        //        }
        //        if (trim(f.c_ncm[i].value) != trim(f.c_ncm_redigite[i].value)) {
        //            alert("Falha na conferência do NCM redigitado do produto " + f.c_produto[i].value + "!!");
        //            f.c_ncm_redigite[i].focus();
        //            return;
        //        }
        //        if (trim(f.c_cst[i].value) != trim(f.c_cst_redigite[i].value)) {
        //            alert("Falha na conferência do CST redigitado do produto " + f.c_produto[i].value + "!!");
        //            f.c_cst_redigite[i].focus();
        //            return;
        //        }
        //    }
        //}

        $("#dCONFIRMA").hide();
        window.status = "Aguarde ...";
        f.submit();

        ////  VERIFICA SE OS CÓDIGOS DE NCM ESTÃO CADASTRADOS NA TABELA DO IBPT
        ////  A FUNÇÃO CHAMADA NO CALLBACK IRÁ EXIBIR UMA MENSAGEM NO CASO DE ENCONTRAR
        ////  CÓDIGOS NÃO CADASTRADOS OU FAZER O SUBMIT() SE ESTIVER TUDO OK.
        //s_ibpt_ncm = "";
        //s_ncm_ja_listado = "";
        //for (i = 0; i < f.c_ncm.length; i++) {
        //    s_ncm_aux = f.c_ncm[i].value;
        //    if (s_ncm_aux != "") {
        //        if (s_ncm_ja_listado.indexOf("|" + s_ncm_aux + "|") == -1) {
        //            s_ncm_ja_listado += "|" + s_ncm_aux + "|";
        //            if (s_ibpt_ncm.length > 0) s_ibpt_ncm += ",";
        //            s_ibpt_ncm += s_ncm_aux;
        //        }
        //    }
        //}

        //$("#divAjaxProgress").show();
        //$.getJSON(
        //    "../Global/IbptNcmConsultaBD.asp",
        //    { ncm: s_ibpt_ncm },
        //    trataRetornoConsultaIbpt)
        //.fail(function(jqXHR, textStatus, errorThrown) {
        //    $("#divAjaxProgress").hide();
        //    $("#dCONFIRMA").show();
        //    $("#divDialogBox div").html("<b>Ocorreu um erro ao tentar consultar os dados no servidor!!<br />Por favor, tente novamente.</b>" + "<br /><br /><b><i>Status:</i></b><br />" + textStatus + "<br /><br /><b><i>Erro ocorrido:</i></b><br />" + errorThrown + "<br /><br /><b><i>Descrição do erro:</i></b><br />" + jqXHR.responseText);
        //    $("#divDialogBox").dialog("option", "title", "Erro!");
        //    $("#divDialogBox").dialog("open");
        //})
        //.always(function() { $("#divAjaxProgress").hide() });
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
<link href="<%=URL_FILE__JQUERY_UI_CSS%>" rel="stylesheet" type="text/css">

<style type="text/css">
    #ckb_especial_aux {
        margin: 0pt 2pt 1pt 15pt;
        vertical-align: top;
    }

    .rbOpt {
        vertical-align: bottom;
    }

    .lblOpt {
        vertical-align: bottom;
    }

    .divMsgAlerta {
        margin-top: 30px;
        margin-bottom: 10px;
        border: solid 2px #000000;
        font-weight: bold;
        text-align: center;
        padding: 10px;
        width: 760px;
        color: #EF0000;
        background-color: #FFFFC4;
    }

    .divAjaxProgress {
        margin-top: 15px;
        text-align: center;
        vertical-align: middle;
    }

    .TxtEditavel{
    	color: blue;
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
<body onload="$('#c_obs').focus();">
    <center>

<form id="fESTOQ" name="fESTOQ" method="post" action="EstoqueTransfereEntreCDsConfirma.asp" autocomplete="off">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="c_log_edicao" id="c_log_edicao" value="">
<input type="hidden" name="c_nfe_emitente_origem" id="c_nfe_emitente_origem" value="<%=c_nfe_emitente_origem%>">
<input type="hidden" name="c_nfe_emitente_destino" id="c_nfe_emitente_destino" value="<%=c_nfe_emitente_destino%>">
<!-- É NECESSÁRIO CRIAR UM CAMPO DO TIPO HIDDEN PARA QUE A PÁGINA SEGUINTE CONSIGA
	 RECUPERAR A INFORMAÇÃO REFERENTE A ESTE CAMPO, JÁ QUE REQUEST.FORM() EM UM
	 CAMPO DO TIPO CHECKBOX QUE ESTÁ DISABLED RETORNA VAZIO.
-->
<input type="hidden" name="ckb_especial" id="ckb_especial" value="<%=ckb_especial%>">

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="780" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Transferência de Produtos Entre CD's</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>
<br>

<!--  CADASTRO DA TRANSFERÊNCIA DE ESTOQUES  -->
<table class="Qx" cellspacing="0" cellpadding="0">
<!--  EMPRESA  ORIGEM -->
	<tr class="trWmsCd">
		<td>
			<table width="100%" cellpadding="0" cellspacing="0">
				<tr>
					<td class="MT" align="left">
						<span class="PLTe">Empresa Origem</span>
						<br />
						<span class="PLLe" style="margin-left:2pt;"><%=obtem_apelido_empresa_NFe_emitente(c_nfe_emitente_origem)%></span>
					</td>
				</tr>
			</table>
		</td>
	</tr>
<!--  EMPRESA DESTINO -->
	<tr class="trWmsCd">
		<td>
			<table width="100%" cellpadding="0" cellspacing="0">
				<tr>
					<td class="MT" align="left">
						<span class="PLTe">Empresa Destino</span>
						<br />
						<span class="PLLe" style="margin-left:2pt;"><%=obtem_apelido_empresa_NFe_emitente(c_nfe_emitente_destino)%></span>
					</td>
				</tr>
			</table>
		</td>
	</tr>
<!--  DOCUMENTO  -->
	<tr bgcolor="#FFFFFF">
    <td class="MDBE" align="left" nowrap><span class="PLTe">Documento da Transferência</span>
		<br><input name="c_documento_transf" id="c_documento_transf" class="PLLe" style="width:270px;margin-left:2pt;color:blue;" maxlength="30" value="Transferencia">
    </td>
	</tr>
<!--  OBSERVAÇÃO  -->
	<tr bgcolor="#FFFFFF">
	    <td class="MDBE" style="padding-top:3px;">
		    <table cellspacing="0" cellpadding="0" border="0" width="100%">
			    <tr>
				    <td width="50%">
					    <span class="PLTe">Observações</span>
				    </td>
				    <td width="50%" align="right">
					    <span class="PLLd">Tamanho restante:</span><input name="c_tamanho_restante" id="c_tamanho_restante" tabindex=-1 readonly class="TA" style="width:35px;text-align:right;" value="<%=Cstr(500)%>" />
				    </td>
			    </tr>
			    <tr>
				    <td colspan="2">
					    <textarea name="c_obs" id="c_obs" class="PLLe" rows="<%=Cstr(MAX_LINHAS_ESTOQUE_OBS)%>"
							    style="width:642px;margin-left:2pt;" onkeypress="limita_tamanho(this,500);" onblur="this.value=trim(this.value);calcula_tamanho_restante();"
							    onkeyup="calcula_tamanho_restante();"
					    ></textarea>
				    </td>
			    </tr>
		    </table>
	    </td>
	</tr>
</table>
<br>

<!--  R E L A Ç Ã O   D E   P R O D U T O S  -->
<table id="tableProduto" class="Qx" cellspacing="0" border="0">
	<thead>
	<tr bgcolor="#FFFFFF">
	<th>&nbsp;</th>
    <th class="MB" align="left" valign="bottom"><span class="PLTe">Fabricante</span></th>
	<th class="MB" align="left" valign="bottom"><span class="PLTe">Produto</span></th>
	<th class="MB" align="left" valign="bottom"><span class="PLTe">Descrição</span></th>
	<th class="MB" align="left" valign="bottom"><span class="PLTe">Documento</span></th>
    <th class="MB" align="center" valign="bottom" align="center"><span class="PLTe">Tipo de<br /><span style="font-size:7pt;">Entrada</span></span></th>
	<th class="MB" align="center" valign="bottom" align="center"><span class="PLTe">EAN</span></th>
	<th class="MB" align="center" valign="bottom" align="center"><span class="PLTe">NCM</span></th>
	<th class="MB" align="center" valign="bottom" align="center"><span class="PLTe">CST</span></th>
	<th class="MB" align="right" valign="bottom"><span class="PLTd">Qtde</span></th>
	<th class="MB" align="center" valign="bottom" align="center"><span class="PLTe">Valor<br />Unitário</span></th>
	<th class="MB" align="center" valign="bottom" align="center"><span class="PLTe">Valor<br />Referência</span></th>
    <th class="MB" align="center" valign="bottom" align="center"><span class="PLTe">Aliq<br /><span style="font-size:7pt;">IPI (%)</span></span></th>
    <th class="MB" align="center" valign="bottom" align="center"><span class="PLTe">Valor<br /><span style="font-size:7pt;">IPI</span></span></th>
    <th class="MB" align="center" valign="bottom" align="center"><span class="PLTe">Aliq<br /><span style="font-size:7pt;">ICMS (%)</span></span></th>
    <th class="MB" align="center" valign="bottom" align="center"><span class="PLTe">NF Entrada<br /><span style="font-size:7pt;">N°</span></span></th>
	</tr>
	</thead>

	<tbody>
<%	for n=LBound(v_item_transf) to UBound(v_item_transf)
		with v_item_transf(n)
            s_fabricante = .fabricante
			s_produto = .produto
            s_id_estoque_origem = .id_estoque_origem
			s_descricao_html = produto_formata_descricao_em_html(.descricao_html)
            s_documento = .documento
            if .entrada_tipo = 1 then
                s_entrada_tipo = "XML"
            else
                s_entrada_tipo = "Manual"
                end if
			s_qtde = .qtde
            s_preco_fabricante = formata_moeda(.preco_fabricante)	
            s_vl_custo2 = formata_moeda(.vl_custo2)
			s_vl_BC_ICMS_ST = formata_moeda(.vl_BC_ICMS_ST)
            s_vl_ICMS_ST = formata_moeda(.vl_ICMS_ST)
            s_ncm = .ncm
            s_cst = .cst
            s_st_ncm_cst_herdado_tabela_produto = CStr(.st_ncm_cst_herdado_tabela_produto)
			s_ean = .ean
            s_aliq_ipi = formata_numero(.aliq_ipi, 0)
            s_aliq_icms = formata_numero(.aliq_icms, 0)
            s_vl_ipi = formata_moeda(.vl_ipi)
            s_preco_origem = formata_moeda(.preco_origem)
            s_produto_xml = .produto_xml
            s_nfe_entrada_numero = .nfe_entrada_numero
            end with
%>
	<tr>
	<td align="left"><input name="c_linha" readonly tabindex=-1 class="PLLe" maxlength="2" style="width:24px;text-align:right;color:#808080;" value="<%=Cstr(n+1) & ". " %>"></td>
	<td class="MDBE" align="left">
		<input name="c_fabricante" readonly tabindex=-1 class="PLLe" style="width:50px;"
			value="<%=s_fabricante%>"></td>
	<td class="MDBE" align="left">
		<input name="c_produto" readonly tabindex=-1 class="PLLe" style="width:50px;"
			value="<%=s_produto%>">
        <!-- incluir num campo oculto o id do estoque origem -->
        <input type="hidden" name="c_id_estoque_origem" value="<%=s_id_estoque_origem%>">
	</td>
	<td class="MDB" align="left">
		<span class="PLLe" style="width:240px;"><%=s_descricao_html%></span></td>
	<td class="MDBE" align="left">
		<input name="c_documento" readonly tabindex=-1 class="PLLe" style="width:50px;"
			value="<%=s_documento%>"></td>
	<td class="MDBE" align="left">
		<input name="c_entrada_tipo" readonly tabindex=-1 class="PLLe" style="width:50px;"
			value="<%=s_entrada_tipo%>"></td>
	<td class="MDB" align="left">
		<input name="c_ean" readonly tabindex=-1 class="PLLe" style="width:80px;"
			value="<%=s_ean%>"></td>
	<td class="MDB" align="left">
		<input name="c_ncm" readonly tabindex=-1 class="PLLc" maxlength="8" style="width:56px;"
            value="<%=s_ncm%>">
	</td>
	<td class="MDB" align="left">
		<input name="c_cst" readonly tabindex=-1 class="PLLc" maxlength="3" style="width:40px;"
            value="<%=s_cst%>">
	</td>
	<td class="MDB" align="right">
		<input name="c_qtde" readonly tabindex=-1 class="PLLd" style="width:30px;"
			value="<%=s_qtde%>"></td>
	<td class="MDB" align="right">
		<input name="c_vl_unitario" readonly tabindex=-1 class="PLLd" maxlength="12" style="width:62px;"
			value="<%=s_preco_fabricante%>">
		</td>
	<td class="MDB" align="right">
		<input name="c_vl_custo2" readonly tabindex=-1 class="PLLd" maxlength="12" style="width:62px;"
			value="<%=s_vl_custo2%>">
		</td>
	<td class="MDB" align="right">
		<input name="c_aliq_ipi" class="PLLd" maxlength="12" style="width:62px; color:blue;"
			onkeypress="if (digitou_enter(true)) $(this).hUtil('focusNext'); filtra_numerico();"
			onblur="if (trim(this.value)!='') this.value=formata_numero(this.value, 0); if (converte_numero(this.value)<0) {alert('Valor inválido!!');this.focus();});"
			value="<%=s_aliq_ipi%>">
		</td>
	<td class="MDB" align="right">
		<input name="c_vl_ipi" class="PLLd" maxlength="12" style="width:62px; color:blue;"
			onkeypress="if (digitou_enter(true)) $(this).hUtil('focusNext'); filtra_moeda();"
			onblur="if (trim(this.value)!='') this.value=formata_moeda(this.value); if (converte_numero(this.value)<0) {alert('Valor inválido!!');this.focus();});"
			value="<%=s_vl_ipi%>">
		</td>
	<td class="MDB" align="right">
		<input name="c_aliq_icms" class="PLLd" maxlength="12" style="width:62px; color:blue;"
			onkeypress="if (digitou_enter(true)) $(this).hUtil('focusNext'); filtra_numerico();"
			onblur="if (trim(this.value)!='') this.value=formata_numero(this.value, 0); if (converte_numero(this.value)<0) {alert('Valor inválido!!');this.focus();});"
			value="<%=s_aliq_icms%>">
		</td>
	<td class="MDB" align="left">
		<input name="c_nfe_entrada_numero" class="PLLd" maxlength="12" style="width:62px; color:blue;"
			onkeypress="if (digitou_enter(true)) $(this).hUtil('focusNext'); filtra_numerico();"
			onblur="if (trim(this.value)!='') this.value=formata_numero(this.value, 0); if (converte_numero(this.value)<0) {alert('Valor inválido!!');this.focus();});"
			value="<%=s_nfe_entrada_numero%>">
	</tr>
<% next %>
	</tbody>

</table>


<!--  ASSEGURA CRIAÇÃO DE UM ARRAY DE CAMPOS, MESMO QUANDO HOUVER SOMENTE 1 PRODUTO!! -->
<!--<input type=HIDDEN name="c_linha" value="">
<input type=HIDDEN name="c_fabricante" value="">
<input type=HIDDEN name="c_produto" value="">
<input type=HIDDEN name="c_id_estoque_origem" value="">
<input type=HIDDEN name="c_descricao_xml" value="">
<input type=HIDDEN name="c_documento" value="">
<input type=HIDDEN name="c_entrada_tipo" value="">
<input type=HIDDEN name="c_ean" value="">
<input type=HIDDEN name="c_ncm" value="">
<input type=HIDDEN name="c_cst" value="">
<input type=HIDDEN name="c_qtde" value="">
<input type=HIDDEN name="c_vl_custo2" value="">
<input type=HIDDEN name="c_aliq_ipi" value="">
<input type=HIDDEN name="c_aliq_icms" value="">
<input type=HIDDEN name="c_vl_ipi" value="">-->

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
'	FECHA CONEXAO COM O BANCO DE DADOS
	if rs.State <> 0 then rs.Close
	set rs = nothing

	cn.Close
	set cn = nothing
%>