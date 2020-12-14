<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/Global.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ==================================
'	  EstoqueEntradaViaXmlConsiste.asp
'     ==================================
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

	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	if Not operacao_permitida(OP_CEN_ENTRADA_MERCADORIAS_ESTOQUE, s_lista_operacoes_permitidas) then
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	dim uploaded_file_guid
    dim uploaded_file_guid2
	uploaded_file_guid = Trim(Request("uploaded_file_guid"))
    uploaded_file_guid2 = Trim(Request("uploaded_file_guid2"))
	if uploaded_file_guid = "" then
		alerta=texto_add_br(alerta)
		alerta=alerta & "Nenhum identificador de arquivo foi informado."
		end if

	dim s, i, iQtdeItens
	dim c_nfe_qtde_itens, c_nfe_numero_nf, c_nfe_emitente_cnpj, c_nfe_destinatario_cnpj, c_nfe_emitente_nome, c_nfe_emitente_nome_fantasia
    dim c_nfe_dt_hr_emissao, c_nfe_dt_hr_emissao2
    dim c_perc_agio, c_total_nf, c_nfe_vl_total_geral
    dim arquivo_nfe, arquivo_nfe2
    dim c_op_upload

    c_op_upload = Trim(Request("c_op_upload"))
    if c_op_upload = "M" then
	    'c_nfe_qtde_itens = Trim(Request("iQtdeItensPreenchidos"))
        c_nfe_qtde_itens = MAX_PRODUTOS_ENTRADA_ESTOQUE
    else
        c_nfe_qtde_itens = Trim(Request("c_nfe_qtde_itens"))
        end if
    'c_nfe_qtde_itens = Trim(Request("c_nfe_qtde_itens"))

	c_nfe_numero_nf = Trim(Request("c_nfe_numero_nf"))
	c_nfe_emitente_cnpj = retorna_so_digitos(Trim(Request("c_nfe_emitente_cnpj")))
	c_nfe_destinatario_cnpj = retorna_so_digitos(Trim(Request("c_nfe_destinatario_cnpj")))
	c_nfe_emitente_nome = Trim(Request("c_nfe_emitente_nome"))
	c_nfe_emitente_nome_fantasia = Trim(Request("c_nfe_emitente_nome_fantasia"))
    c_nfe_dt_hr_emissao = Trim(Request("c_dt_hr_emissao"))
    c_nfe_dt_hr_emissao2 = Trim(Request("c_dt_hr_emissao2"))
    
    c_perc_agio = Trim(Request.Form("c_perc_agio"))
    c_total_nf = Trim(Request.Form("c_total_nf"))
    c_nfe_vl_total_geral = Trim(Request.Form("c_nfe_vl_total_geral"))

	iQtdeItens = converte_numero(c_nfe_qtde_itens)

    arquivo_nfe = Trim(Request.Form("arquivo_nfe"))
    arquivo_nfe2 = Trim(Request.Form("arquivo_nfe2"))

    dim c_xml_ide__cNF_1 
    dim c_xml_ide__serie_1 
    dim c_xml_ide__nNF_1 
    dim c_xml_emit__CNPJ_1 
    dim c_xml_emit__xNome_1 
    dim c_xml_dest__CNPJ_1 
    dim c_xml_dest__xNome_1 
    dim c_xml_transp__xNome_1 
    dim c_xml_transp__CNPJ_1 
    dim c_xml_det_nItem_1 
    dim c_xml_ide__cNF_2 
    dim c_xml_ide__serie_2 
    dim c_xml_ide__nNF_2 
    dim c_xml_emit__CNPJ_2 
    dim c_xml_emit__xNome_2 
    dim c_xml_dest__CNPJ_2 
    dim c_xml_dest__xNome_2 
    dim c_xml_transp__CNPJ_2 
    dim c_xml_transp__xNome_2 
    dim c_xml_det_nItem_2 
 

    c_xml_ide__cNF_1  = Trim(Request.Form("c_xml_ide__cNF_1"))
    c_xml_ide__serie_1  = Trim(Request.Form("c_xml_ide__serie_1"))
    c_xml_ide__nNF_1  = Trim(Request.Form("c_xml_ide__nNF_1"))
    c_xml_emit__CNPJ_1  = Trim(Request.Form("c_xml_emit__CNPJ_1"))
    c_xml_emit__xNome_1  = Trim(Request.Form("c_xml_emit__xNome_1"))
    c_xml_dest__CNPJ_1  = Trim(Request.Form("c_xml_dest__CNPJ_1"))
    c_xml_dest__xNome_1  = Trim(Request.Form("c_xml_dest__xNome_1"))
    c_xml_det_nItem_1  = Trim(Request.Form("c_xml_det_nItem_1"))
    c_xml_transp__xNome_1  = Trim(Request.Form("c_xml_transp__xNome_1"))
    c_xml_transp__CNPJ_1  = Trim(Request.Form("c_xml_transp__CNPJ_1"))
    c_xml_ide__cNF_2  = Trim(Request.Form("c_xml_ide__cNF_2"))
    c_xml_ide__serie_2  = Trim(Request.Form("c_xml_ide__serie_2"))
    c_xml_ide__nNF_2  = Trim(Request.Form("c_xml_ide__nNF_2"))
    c_xml_emit__CNPJ_2  = Trim(Request.Form("c_xml_emit__CNPJ_2"))
    c_xml_emit__xNome_2  = Trim(Request.Form("c_xml_emit__xNome_2"))
    c_xml_dest__CNPJ_2  = Trim(Request.Form("c_xml_dest__CNPJ_2"))
    c_xml_dest__xNome_2  = Trim(Request.Form("c_xml_dest__xNome_2"))
    c_xml_transp__CNPJ_2  = Trim(Request.Form("c_xml_transp__CNPJ_2"))
    c_xml_det_nItem_2  = Trim(Request.Form("c_xml_det_nItem_2"))
    c_xml_transp__xNome_2  = Trim(Request.Form("c_xml_transp__xNome_2"))



'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs, msg_erro
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

	dim alerta
	alerta = ""

	dim s_id_nfe_emitente
	s_id_nfe_emitente = Trim(Request.Form("c_id_nfe_emitente"))

'	if alerta = "" then
'		if c_id_nfe_emitente <> "" then
'			s = "SELECT id FROM t_NFe_EMITENTE WHERE (apelido = '" & c_id_nfe_emitente & "')"
'			if rs.State <> 0 then rs.Close
'			rs.Open s, cn
'			if Not rs.Eof then
'				s_id_nfe_emitente = Trim("" & rs("id"))
'				rs.MoveNext
'				if Not rs.Eof then
'				'	HÁ MAIS DE UM REGISTRO COM O MESMO APELIDO
'					s_id_nfe_emitente = ""
'					end if
'				end if
'			end if
'		end if

	dim s_fabricante_codigo, s_fabricante_nome, s_documento, s_nfe_dt_hr_emissao
	s_fabricante_codigo = ""
	s_fabricante_nome = ""
    s_documento = ""

	s_fabricante_codigo = Trim(Request("c_fabricante"))
	s_fabricante_nome = Trim(Request("c_nfe_emitente_nome"))
    s_documento = Trim(Request("c_documento"))
	' if alerta = "" then
		' if c_nfe_emitente_cnpj <> "" then
			' s = "SELECT fabricante, nome, razao_social FROM t_FABRICANTE WHERE (cnpj = '" & c_nfe_emitente_cnpj & "')"
			' if rs.State <> 0 then rs.Close
			' rs.Open s, cn
			' if Not rs.Eof then
				' s_fabricante_codigo = Trim("" & rs("fabricante"))
				' s_fabricante_nome = Trim("" & rs("nome"))
				' if s_fabricante_nome = "" then s_fabricante_nome = Trim("" & rs("razao_social"))
				' rs.MoveNext
				' if Not rs.Eof then
				' '	HÁ MAIS DE UM REGISTRO COM O MESMO CNPJ
					' s_fabricante_codigo = ""
					' s_fabricante_nome = ""
					' end if
				' end if
			' end if
		
		' if (s_fabricante_codigo = "") And (c_nfe_emitente_nome_fantasia <> "") then
			' s = "SELECT fabricante, nome, razao_social FROM t_FABRICANTE WHERE (nome LIKE '" & BD_CURINGA_TODOS & c_nfe_emitente_nome_fantasia & BD_CURINGA_TODOS & "')"
			' if rs.State <> 0 then rs.Close
			' rs.Open s, cn
			' if Not rs.Eof then
				' s_fabricante_codigo = Trim("" & rs("fabricante"))
				' s_fabricante_nome = Trim("" & rs("nome"))
				' if s_fabricante_nome = "" then s_fabricante_nome = Trim("" & rs("razao_social"))
				' rs.MoveNext
				' if Not rs.Eof then
				' '	HÁ MAIS DE UM REGISTRO NO RESULTADO
					' s_fabricante_codigo = ""
					' s_fabricante_nome = ""
					' end if
				' end if
			' end if
		' end if
 
	dim vDtHr, vDt, vHr
	if alerta = "" then
		s_nfe_dt_hr_emissao = Trim(Request("c_dt_hr_emissao"))
        's_nfe_dt_hr_emissao = c_nfe_dt_hr_emissao
		if s_nfe_dt_hr_emissao <> "" then
		'if false then
			vDtHr = Split(s_nfe_dt_hr_emissao, "T")
			vDt = Split(vDtHr(LBound(vDtHr)), "-")
			s_nfe_dt_hr_emissao = vDt(LBound(vDt)+2) & "/" & vDt(LBound(vDt)+1) & "/" & vDt(LBound(vDt))
			end if
		end if

    dim s_visibilidade
	dim s_ckb_importa
	dim s_prod_ean
	dim s_prod_cod
    dim s_prod_cod_nota
	dim s_prod_ncm
	dim s_prod_cst
	dim s_prod_cfop
	dim s_prod_unid
	dim s_prod_qtde
	dim s_prod_vl_unitario_nota
	dim s_prod_vl_unitario
	dim s_prod_vl_total
	dim s_prod_vl_base_icms
	dim s_prod_vl_icms
	dim s_prod_vl_ipi
	dim s_prod_aliq_icms
	dim s_prod_aliq_ipi	
    dim s_prod_vl_frete
	dim s_cod_produto_import
	dim s_desc_produto_import
    dim s_linha_importa
	dim tabela_confirma
    dim campos_ocultos
	dim icont
    dim v_prod_cod
    dim s_c1_xml_prod_cProd
    dim s_c1_xml_prod_cEAN
    dim s_c1_xml_prod__NCM
    dim s_c1_xml_prod__CFOP
    dim s_c1_xml_prod__qCom
    dim s_c1_xml_prod__vUnCom
    dim s_c1_xml_prod__vProd
    dim s_c1_xml_prod__vFrete
    dim s_c1_xml_imposto__pICMS
    dim s_c1_xml_imposto__pIPI
    dim s_c1_xml_imposto__vIPI
    dim s_c2_xml_prod_cProd
    dim s_c2_xml_prod_cEAN
    dim s_c2_xml_prod__NCM
    dim s_c2_xml_prod__CFOP
    dim s_c2_xml_prod__qCom
    dim s_c2_xml_prod__vUnCom
    dim s_c2_xml_prod__vProd
    dim s_c2_xml_prod__vFrete
    dim s_c2_xml_imposto__pICMS
    dim s_c2_xml_imposto__pIPI
    dim s_c2_xml_imposto__vIPI

	if alerta = "" then
		tabela_confirma = ""
		icont = 0
        redim v_prod_cod(0)
        set v_prod_cod(ubound(v_prod_cod)) = New cl_TRES_COLUNAS
        v_prod_cod(ubound(v_prod_cod)).c1 = ""		
		For i = 1 to iQtdeItens
            s_ckb_importa = Trim(Request.Form("ckb_importa_" & Trim(i)))            
		    s_prod_ean = Trim(Request.Form("c_ean_" & Trim(i)))
            s_prod_cod = Trim(Request.Form("c_erp_codigo_" & Trim(i)))
            s_prod_cod_nota = Trim(Request.Form("c_nfe_codigo_" & Trim(i)))
            'se  o código do produto na nota não está preenchido, trata-se de uma linha vazia
            'neste caso, ignorar geração do HTML
            if (s_prod_cod <> "") then
		        s_prod_ncm = Trim(Request.Form("c_nfe_ncm_sh_" & Trim(i)))
		        s_prod_cst = Trim(Request.Form("c_erp_cst_" & Trim(i)))
		        s_prod_cfop = Trim(Request.Form("c_nfe_cfop_" & Trim(i)))
		        s_prod_unid = Trim(Request.Form("c_nfe_unid_" & Trim(i)))
		        s_prod_qtde = Trim(Request.Form("c_nfe_qtde_" & Trim(i)))
		        s_prod_vl_unitario_nota = Trim(Request.Form("c_nfe_vl_unitario_nota_" & Trim(i)))
		        s_prod_vl_unitario = Trim(Request.Form("c_nfe_vl_unitario_" & Trim(i)))
		        s_prod_vl_total = Trim(Request.Form("c_nfe_vl_total_" & Trim(i)))
		        s_prod_vl_base_icms = Trim(Request.Form("c_nfe_vl_base_icms_" & Trim(i)))
		        s_prod_vl_icms = Trim(Request.Form("c_nfe_vl_icms_" & Trim(i)))
		        s_prod_vl_ipi = Trim(Request.Form("c_nfe_vl_ipi_" & Trim(i)))
		        s_prod_aliq_icms = Trim(Request.Form("c_nfe_aliq_icms_" & Trim(i)))
		        s_prod_aliq_ipi = Trim(Request.Form("c_nfe_aliq_ipi_" & Trim(i)))
                s_prod_vl_frete = Trim(Request.Form("c_nfe_vl_frete_" & Trim(i)))
                s_c1_xml_prod_cProd = Trim(Request.Form("c1_xml_prod_cProd_" & Trim(i)))
                s_c1_xml_prod_cEAN = Trim(Request.Form("c1_xml_prod_cEAN_" & Trim(i)))
                s_c1_xml_prod__NCM = Trim(Request.Form("c1_xml_prod__NCM_" & Trim(i)))
                s_c1_xml_prod__CFOP = Trim(Request.Form("c1_xml_prod__CFOP_" & Trim(i)))
                s_c1_xml_prod__qCom = Trim(Request.Form("c1_xml_prod__qCom_" & Trim(i)))
                s_c1_xml_prod__vUnCom = Trim(Request.Form("c1_xml_prod__vUnCom_" & Trim(i)))
                s_c1_xml_prod__vProd = Trim(Request.Form("c1_xml_prod__vProd_" & Trim(i)))
                s_c1_xml_imposto__pICMS = Trim(Request.Form("c1_xml_imposto__pICMS_" & Trim(i)))
                s_c1_xml_imposto__pIPI = Trim(Request.Form("c1_xml_imposto__pIPI_" & Trim(i)))
                s_c1_xml_imposto__vIPI = Trim(Request.Form("c1_xml_imposto__vIPI_" & Trim(i)))                
                s_c1_xml_prod__vFrete = Trim(Request.Form("c1_xml_prod__vFrete_" & Trim(i)))
                s_c2_xml_prod_cProd = Trim(Request.Form("c2_xml_prod_cProd_" & Trim(i)))
                s_c2_xml_prod_cEAN = Trim(Request.Form("c2_xml_prod_cEAN_" & Trim(i)))
                s_c2_xml_prod__NCM = Trim(Request.Form("c2_xml_prod__NCM_" & Trim(i)))
                s_c2_xml_prod__CFOP = Trim(Request.Form("c2_xml_prod__CFOP_" & Trim(i)))
                s_c2_xml_prod__qCom = Trim(Request.Form("c2_xml_prod__qCom_" & Trim(i)))
                s_c2_xml_prod__vUnCom = Trim(Request.Form("c2_xml_prod__vUnCom_" & Trim(i)))
                s_c2_xml_prod__vProd = Trim(Request.Form("c2_xml_prod__vProd_" & Trim(i)))
                s_c2_xml_imposto__pICMS = Trim(Request.Form("c2_xml_imposto__pICMS_" & Trim(i)))
                s_c2_xml_imposto__pIPI = Trim(Request.Form("c2_xml_imposto__pIPI_" & Trim(i)))
                s_c2_xml_imposto__vIPI = Trim(Request.Form("c2_xml_imposto__vIPI_" & Trim(i)))
                s_c2_xml_prod__vFrete = Trim(Request.Form("c2_xml_prod__vFrete_" & Trim(i)))
                


			    s_visibilidade = "style='display:none'" 'campos não visíveis entrarão como não importados
			    s_cod_produto_import = "COD_NOP"
			    s_desc_produto_import = "DESC_NOP"
                s_linha_importa = "IMPORTA_N"
			    if (s_ckb_importa = "IMPORTA_ON") then
				    if (s_prod_cod <> "") then
					    s_visibilidade = "" 'campos visíveis entrarão como importados
					    s_cod_produto_import = ""
					    s_desc_produto_import = ""
					    s = "SELECT * FROM t_PRODUTO WHERE (fabricante = '" & s_fabricante_codigo & "') AND (produto = '" & s_prod_cod & "')"
					    if rs.State <> 0 then rs.Close
					    rs.Open s, cn
					    if Not rs.Eof then
						    s_cod_produto_import = Trim("" & rs("produto"))
						    s_desc_produto_import = Trim("" & rs("descricao"))
                            s_linha_importa = "IMPORTA_S"
						    End If
					    end if
				    end if
			    icont = icont + 1
			    if s_cod_produto_import <> "" then
				    tabela_confirma = tabela_confirma & "<tr id='TR_" & Cstr(icont) & "' " & s_visibilidade & ">" & vbCRLF
				    tabela_confirma = tabela_confirma & "<td align='left'>" & vbCRLF
				    tabela_confirma = tabela_confirma & "	<input name='c_linha_" & Cstr(icont) & "' id='c_linha_" & Cstr(icont) & "' readonly tabindex=-1 class='PLLe' maxlength='2' style='width:30px;text-align:right;color:#808080;' " & vbCRLF
				    tabela_confirma = tabela_confirma & "		value='" & Cstr(icont) & "' />" & vbCRLF
				    tabela_confirma = tabela_confirma & "</td>" & vbCRLF
                    tabela_confirma = tabela_confirma & "<td class='MDBE TdErpCodigo' align='left'>" & vbCRLF
				    tabela_confirma = tabela_confirma & "	<input type='hidden' name='c_nfe_nItem_" & Cstr(icont) & "' id='c_nfe_nItem_" & Cstr(icont) & "' value='" & s_linha_importa & "' />" & vbCRLF
				    tabela_confirma = tabela_confirma & "	<input name='c_erp_codigo_" & Cstr(icont) & "' id='c_erp_codigo_" & Cstr(icont) & "' class='PLLe' style='width:70px' value='" & s_cod_produto_import & "' />" & vbCRLF
				    tabela_confirma = tabela_confirma & "</td>" & vbCRLF

   				    tabela_confirma = tabela_confirma & "<td class='MDB TdNfeCodigo' align='left'>" & vbCRLF
                    tabela_confirma = tabela_confirma & "   <input type='hidden' name='c_nfe_codigo_" & Cstr(icont) & "' id='c_nfe_codigo_" & Cstr(icont) & "' class='PLLe' value='" & s_prod_cod_nota & "' />" & vbCRLF
				    tabela_confirma = tabela_confirma & "	<input name='c_ean_" & Cstr(icont) & "' id='c_ean_" & Cstr(icont) & "' class='PLLe' style='width:100px' value='" & s_prod_ean & "' />" & vbCRLF
				    tabela_confirma = tabela_confirma & "</td>" & vbCRLF

				    tabela_confirma = tabela_confirma & "<td class='MDB TdNfeDescricao' align='left'>" & vbCRLF
				    tabela_confirma = tabela_confirma & "	<input name='c_nfe_descricao_" & Cstr(icont) & "' id='c_nfe_descricao_" & Cstr(icont) & "' class='PLLe TdNfeDescricao' style='width:240px' value='" & s_desc_produto_import & "' />" & vbCRLF
				    tabela_confirma = tabela_confirma & "</td>" & vbCRLF
				    tabela_confirma = tabela_confirma & "<td class='MDB TdNfeNcm' align='left'>" & vbCRLF
				    tabela_confirma = tabela_confirma & "	<input name='c_nfe_ncm_sh_" & Cstr(icont) & "' id='c_nfe_ncm_sh_'" & Cstr(icont) & "' class='PLLe TdNfeNcm' value='" & s_prod_ncm & "' />" & vbCRLF
				    tabela_confirma = tabela_confirma & "</td>" & vbCRLF
				    tabela_confirma = tabela_confirma & "<td class='MDB TdErpCst' align='left'>" & vbCRLF
				    tabela_confirma = tabela_confirma & "	<input name='c_nfe_cst_" & Cstr(icont) & "' id='c_nfe_cst_" & Cstr(icont) & "' class='PLLe TdErpCst' value='" & s_prod_cst & "' />" & vbCRLF
				    tabela_confirma = tabela_confirma & "</td>" & vbCRLF
				    tabela_confirma = tabela_confirma & "<td class='MDB TdNfeCfop' align='left'>" & vbCRLF
				    tabela_confirma = tabela_confirma & "	<input name='c_nfe_cfop_" & Cstr(icont) & "' id='c_nfe_cfop_" & Cstr(icont) & "' class='PLLe TdNfeCfop' value='" & s_prod_cfop & "' />" & vbCRLF
				    tabela_confirma = tabela_confirma & "</td>" & vbCRLF
				    tabela_confirma = tabela_confirma & "<td class='MDB TdNfeQtde' align='left'>" & vbCRLF
				    tabela_confirma = tabela_confirma & "	<input name='c_nfe_qtde_" & Cstr(icont) & "' id='c_nfe_qtde_" & Cstr(icont) & "' class='PLLe TdNfeQtde' value='" & s_prod_qtde & "' />" & vbCRLF
				    tabela_confirma = tabela_confirma & "</td>" & vbCRLF
				    tabela_confirma = tabela_confirma & "<td class='MDB TdNfeVlUnit' align='left'>" & vbCRLF
				    tabela_confirma = tabela_confirma & "	<input name='c_nfe_vl_unitario_nota_" & Cstr(icont) & "' id='c_nfe_vl_unitario_nota_" & Cstr(icont) & "' class='PLLe TdNfeVlUnit' value='" & s_prod_vl_unitario_nota & "' />" & vbCRLF
				    tabela_confirma = tabela_confirma & "<td class='MDB TdNfeVlUnit' align='left'>" & vbCRLF
				    tabela_confirma = tabela_confirma & "	<input name='c_nfe_vl_unitario_" & Cstr(icont) & "' id='c_nfe_vl_unitario_" & Cstr(icont) & "' class='PLLe TdNfeVlUnit' value='" & s_prod_vl_unitario & "' />" & vbCRLF
				    tabela_confirma = tabela_confirma & "</td>" & vbCRLF
				    tabela_confirma = tabela_confirma & "<td class='MDB TdNfeAliqIpi' align='left'>" & vbCRLF
				    tabela_confirma = tabela_confirma & "	<input name='c_nfe_aliq_ipi_" & Cstr(icont) & "' id='c_nfe_aliq_ipi_" & Cstr(icont) & "' class='PLLe TdNfeAliqIpi' value='" & s_prod_aliq_ipi & "' />" & vbCRLF
				    tabela_confirma = tabela_confirma & "</td>" & vbCRLF
				    tabela_confirma = tabela_confirma & "<td class='MDB TdNfeVlIpi' align='left'>" & vbCRLF
				    tabela_confirma = tabela_confirma & "	<input name='c_nfe_vl_ipi_" & Cstr(icont) & "' id='c_nfe_vl_ipi_" & Cstr(icont) & "' class='PLLe TdNfeVlIpi' value='" & s_prod_vl_ipi & "' />" & vbCRLF
				    tabela_confirma = tabela_confirma & "</td>" & vbCRLF
				    tabela_confirma = tabela_confirma & "<td class='MDB TdNfeAliqIcms' align='left'>" & vbCRLF
				    tabela_confirma = tabela_confirma & "	<input name='c_nfe_aliq_icms_" & Cstr(icont) & "' id='c_nfe_aliq_icms_" & Cstr(icont) & "' class='PLLe TdNfeAliqIcms' value='" & s_prod_aliq_icms & "' />" & vbCRLF
				    tabela_confirma = tabela_confirma & "</td>" & vbCRLF
				    tabela_confirma = tabela_confirma & "<td class='MDB TdNfeVlIpi' align='left'>" & vbCRLF
				    tabela_confirma = tabela_confirma & "	<input name='c_nfe_vl_frete_" & Cstr(icont) & "' id='c_nfe_vl_frete_" & Cstr(icont) & "' class='PLLe TdNfeVlIpi' value='" & s_prod_vl_frete & "' />" & vbCRLF
				    tabela_confirma = tabela_confirma & "</td>" & vbCRLF
				    tabela_confirma = tabela_confirma & "<td class='MDB TdNfeVlTotal' align='left'>" & vbCRLF
				    tabela_confirma = tabela_confirma & "	<input name='c_nfe_vl_total_" & Cstr(icont) & "' id='c_nfe_vl_total_" & Cstr(icont) & "' class='PLLe TdNfeVlTotal' value='" & s_prod_vl_total & "' />" & vbCRLF
				    tabela_confirma = tabela_confirma & "</td>" & vbCRLF
				    tabela_confirma = tabela_confirma & "</tr>" & vbCRLF

                    'incluir o código do produto no vetor para testar duplicidade
                    if Trim(v_prod_cod(ubound(v_prod_cod)).c1) <> "" then
				        redim preserve v_prod_cod(ubound(v_prod_cod)+1)
				        set v_prod_cod(ubound(v_prod_cod)) = New cl_TRES_COLUNAS
				        end if                			
                    v_prod_cod(ubound(v_prod_cod)).c1 = s_fabricante_codigo
                    v_prod_cod(ubound(v_prod_cod)).c2 = s_prod_cod
                    v_prod_cod(ubound(v_prod_cod)).c3 = s_linha_importa
			    else
				    if alerta <> "" then alerta = alerta & "<br>"
				    alerta = alerta & "O produto de código " & s_prod_cod & " do fabricante " & s_fabricante_codigo & " não foi encontrado."
				    end if						

                campos_ocultos = campos_ocultos & "   <input type='hidden' name='c1_xml_prod_cProd_" & Cstr(icont) & "' id='c1_xml_prod_cProd_" & Cstr(icont) & "' class='PLLe' value='" & s_c1_xml_prod_cProd & "' />" & vbCRLF
                campos_ocultos = campos_ocultos & "   <input type='hidden' name='c1_xml_prod_cEAN_" & Cstr(icont) & "' id='c1_xml_prod_cEAN_" & Cstr(icont) & "' class='PLLe' value='" & s_c1_xml_prod_cEAN & "' />" & vbCRLF
                campos_ocultos = campos_ocultos & "   <input type='hidden' name='c1_xml_prod__NCM_" & Cstr(icont) & "' id='c1_xml_prod__NCM_" & Cstr(icont) & "' class='PLLe' value='" & s_c1_xml_prod__NCM & "' />" & vbCRLF
                campos_ocultos = campos_ocultos & "   <input type='hidden' name='c1_xml_prod__CFOP_" & Cstr(icont) & "' id='c1_xml_prod__CFOP_" & Cstr(icont) & "' class='PLLe' value='" & s_c1_xml_prod__CFOP & "' />" & vbCRLF
                campos_ocultos = campos_ocultos & "   <input type='hidden' name='c1_xml_prod__qCom_" & Cstr(icont) & "' id='c1_xml_prod__qCom_" & Cstr(icont) & "' class='PLLe' value='" & s_c1_xml_prod__qCom & "' />" & vbCRLF
                campos_ocultos = campos_ocultos & "   <input type='hidden' name='c1_xml_prod__vUnCom_" & Cstr(icont) & "' id='c1_xml_prod__vUnCom_" & Cstr(icont) & "' class='PLLe' value='" & s_c1_xml_prod__vUnCom & "' />" & vbCRLF
                campos_ocultos = campos_ocultos & "   <input type='hidden' name='c1_xml_prod__vProd_" & Cstr(icont) & "' id='c1_xml_prod__vProd_" & Cstr(icont) & "' class='PLLe' value='" & s_c1_xml_prod__vProd & "' />" & vbCRLF
                campos_ocultos = campos_ocultos & "   <input type='hidden' name='c1_xml_prod__vFrete_" & Cstr(icont) & "' id='c1_xml_prod__vFrete_" & Cstr(icont) & "' class='PLLe' value='" & s_c1_xml_prod__vFrete & "' />" & vbCRLF
                campos_ocultos = campos_ocultos & "   <input type='hidden' name='c1_xml_imposto__pICMS_" & Cstr(icont) & "' id='c1_xml_imposto__pICMS_" & Cstr(icont) & "' class='PLLe' value='" & s_c1_xml_imposto__pICMS & "' />" & vbCRLF
                campos_ocultos = campos_ocultos & "   <input type='hidden' name='c1_xml_imposto__pIPI_" & Cstr(icont) & "' id='c1_xml_imposto__pIPI_" & Cstr(icont) & "' class='PLLe' value='" & s_c1_xml_imposto__pIPI & "' />" & vbCRLF
                campos_ocultos = campos_ocultos & "   <input type='hidden' name='c1_xml_imposto__vIPI_" & Cstr(icont) & "' id='c1_xml_imposto__vIPI_" & Cstr(icont) & "' class='PLLe' value='" & s_c1_xml_imposto__vIPI & "' />" & vbCRLF
                campos_ocultos = campos_ocultos & "   <input type='hidden' name='c2_xml_prod_cProd_" & Cstr(icont) & "' id='c2_xml_prod_cProd_" & Cstr(icont) & "' class='PLLe' value='" & s_c2_xml_prod_cProd & "' />" & vbCRLF
                campos_ocultos = campos_ocultos & "   <input type='hidden' name='c2_xml_prod_cEAN_" & Cstr(icont) & "' id='c2_xml_prod_cEAN_" & Cstr(icont) & "' class='PLLe' value='" & s_c2_xml_prod_cEAN & "' />" & vbCRLF
                campos_ocultos = campos_ocultos & "   <input type='hidden' name='c2_xml_prod__NCM_" & Cstr(icont) & "' id='c2_xml_prod__NCM_" & Cstr(icont) & "' class='PLLe' value='" & s_c2_xml_prod__NCM & "' />" & vbCRLF
                campos_ocultos = campos_ocultos & "   <input type='hidden' name='c2_xml_prod__CFOP_" & Cstr(icont) & "' id='c2_xml_prod__CFOP_" & Cstr(icont) & "' class='PLLe' value='" & s_c2_xml_prod__CFOP & "' />" & vbCRLF
                campos_ocultos = campos_ocultos & "   <input type='hidden' name='c2_xml_prod__qCom_" & Cstr(icont) & "' id='c2_xml_prod__qCom_" & Cstr(icont) & "' class='PLLe' value='" & s_c2_xml_prod__qCom & "' />" & vbCRLF
                campos_ocultos = campos_ocultos & "   <input type='hidden' name='c2_xml_prod__vUnCom_" & Cstr(icont) & "' id='c2_xml_prod__vUnCom_" & Cstr(icont) & "' class='PLLe' value='" & s_c2_xml_prod__vUnCom & "' />" & vbCRLF
                campos_ocultos = campos_ocultos & "   <input type='hidden' name='c2_xml_prod__vProd_" & Cstr(icont) & "' id='c2_xml_prod__vProd_" & Cstr(icont) & "' class='PLLe' value='" & s_c2_xml_prod__vProd & "' />" & vbCRLF
                campos_ocultos = campos_ocultos & "   <input type='hidden' name='c2_xml_prod__vFrete_" & Cstr(icont) & "' id='c2_xml_prod__vFrete_" & Cstr(icont) & "' class='PLLe' value='" & s_c2_xml_prod__vFrete & "' />" & vbCRLF
                campos_ocultos = campos_ocultos & "   <input type='hidden' name='c2_xml_imposto__pICMS_" & Cstr(icont) & "' id='c2_xml_imposto__pICMS_" & Cstr(icont) & "' class='PLLe' value='" & s_c2_xml_imposto__pICMS & "' />" & vbCRLF
                campos_ocultos = campos_ocultos & "   <input type='hidden' name='c2_xml_imposto__pIPI_" & Cstr(icont) & "' id='c2_xml_imposto__pIPI_" & Cstr(icont) & "' class='PLLe' value='" & s_c2_xml_imposto__pIPI & "' />" & vbCRLF
                campos_ocultos = campos_ocultos & "   <input type='hidden' name='c2_xml_imposto__vIPI_" & Cstr(icont) & "' id='c2_xml_imposto__vIPI_" & Cstr(icont) & "' class='PLLe' value='" & s_c2_xml_imposto__vIPI & "' />" & vbCRLF

                end if 'if (s_prod_cod_nota <> "")

			Next
		
        'verificando a existência de códigos repetidos
        for i=lbound(v_prod_cod) to ubound(v_prod_cod)-1
            for icont=i+1 to ubound(v_prod_cod)
                if (v_prod_cod(i).c2 = v_prod_cod(icont).c2) and _
                   (v_prod_cod(i).c3 = "IMPORTA_S") and _
                   (v_prod_cod(icont).c3 = "IMPORTA_S") then
                    alerta=texto_add_br(alerta)
					alerta=alerta & "Produto " & v_prod_cod(i).c2 & ": linha " & renumera_com_base1(Lbound(v_prod_cod),icont) & " repete o mesmo produto da linha " & renumera_com_base1(Lbound(v_prod_cod),i) & "."
                    end if
                next
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
<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>

<script type="text/javascript">
	var serverVariableUrl;
	var uploaded_file_guid;
	var nfe;
	var iQtdeItens = '<%=iQtdeItens%>';

	serverVariableUrl = '<%=Request.ServerVariables("URL")%>';
	serverVariableUrl = serverVariableUrl.toUpperCase();
	serverVariableUrl = serverVariableUrl.substring(0, serverVariableUrl.indexOf("CENTRAL"));
	uploaded_file_guid = '<%=uploaded_file_guid%>';

//Dynamically assign height
function sizeDivAjaxRunning() {
	var newTop = $(window).scrollTop() + "px";
	$("#divAjaxRunning").css("top", newTop);
}

//' $(function () {
//	' $("#divAjaxRunning").hide();

//	' $("#divAjaxRunning").css('filter', 'alpha(opacity=60)'); // TRANSPARÊNCIA NO IE8

//	' $(document).ajaxStart(function () {
//		' $("#divAjaxRunning").show();
//	' })
//	' .ajaxStop(function () {
//		' $("#divAjaxRunning").hide();
//	' });

//	' //Every resize of window
//	' $(window).resize(function() {
//		' sizeDivAjaxRunning();
//	' });

//	' //Every scroll of window
//	' $(window).scroll(function() {
//		' sizeDivAjaxRunning();
//	' });

//	' // Trata o problema em que os campos do formulário são limpos após retornar à esta página c/ o history.back() pela 2ª vez quando ocorre erro de consistência
//	' if (trim(fESTOQ.c_FormFieldValues.value) != "")
//	' {
//		' stringToForm(fESTOQ.c_FormFieldValues.value, $('#fESTOQ'));
//	' }

//	' var jqxhr = $.ajax({
//		' url: '<%=getProtocoloEmUsoHttpOrHttps%>://<%=Request.ServerVariables("SERVER_NAME")%>:<%=Request.ServerVariables("SERVER_PORT")%>' + serverVariableUrl + 'WebAPI/api/UploadedFile/ConvertXmlToJson',
//		' type: "GET",
//		' dataType: 'json',
//		' data: {
//			' id: uploaded_file_guid
//		' }
//	' })
//	' .done(function (response) {
//		' $("#divAjaxRunning").hide();
//		' nfe = response;
//		' preencheForm();
//	' })
//	' .fail(function (jqXHR, textStatus) {
//		' $("#divAjaxRunning").hide();
//		' var msgErro = "";
//		' if (textStatus.toString().length > 0) msgErro = "Mensagem de Status: " + textStatus.toString();
//		' try {
//			' if (jqXHR.status.toString().length > 0) {if (msgErro.length > 0) msgErro += "\n\n"; msgErro += "Status: " + jqXHR.status.toString();}
//		' } catch (e) { }

//		' try {
//			' if (jqXHR.statusText.toString().length > 0) {if (msgErro.length > 0) msgErro += "\n\n"; msgErro += "Descrição do Status: " + jqXHR.statusText.toString();}
//		' } catch (e) { }
		
//		' try {
//			' if (jqXHR.responseText.toString().length > 0) {if (msgErro.length > 0) msgErro += "\n\n"; msgErro += "Mensagem de Resposta: " + jqXHR.responseText.toString();}
//		' } catch (e) { }
		
//		' alert("Falha ao tentar processar a requisição!!\n\n" + msgErro);
//	' });
//' });

//function preencheForm()
//{
//    var f, i, sIdx, childProp;
//    var iQtdeItens = <%=iQtdeItens%>;

//	if (nfe == null) {
//		alert("Os dados da NFe não foram recuperados corretamente!!");
//		return;
//	}

//	f = fESTOQ;
//    //for (i = 0; i < f.c_erp_codigo.length; i++) {
//	for (i = 0; i < iQtdeItens; i++) {
//		sIdx = (i + 1).toString();
//		if (i < nfe.nfeProc.NFe.infNFe.det.length) {
//			$("#c_nfe_nItem_" + sIdx).val(nfe.nfeProc.NFe.infNFe.det[i]['@nItem']);
//			$("#c_nfe_codigo_" + sIdx).text(nfe.nfeProc.NFe.infNFe.det[i].prod.cProd);
//			$("#c_nfe_descricao_" + sIdx).text(nfe.nfeProc.NFe.infNFe.det[i].prod.xProd);
//			$("#c_nfe_ncm_sh_" + sIdx).text(nfe.nfeProc.NFe.infNFe.det[i].prod.NCM);
//			$("#c_nfe_cfop_" + sIdx).text(nfe.nfeProc.NFe.infNFe.det[i].prod.CFOP);
//			$("#c_nfe_unid_" + sIdx).text(nfe.nfeProc.NFe.infNFe.det[i].prod.uCom);
//			$("#c_nfe_qtde_" + sIdx).text(formata_numero(nfe.nfeProc.NFe.infNFe.det[i].prod.qCom, 1));
//			$("#c_nfe_vl_unitario_" + sIdx).text(formata_numero(nfe.nfeProc.NFe.infNFe.det[i].prod.vUnCom, 4));
//			$("#c_nfe_vl_total_" + sIdx).text(formata_numero(nfe.nfeProc.NFe.infNFe.det[i].prod.vProd, 2));
//			for (var key in nfe.nfeProc.NFe.infNFe.det[i].imposto.ICMS) {
//				if (nfe.nfeProc.NFe.infNFe.det[i].imposto.ICMS.hasOwnProperty(key)) {
//					childProp = nfe.nfeProc.NFe.infNFe.det[i].imposto.ICMS[key];
//					$("#c_nfe_cst_" + sIdx).text(childProp.orig.toString() + childProp.CST.toString());
//					$("#c_erp_cst_" + sIdx).val(converte_cst_nfe_fabricante_para_entrada_estoque($("#c_nfe_cst_" + sIdx).text()));
//					$("#c_nfe_vl_base_icms_" + sIdx).text(formata_numero(childProp.vBC, 2));
//					$("#c_nfe_vl_icms_" + sIdx).text(formata_numero(childProp.vICMS, 2));
//					$("#c_nfe_aliq_icms_" + sIdx).text(formata_numero(childProp.pICMS, 2));
//					break;
//				}
//			}
//			if (nfe.nfeProc.NFe.infNFe.det[i].imposto.IPI.IPITrib.hasOwnProperty('vIPI')) {
//				$("#c_nfe_vl_ipi_" + sIdx).text(formata_numero(nfe.nfeProc.NFe.infNFe.det[i].imposto.IPI.IPITrib.vIPI, 2));
//			}
//			if (nfe.nfeProc.NFe.infNFe.det[i].imposto.IPI.IPITrib.hasOwnProperty('pIPI')) {
//				$("#c_nfe_aliq_ipi_" + sIdx).text(formata_numero(nfe.nfeProc.NFe.infNFe.det[i].imposto.IPI.IPITrib.pIPI, 2));
//			}
//		}
//		else {
//			$("#c_erp_codigo_" + sIdx).prop("readonly", true);
//			$("#c_erp_cst_" + sIdx).prop("readonly", true);
//			$("#c_erp_codigo_" + sIdx).attr("tabindex", -1);
//			$("#c_erp_cst_" + sIdx).attr("tabindex", -1);
//		}
//	}

//	// Ajusta a altura dos campos input para ficar na mesma altura da linha da tabela
//	$(".TxtErpCodigo").each(function () {
//		$(this).height($(this).parent().height());
//	});

//	$(".TxtErpCst").each(function () {
//		$(this).height($(this).parent().height());
//	});

//	// Aceita somente dígitos
//	$(".TxtErpFabr, .TxtErpCodigo, .TxtErpCst").keydown(function (e) {
//		// Allow: delete, backspace, tab, escape, enter
//		if ($.inArray(e.keyCode, [46, 8, 9, 27, 13]) !== -1 ||
//			// Allow: Ctrl+A, Command+A, Ctrl+C, Ctrl+V, Ctrl+X
//			(((e.keyCode === 65) || (e.keyCode === 67) || (e.keyCode === 86) || (e.keyCode === 88)) && (e.ctrlKey === true || e.metaKey === true)) ||
//			// Allow: home, end, left, right, down, up
//			(e.keyCode >= 35 && e.keyCode <= 40)) {
//			// let it happen, don't do anything
//			return;
//		}
//		// Ensure that it is a number and stop the keypress
//		if ((e.shiftKey || (e.keyCode < 48 || e.keyCode > 57)) && (e.keyCode < 96 || e.keyCode > 105)) {
//			e.preventDefault();
//		}
//	});
//}

function realca_cor_linha(c, indice_row) {
	$("#TR_" + indice_row).css("background-color","palegreen");
	$("#TR_"+indice_row + " td input").css("background-color","palegreen");
	$("#TR_"+indice_row + " td span").css("background-color","palegreen");
	$(c).css("background-color","lightgray");
}

function normaliza_cor_linha(c, indice_row) {
	$("#TR_" + indice_row).css("background-color","");
	$("#TR_"+indice_row + " td input").css("background-color","");
	$("#TR_"+indice_row + " td span").css("background-color","");
	$(c).css("background-color","");
}

function fESTOQConfirma(f) {
	var s_id;

	s_id = "#c_id_nfe_emitente";
	if ($(s_id).val() == "") {
		alert("Selecione o CD!");
		$(s_id).focus();
		return;
	}

	s_id = "#c_fabricante";
	if ($(s_id).val() == "") {
		alert("Informe o código do fabricante!");
		$(s_id).focus();
		return;
	}

	s_id = "#c_documento";
	if ($(s_id).val() == "") {
		alert("Informe o número do documento!");
		$(s_id).focus();
		return;
	}

	for (var i = 1; i <= iQtdeItens; i++) {
		s_id = "#c_erp_codigo_" + i.toString();
		if ($(s_id).val() == "") {
			alert("Informe o código do produto no ERP!");
			$(s_id).focus();
			return;
		}

		s_id = "#c_erp_cst_" + i.toString();
		if ($(s_id).val() == "") {
			alert("Informe o CST para entrada no estoque!");
			$(s_id).focus();
			return;
		}
	}

	fESTOQ.c_FormFieldValues.value = formToStringAll($("#fESTOQ"));

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

<style type="text/css">
select
{
	margin-left:8px;
}
#divAjaxRunning
{
	position:absolute;
	top:0;
	left:0;
	width:100%;
	height:100%;
	z-index:1001;
	background-color:grey;
	opacity: .6;
}
.AjaxImgLoader
{
	position: absolute;
	left: 50%;
	top: 50%;
	margin-left: -128px; /* -1 * image width / 2 */
	margin-top: -128px;  /* -1 * image height / 2 */
	display: block;
}
.PLTe{
	margin-left:1pt;
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
.TxtErpCst{
	width: 30px;
	text-align:center;
}
.TdErpCodigo{
	width:50px;
	vertical-align: middle;
}
.TdNfeCodigo{
	width: 80px;
	vertical-align: middle;
}
.TdNfeDescricao{
	width: 160px;
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

<body>
<center>

<!-- AJAX EM ANDAMENTO -->
<div id="divAjaxRunning" style="display:none;"><img src="../Imagem/ajax_loader_gray_256.gif" class="AjaxImgLoader"/></div>

<!--form id="fESTOQ" name="fESTOQ" method="post" action="EstoqueEntradaViaXmlConsiste.asp" -->
<form id="fESTOQ" name="fESTOQ" method="post" action="EstoqueEntradaViaXmlConfirma.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="c_FormFieldValues" id="c_FormFieldValues" value="" / >
<input type="hidden" name="uploaded_file_guid" id="uploaded_file_guid" value="<%=uploaded_file_guid%>" / >
<input type="hidden" name="uploaded_file_guid2" id="uploaded_file_guid2" value="<%=uploaded_file_guid2%>" / >
<input type="hidden" name="c_nfe_qtde_itens" id="c_nfe_qtde_itens" />
<input type="hidden" name="iQtdeItens" id="iQtdeItens" value="<%=iQtdeItens%>" />
<input type="hidden" name="arquivo_nfe" id="arquivo_nfe" value="<%=arquivo_nfe%>"/>
<input type="hidden" name="arquivo_nfe2" id="arquivo_nfe2" value="<%=arquivo_nfe2%>"/>
<input type="hidden" name="c_dt_hr_emissao" id="c_dt_hr_emissao" value="<%=c_nfe_dt_hr_emissao%>"/>
<input type="hidden" name="c_dt_hr_emissao2" id="c_dt_hr_emissao2" value="<%=c_nfe_dt_hr_emissao2%>"/>

<input type="hidden" name="c_xml_ide__cNF_1" id="c_xml_ide__cNF_1"  value="<%=c_xml_ide__cNF_1%>" />
<input type="hidden" name="c_xml_ide__serie_1" id="c_xml_ide__serie_1" value="<%=c_xml_ide__serie_1%>" />
<input type="hidden" name="c_xml_ide__nNF_1" id="c_xml_ide__nNF_1" value="<%=c_xml_ide__nNF_1%>" />
<input type="hidden" name="c_xml_emit__CNPJ_1" id="c_xml_emit__CNPJ_1" value="<%=c_xml_emit__CNPJ_1%>" />
<input type="hidden" name="c_xml_emit__xNome_1" id="c_xml_emit__xNome_1" value="<%=c_xml_emit__xNome_1%>" />
<input type="hidden" name="c_xml_dest__CNPJ_1" id="c_xml_dest__CNPJ_1" value="<%=c_xml_dest__CNPJ_1%>" />
<input type="hidden" name="c_xml_dest__xNome_1" id="c_xml_dest__xNome_1" value="<%=c_xml_dest__xNome_1%>" />
<input type="hidden" name="c_xml_transp__CNPJ_1" id="c_xml_transp__CNPJ_1" value="<%=c_xml_transp__CNPJ_1%>" />
<input type="hidden" name="c_xml_det_nItem_1" id="c_xml_det_nItem_1" value="<%=c_xml_det_nItem_1%>" />
<input type="hidden" name="c_xml_transp__xNome_1" id="c_xml_transp__xNome_1" value="<%=c_xml_transp__xNome_1%>" />
<input type="hidden" name="c_xml_ide__cNF_2" id="c_xml_ide__cNF_2"  value="<%=c_xml_ide__cNF_2%>" />
<input type="hidden" name="c_xml_ide__serie_2" id="c_xml_ide__serie_2" value="<%=c_xml_ide__serie_2%>" />
<input type="hidden" name="c_xml_ide__nNF_2" id="c_xml_ide__nNF_2" value="<%=c_xml_ide__nNF_2%>" />
<input type="hidden" name="c_xml_emit__CNPJ_2" id="c_xml_emit__CNPJ_2" value="<%=c_xml_emit__CNPJ_2%>" />
<input type="hidden" name="c_xml_emit__xNome_2" id="c_xml_emit__xNome_2" value="<%=c_xml_emit__xNome_2%>" />
<input type="hidden" name="c_xml_dest__CNPJ_2" id="c_xml_dest__CNPJ_2" value="<%=c_xml_dest__CNPJ_2%>" />
<input type="hidden" name="c_xml_dest__xNome_2" id="c_xml_dest__xNome_2" value="<%=c_xml_dest__xNome_2%>" />
<input type="hidden" name="c_xml_transp__CNPJ_2" id="c_xml_transp__CNPJ_2" value="<%=c_xml_transp__CNPJ_2%>" />
<input type="hidden" name="c_xml_det_nItem_2" id="c_xml_det_nItem_2" value="<%=c_xml_det_nItem_2%>" />
<input type="hidden" name="c_xml_transp__xNome_2" id="c_xml_transp__xNome_2" value="<%=c_xml_transp__xNome_2%>" />


<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Entrada de Mercadorias no Estoque via XML</span>
	<br /><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>
<br />

<!--  CADASTRO DA ENTRADA DE MERCADORIAS NO ESTOQUE VIA XML  -->
<table class="Qx" cellspacing="0" cellpadding="0">
<!--  EMPRESA COMPRADORA / CENTRO DE DISTRIBUIÇÃO  -->
	<tr bgcolor="#FFFFFF" class="trWmsCd">
		<td colspan="3">
		<table width="100%" cellpadding="0" cellspacing="0">
		<tr>
			<td class="MT" align="left" width="50%"><span class="PLTe">Empresa</span>
			<br />
			<select id="c_id_nfe_emitente" name="c_id_nfe_emitente" style="margin-top:4pt;margin-bottom:4pt;min-width:100px;">
			<%=wms_apelido_empresa_nfe_emitente_monta_itens_select(s_id_nfe_emitente)%>
			</select>
			</td>
		</tr>
		</table>
		</td>
	</tr>

<!--  FABRICANTE/DOCUMENTO  -->
	<tr bgcolor="#FFFFFF">
		<td colspan="2" class="MDBE" align="left"><span class="PLTe">Fabricante</span>
			<br /><input name="c_nfe_emitente_nome" id="c_nfe_emitente_nome" class="PLLe TxtNfeEmitNome" readonly tabindex="-1" value="<%=filtra_nome_identificador(c_nfe_emitente_nome)%>" />
		</td>
    	<td class="MDB" style="border-left:0pt;" align="left"><span class="PLTe">% Ágio</span>
		<br><input name="c_perc_agio" id="c_perc_agio" class="PLLe TxtEditavel" maxlength="8" value="<%=c_perc_agio%>" onkeypress="if (digitou_enter(true)) $('#c_fabricante').focus();" onblur="this.value=formata_numero(this.value, 4); recalcula_itens();"></td>
	</tr>
	<tr bgcolor="#FFFFFF">
	<td class="MDBE" align="left" nowrap><span class="PLTe">Cód Fabricante (ERP)</span>
		<br><input name="c_fabricante" id="c_fabricante" class="PLLe TxtErpFabr TxtEditavel" maxlength="4" value="<%=s_fabricante_codigo%>" onkeypress="if ((digitou_enter(true))&&tem_info(this.value)) fESTOQ.c_documento.focus();" onblur="this.value=normaliza_codigo(this.value,TAM_MIN_FABRICANTE);"></td>
	<td class="MDB" style="border-left:0pt;" align="left"><span class="PLTe">Documento</span>
		<br><input name="c_documento" id="c_documento" class="PLLe TxtErpDocumento TxtEditavel" maxlength="30" value="<%=c_nfe_numero_nf%>" onkeypress="if (digitou_enter(true)) $('#c_erp_codigo_1').focus(); filtra_nome_identificador();" onblur="this.value=trim(this.value);"></td>
	<td class="MDB" style="border-left:0pt;" align="left"><span class="PLTe">Emissão</span>
		<br /><input name="c_nfe_dt_emissao" id="c_nfe_dt_emissao" class="PLLe TxtNfeDtHrEmissao" readonly tabindex="-1" value="<%=s_nfe_dt_hr_emissao%>" />
	</td>
	</tr>

<!--  ENTRADA ESPECIAL  -->
	<tr bgcolor="#FFFFFF">
	<td colspan="3" class="MDBE" align="left" nowrap><span class="PLTe">Tipo de Cadastramento</span>
		<br><input type="checkbox" class="rbOpt" tabindex="-1" id="ckb_especial" name="ckb_especial" value="ESPECIAL_ON"
		<%if Not operacao_permitida(OP_CEN_ENTRADA_ESPECIAL_ESTOQUE, s_lista_operacoes_permitidas) then Response.Write " disabled" %>
		><span class="C lblOpt" style="cursor:default" onclick="fESTOQ.ckb_especial.click();">Entrada Especial</span>
	</td>
	</tr>

	<tr bgColor="#FFFFFF">
	<td colspan="3" class="MDBE" align="left" nowrap><span class="PLTe">Observações</span>
		<br><textarea name="c_obs" id="c_obs" class="PLLe TxtErpObs TxtEditavel" rows="<%=Cstr(MAX_LINHAS_ESTOQUE_OBS)%>"
				onkeypress="limita_tamanho(this,MAX_TAM_T_ESTOQUE_CAMPO_OBS);" onblur="this.value=trim(this.value);"
				></textarea>
	</td>
	</tr>
</table>

<br />
<br />

<!--  R E L A Ç Ã O   D E   P R O D U T O S  -->
<table class="Qx" cellspacing="0">
	<tr bgColor="#FFFFFF">
	<td align="left">&nbsp;</td>
	<td class="MB TdErpCodigo" align="left" style="vertical-align:bottom;"><span class="PLTe">(ERP)<br />CÓD PROD</span></td>
	<td class="MB TdNfeCodigo" align="left" style="vertical-align:bottom;"><span class="PLTe">EAN</span></td>
	<td class="MB TdNfeDescricao" align="left" style="vertical-align:bottom;"><span class="PLTe">DESCRIÇÃO DO PROD/SERV</span></td>
	<td class="MB TdNfeNcm" align="left" style="vertical-align:bottom;"><span class="PLTe">NCM/SH</span></td>
	<td class="MB TdNfeCst" align="left" style="vertical-align:bottom;"><span class="PLTe">CST (ENTR)</span></td>
	<td class="MB TdNfeCfop" align="left" style="vertical-align:bottom;"><span class="PLTe">CFOP</span></td>
	<td class="MB TdNfeQtde" align="left" style="vertical-align:bottom;"><span class="PLTe">QUANT</span></td>
	<td class="MB TdNfeVlUnit" align="left" style="vertical-align:bottom;"><span class="PLTe">VL Nota</span></td>
	<td class="MB TdNfeVlUnit" align="left" style="vertical-align:bottom;"><span class="PLTe">VL Referência</span></td>
	<td class="MB TdNfeAliqIpi" align="left" style="vertical-align:bottom;"><span class="PLTe">A.IPI</span></td>
	<td class="MB TdNfeVlIpi" align="left" style="vertical-align:bottom;"><span class="PLTe">VL IPI</span></td>
	<td class="MB TdNfeAliqIcms" align="left" style="vertical-align:bottom;"><span class="PLTe">A.ICMS</span></td>
	<td class="MB TdNfeVlFrete" align="left" style="vertical-align:bottom;"><span class="PLTe">VL Frete</span></td>
	<td class="MB TdNfeVlTotal" align="left" style="vertical-align:bottom;"><span class="PLTe">VL TOTAL</span></td>
	</tr>
    <tbody>
<% if trim(tabela_confirma)<>"" then        		
        Response.Write tabela_confirma        
		end if
%>
    </tbody>
    <tfoot>
	    <tr>	
	    <td colspan="7" class="MD">&nbsp;</td>
	    <td class="MDB" align="left"><p class="Cd">Total NF</p></td>	
	    <td class="MDB" align="right"><input name="c_total_nf" id="c_total_nf" class="PLLd" style="width:62px;color:black;" 
	        value="<%=c_total_nf%>"></td>	
	    <td>&nbsp;</td>
        <td>&nbsp;</td>
	    <td>&nbsp;</td>
	    <td>&nbsp;</td>
	    <td class="MD">&nbsp;</td>
	    <td class="MDB" align="right"><input name="c_nfe_vl_total_geral" id="c_nfe_vl_total_geral" class="PLLd" style="width:70px;color:black;"
		    value="<%=c_nfe_vl_total_geral%>" readonly tabindex=-1 /></td>	
	    </tr>
	</tfoot>
</table>

<% if trim(tabela_confirma)<>"" then        		
        Response.Write campos_ocultos
		end if
%>

</form>

<br />

<!-- ************   SEPARADOR   ************ -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc" align="left">&nbsp;</td></tr>
</table>
<br />

<table width="649" cellspacing="0">
<tr>
	<td align="left"><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()" title="retorna para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
	<td align="right"><div name="dCONFIRMA" id="dCONFIRMA">
		<a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fESTOQConfirma(fESTOQ)" title="confirma a entrada das mercadorias no estoque">
		<img src="../botao/confirmar.gif" width="176" height="55" border="0"></a></div>
	</td>
</tr>
</table>

</center>
</body>
<% end if %>

</html>


<%
'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>