<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/estoque.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ================================================================
'	  E S T O Q U E E N T R A D A V I A X M L C O N F I R M A . A S P
'     ================================================================
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

'	class cl_ITEM_ESTOQUE_ENTRADA_XML
'		dim id_estoque
'		dim fabricante
'		dim produto
'		dim qtde
'		dim qtde_utilizada
'		dim preco_fabricante
'		dim data_ult_movimento
'		dim sequencia
'		dim vl_custo2
'		dim vl_BC_ICMS_ST
'		dim vl_ICMS_ST
'		dim ncm
'		dim ncm_redigite
'		dim cst
'		dim cst_redigite
'       dim ean
'        dim ean_original
'        dim aliq_ipi
'        dim aliq_icms
'        dim vl_ipi
'        dim preco_origem
'        dim produto_xml
'		end class
    
    class cl_ITEM_ESTOQUE_EAN
		dim fabricante
		dim produto
        dim ean
        dim ean_original
        end class
	
    dim s, s_log, i, n, usuario, msg_erro, c_log_edicao, s_nfe_dt_hr_emissao, s_nfe_hr_emissao, s_nfe_dt_hr_emissao2, s_nfe_hr_emissao2
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 
	
	dim alerta
	alerta=""

	dim r_estoque, v_item
    dim c_perc_agio
    dim arquivo_nfe, arquivo_nfe2
    dim s_sql    
    dim id_estoque_xml

'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	
	c_log_edicao = Trim(Request.Form("c_log_edicao"))
	c_perc_agio = Trim(Request.Form("c_perc_agio"))		
    arquivo_nfe = Trim(Request.Form("arquivo_nfe"))
    arquivo_nfe2 = Trim(Request.Form("arquivo_nfe2"))										

    dim uploaded_file_guid
    dim uploaded_file_guid2
	uploaded_file_guid = Trim(Request("uploaded_file_guid"))
    uploaded_file_guid2 = Trim(Request("uploaded_file_guid2"))

    dim vDtHr, vDt, vHr
	if alerta = "" then
		's_nfe_dt_hr_emissao = c_nfe_dt_hr_emissao
        s_nfe_dt_hr_emissao = Trim(Request("c_dt_hr_emissao"))	
		if s_nfe_dt_hr_emissao <> "" then
			vDtHr = Split(s_nfe_dt_hr_emissao, "T")
			vDt = Split(vDtHr(LBound(vDtHr)), "-")
            vHr = Split(vDtHr(UBound(vDtHr)), "-")
			s_nfe_dt_hr_emissao = vDt(LBound(vDt)+2) & "/" & vDt(LBound(vDt)+1) & "/" & vDt(LBound(vDt))
            s_nfe_hr_emissao = Mid(vHr(LBound(vHr)), 1, 2) & Mid(vHr(LBound(vHr)), 4, 2) & Mid(vHr(LBound(vHr)), 7, 2)
			end if
		end if
    
    dim vDtHr2, vDt2, vHr2
	if alerta = "" then
		s_nfe_dt_hr_emissao2 = Trim(Request("c_dt_hr_emissao2"))
		if s_nfe_dt_hr_emissao2 <> "" then
			vDtHr2 = Split(s_nfe_dt_hr_emissao2, "T")
			vDt2 = Split(vDtHr2(LBound(vDtHr2)), "-")
            vHr2 = Split(vDtHr2(UBound(vDtHr2)), "-")
			s_nfe_dt_hr_emissao2 = vDt2(LBound(vDt2)+2) & "/" & vDt2(LBound(vDt2)+1) & "/" & vDt2(LBound(vDt2))
            s_nfe_hr_emissao2 = Mid(vHr2(LBound(vHr2)), 1, 2) & Mid(vHr2(LBound(vHr2)), 4, 2) & Mid(vHr2(LBound(vHr2)), 7, 2)
			end if
		end if
    
    n = converte_numero(Request.Form("iQtdeItens"))

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
 
    dim v_item1, v_item2

   	redim v_item1(0)
	set v_item1(0) = New cl_ITEM_ESTOQUE_XML
    
    for i = 1 to n
        if Trim(Request.Form("c1_xml_prod_cProd_" & trim(cstr(i)))) <> "" then 
            if Trim(v_item1(ubound(v_item1)).xml_prod_cProd) <> "" then
				redim preserve v_item1(ubound(v_item1)+1)
				set v_item1(ubound(v_item1)) = New cl_ITEM_ESTOQUE_XML
				end if
			with v_item1(ubound(v_item1))
                .xml_prod_cProd  = Trim(Request.Form("c1_xml_prod_cProd_" & trim(cstr(i))))
                .xml_prod_cEAN  = Trim(Request.Form("c1_xml_prod_cEAN_" & trim(cstr(i))))
                .xml_prod__NCM  = Trim(Request.Form("c1_xml_prod__NCM_" & trim(cstr(i))))
                .xml_prod__CFOP  = Trim(Request.Form("c1_xml_prod__CFOP_" & trim(cstr(i))))
                .xml_prod__qCom  = Trim(Request.Form("c1_xml_prod__qCom_" & trim(cstr(i))))
                .xml_prod__vUnCom  = Trim(Request.Form("c1_xml_prod__vUnCom_" & trim(cstr(i))))
                .xml_prod__vProd  = Trim(Request.Form("c1_xml_prod__vProd_" & trim(cstr(i))))
                .xml_prod__vFrete  = Trim(Request.Form("c1_xml_prod__vFrete_" & trim(cstr(i))))
                .xml_imposto__pICMS  = Trim(Request.Form("c1_xml_imposto__pICMS_" & trim(cstr(i))))
                .xml_imposto__pIPI  = Trim(Request.Form("c1_xml_imposto__pIPI_" & trim(cstr(i))))
                .xml_imposto__vIPI  = Trim(Request.Form("c1_xml_imposto__vIPI_" & trim(cstr(i))))
                end with
            end if
        next

   	redim v_item2(0)
	set v_item2(0) = New cl_ITEM_ESTOQUE_XML
    
    for i = 1 to 19
        if Trim(Request.Form("c2_xml_prod_cProd_" & trim(cstr(i)))) <> "" then 
            if Trim(v_item2(ubound(v_item2)).xml_prod_cProd) <> "" then
				redim preserve v_item2(ubound(v_item2)+1)
				set v_item2(ubound(v_item2)) = New cl_ITEM_ESTOQUE_XML
				end if
			with v_item2(ubound(v_item2))
                .xml_prod_cProd  = Trim(Request.Form("c2_xml_prod_cProd_" & trim(cstr(i))))
                .xml_prod_cEAN  = Trim(Request.Form("c2_xml_prod_cEAN_" & trim(cstr(i))))
                .xml_prod__NCM  = Trim(Request.Form("c2_xml_prod__NCM_" & trim(cstr(i))))
                .xml_prod__CFOP  = Trim(Request.Form("c2_xml_prod__CFOP_" & trim(cstr(i))))
                .xml_prod__qCom  = Trim(Request.Form("c2_xml_prod__qCom_" & trim(cstr(i))))
                .xml_prod__vUnCom  = Trim(Request.Form("c2_xml_prod__vUnCom_" & trim(cstr(i))))
                .xml_prod__vProd  = Trim(Request.Form("c2_xml_prod__vProd_" & trim(cstr(i))))
                .xml_prod__vFrete  = Trim(Request.Form("c2_xml_prod__vFrete_" & trim(cstr(i))))
                .xml_imposto__pICMS  = Trim(Request.Form("c2_xml_imposto__pICMS_" & trim(cstr(i))))
                .xml_imposto__pIPI  = Trim(Request.Form("c2_xml_imposto__pIPI_" & trim(cstr(i))))
                .xml_imposto__vIPI  = Trim(Request.Form("c2_xml_imposto__vIPI_" & trim(cstr(i))))
                end with
            end if
        next


	set r_estoque = New cl_ESTOQUE_AGIO
	with r_estoque
		.data_entrada = Date
		.hora_entrada = retorna_so_digitos(formata_hora(Now))
		.data_ult_movimento = Date
		.fabricante = normaliza_codigo(retorna_so_digitos(Request.Form("c_fabricante")), TAM_MIN_FABRICANTE)
		.documento = Trim(Request.Form("c_documento"))
        .perc_agio = converte_numero(c_perc_agio)
        .data_emissao_NF_entrada = StrToDate(s_nfe_dt_hr_emissao)
		if CADASTRAR_WMS_CD_ENTRADA_ESTOQUE then
			.id_nfe_emitente = Trim(Request.Form("c_id_nfe_emitente"))
		else
			.id_nfe_emitente = 0
			end if
		.usuario = usuario
		.kit = 0
	'	ENTRADA ESPECIAL?
		s = Trim(Request.Form("ckb_especial"))
		if s <> "" then
			.entrada_especial = 1
		else
			.entrada_especial = 0
			end if
		.obs = Trim(Request.Form("c_obs"))
		end with
		
	redim v_item(0)
	set v_item(0) = New cl_ITEM_ESTOQUE_ENTRADA_XML
	'n = Request.Form("c_produto").Count
    'n = converte_numero(Request.Form("iQtdeItens"))
	for i = 1 to n
		s=Trim(Request.Form("c_nfe_nItem_" & Trim(i)))
        if s = "IMPORTA_S" then 
            s=Trim(Request.Form("c_erp_codigo_" & Trim(i)))
		    if s <> "" then
			    if Trim(v_item(ubound(v_item)).produto) <> "" then
				    redim preserve v_item(ubound(v_item)+1)
				    set v_item(ubound(v_item)) = New cl_ITEM_ESTOQUE_ENTRADA_XML
				    end if
			    with v_item(ubound(v_item))
				    .fabricante=r_estoque.fabricante
				    .produto=Ucase(Trim(Request.Form("c_erp_codigo_" & Trim(i))))
				    s = Trim(Request.Form("c_nfe_qtde_" & Trim(i)))
				    if IsNumeric(s) then .qtde = CLng(s) else .qtde = 0
			    '	PREÇO FABRICANTE
				    s = Trim(Request.Form("c_nfe_vl_unitario_nota_" & Trim(i)))
				    .preco_fabricante = converte_numero(s)
				    if .preco_fabricante < 0 then
					    alerta=texto_add_br(alerta)
					    alerta=alerta & "Produto " & .produto & " está com valor inválido: " & formata_moeda(.preco_fabricante)
					    end if
			    '	CUSTO 2
				    s = Trim(Request.Form("c_nfe_vl_unitario_" & Trim(i)))
				    .vl_custo2 = converte_numero(s)
				    if .vl_custo2 < 0 then
					    alerta=texto_add_br(alerta)
					    alerta=alerta & "Produto " & .produto & " está com Custo II inválido: " & formata_moeda(.vl_custo2)
					    end if
			    '	BASE CÁLCULO ICMS ST
				    's = Trim(Request.Form("c_vl_BC_ICMS_ST")(i))
				    '.vl_BC_ICMS_ST = converte_numero(s)
                    .vl_BC_ICMS_ST = 0
				    if .vl_BC_ICMS_ST < 0 then
					    alerta=texto_add_br(alerta)
					    alerta=alerta & "Produto " & .produto & " está com valor de base de cálculo do ICMS ST inválido: " & formata_moeda(.vl_BC_ICMS_ST)
					    end if
			    '	VALOR DO ICMS ST
				    's = Trim(Request.Form("c_vl_ICMS_ST")(i))
				    '.vl_ICMS_ST = converte_numero(s)
                    .vl_ICMS_ST = 0
				    if .vl_ICMS_ST < 0 then
					    alerta=texto_add_br(alerta)
					    alerta=alerta & "Produto " & .produto & " está com valor do ICMS ST inválido: " & formata_moeda(.vl_ICMS_ST)
					    end if
			    '	NCM
				    .ncm = Trim(Request.Form("c_nfe_ncm_sh_" & Trim(i)))
				    .ncm_redigite = Trim(Request.Form("c_nfe_ncm_sh_" & Trim(i)))
				    if .ncm = "" then
					    alerta=texto_add_br(alerta)
					    alerta=alerta & "Produto " & .produto & ": informe o NCM"
				    elseif (Len(.ncm) <> 2) And (Len(.ncm) <> 8) then
					    alerta=texto_add_br(alerta)
					    alerta=alerta & "Produto " & .produto & ": NCM com tamanho inválido"
				    elseif .ncm <> .ncm_redigite then
					    alerta=texto_add_br(alerta)
					    alerta=alerta & "Produto " & .produto & ": falha na conferência do NCM redigitado"
					    end if
			    '	CST
				    .cst = Trim(Request.Form("c_nfe_cst_" & Trim(i)))
				    .cst_redigite = Trim(Request.Form("c_nfe_cst_" & Trim(i)))
				    if .cst = "" then
					    alerta=texto_add_br(alerta)
					    alerta=alerta & "Produto " & .produto & ": informe o CST"
				    elseif Len(.cst) <> 3 then
					    alerta=texto_add_br(alerta)
					    alerta=alerta & "Produto " & .produto & ": CST com tamanho inválido"
				    elseif .cst <> .cst_redigite then
					    alerta=texto_add_br(alerta)
					    alerta=alerta & "Produto " & .produto & ": falha na conferência do CST redigitado"
					    end if
                '   EAN e EAN ORIGINAL
			        .ean = Trim(Request.Form("c_ean_" & Trim(i)))
                    .ean_original = Trim(Request.Form("c_ean_" & Trim(i)))
               '   CÓDIGO PRODUTO NO ARQUIVO XML
			        .produto_xml = Trim(Request.Form("c_nfe_codigo_" & Trim(i)))
               '   ALÍQUOTA IPI
                    s = Trim(Request.Form("c_nfe_aliq_ipi_" & Trim(i)))
                    if s = "" then s = "0"
			        .aliq_ipi = s
               '   ALÍQUOTA ICMS
                    s = Trim(Request.Form("c_nfe_aliq_icms_" & Trim(i)))
                    if s = "" then s = "0"
			        .aliq_icms = s
               '   VALOR IPI 
                    s = Trim(Request.Form("c_nfe_vl_ipi_" & Trim(i))) 
                    if s = "" then s = "0"
			        .vl_ipi = s               
               '   VALOR FRETE
                    s = Trim(Request.Form("c_nfe_vl_frete_" & Trim(i))) 
                    if s = "" then s = "0"
			        .vl_frete = s               
				    end with
			    end if
            end if
		next

	if alerta = "" then
	'	VERIFICA SE ESTAS MERCADORIAS JÁ FORAM GRAVADAS!!
		dim estoque_a, vjg
		s = "SELECT t_ESTOQUE.id_estoque, produto, qtde FROM t_ESTOQUE INNER JOIN t_ESTOQUE_ITEM ON (t_ESTOQUE.id_estoque=t_ESTOQUE_ITEM.id_estoque)" & _
			" WHERE (t_ESTOQUE.fabricante='" & r_estoque.fabricante & "')" & _
			" AND (usuario='" & usuario & "')" & _
			" AND (data_entrada=" & bd_formata_data(Date) & ")" & _
			" AND (hora_entrada >= '" & formata_hora_hhnnss(Now-converte_min_to_dec(10))& "')" & _
			" AND (documento='" & r_estoque.documento & "')" & _
			" ORDER BY t_ESTOQUE_ITEM.id_estoque, sequencia"
		set rs = cn.execute(s)
		redim vjg(0)
		set vjg(ubound(vjg)) = New cl_DUAS_COLUNAS
		vjg(ubound(vjg)).c1=""
		estoque_a="--XX--"
		do while Not rs.EOF 
			if estoque_a<>Trim("" & rs("id_estoque")) then
				estoque_a=Trim("" & rs("id_estoque"))
				if vjg(ubound(vjg)).c1 <> "" then 
					redim preserve vjg(ubound(vjg)+1)
					set vjg(ubound(vjg)) = New cl_DUAS_COLUNAS
					vjg(ubound(vjg)).c1=""
					end if
				vjg(ubound(vjg)).c2=estoque_a
				end if
			
			vjg(ubound(vjg)).c1=vjg(ubound(vjg)).c1 & Trim("" & rs("produto")) & "|" & Trim("" & rs("qtde")) & "|"
			rs.MoveNext 
			Loop
		
		if rs.State <> 0 then rs.Close
		
		s=""
		for i=Lbound(v_item) to Ubound(v_item)
			with v_item(i)
				if .produto<>"" then
					s=s & .produto & "|" & Cstr(.qtde) & "|"
					end if
				end with
			next
    
		for i=Lbound(vjg) to Ubound(vjg)
			if s=vjg(i).c1 then
				alerta="Esta operação de entrada de mercadorias no estoque já foi gravada com a identificação " & vjg(i).c2
                exit for
				end if
			next
		end if

	if alerta = "" then
	'	INFORMAÇÕES PARA O LOG
		s_log = ""
		for i = Lbound(v_item) to Ubound(v_item)
			with v_item(i)
				if .produto <> "" then
					s_log = s_log & log_estoque_monta_incremento(.qtde, "", .produto) & _
							"(" & formata_moeda(.preco_fabricante) & "; " & formata_moeda(.vl_custo2) & _
							"; ST: " & formata_moeda(.vl_BC_ICMS_ST) & "; " & formata_moeda(.vl_ICMS_ST) & _
							"; NCM: " & .ncm & "; " & _
							"; CST: " & .cst & ")"
					end if
				end with
			next

		s = "Entrada via XML no estoque de mercadorias do fabricante=" & Trim(r_estoque.fabricante) & "," & _
			" documento=" & Trim(r_estoque.documento)
		if r_estoque.entrada_especial <> 0 then s = s & ", registrado como entrada especial"
		s_log = s & ":" & s_log
		
	'	~~~~~~~~~~~~~
		cn.BeginTrans
	'	~~~~~~~~~~~~~
        if Not estoque_nova_entrada_mercadorias_xml(r_estoque, v_item, msg_erro) then 
		'	~~~~~~~~~~~~~~~~
			cn.RollbackTrans
		'	~~~~~~~~~~~~~~~~
			Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_MOVIMENTO_ESTOQUE)
			end if

    ' 	GRAVA O EAN NA TABELA DE PRODUTOS
        msg_erro = ""
	    For i = LBound(v_item) To UBound(v_item)
		    With v_item(i)
			    If (.ean <> .ean_original) Then
				    s_sql = " UPDATE T_PRODUTO" & _
						    " SET ean = '" & .ean & "' " & _
						    " WHERE fabricante = '" & .fabricante & "'" & _
                           " AND produto = '" & .produto & "'" 
				    cn.Execute(s_sql)
				    if Err <> 0 then
					    msg_erro=Cstr(Err) & ": " & Err.Description
					    exit for
					    end if				
				   End If
			    End With
		    Next
       if msg_erro <> "" then
		'	~~~~~~~~~~~~~~~~
			cn.RollbackTrans
		'	~~~~~~~~~~~~~~~~
			Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_MOVIMENTO_ESTOQUE)
            end if

      ' 	GRAVA OS DADOS DO XML NAS TABELAS T_ESTOQUE_XML E T_ESTOQUE_XML_ITEM
      '---- primeiro xml  
	     If uploaded_file_guid <> "" Then

     		'	GERA O NSU PARA O NOVO REGISTRO
			if Not fin_gera_nsu(T_ESTOQUE_XML, id_estoque_xml, msg_erro) then
				alerta = "FALHA AO GERAR NSU PARA O NOVO REGISTRO (" & msg_erro & ")"
			else
				if id_estoque_xml <= 0 then
					alerta = "NSU GERADO É INVÁLIDO (" & id_estoque_xml & ")"
					end if
				end if
       
			    s_sql = " INSERT INTO T_ESTOQUE_XML " & _
					 " (id, id_estoque, xml_data_import, xml_hora_import, xml_prioridade, xml_usuario, xml_ide__cNF, " & _
                     "  xml_ide__serie, xml_ide__nNF, xml_emit__CNPJ, xml_emit__xNome, xml_dest__CNPJ, xml_dest__xNome, " & _
                     "  xml_data_emissao_NF_entrada, xml_hora_emissao_NF_entrada, xml_conteudo " & _   
                     "  ) VALUES " & _
                     " ("  & _
	                 CStr(id_estoque_xml) & ", " & _   
    				 " '" & r_estoque.id_estoque & "', " & _
                     bd_formata_data(Date) & ", " & _
                     " '" & formata_hora_hhnnss(Now-converte_min_to_dec(10)) & "', " & _
                     " " & "1" & ", " & _
                     " '" & usuario & "', " & _
                     " '" & c_xml_ide__cNF_1  & "', " & _
                     " '" & c_xml_ide__serie_1  & "', " & _
                     " '" & c_xml_ide__nNF_1 & "', " & _
                     " '" & c_xml_emit__CNPJ_1 & "', " & _
                     " '" & QuotedStr(c_xml_emit__xNome_1) & "', " & _
                     " '" & c_xml_dest__CNPJ_1 & "', " & _
                     " '" & QuotedStr(c_xml_dest__xNome_1) & "', " & _
                     bd_formata_data(StrToDate(s_nfe_dt_hr_emissao)) & ", " & _
                     " '" & s_nfe_hr_emissao & "', " & _
                     "(select file_content_text from t_UPLOAD_FILE where guid = '" & uploaded_file_guid & "') " & _
                     " )" 
			     cn.Execute(s_sql)
			     if Err <> 0 then
                     msg_erro= "Problema na inclusão do XML" & vbCrLf
				     msg_erro= msg_erro & Cstr(Err) & ": " & Err.Description
				     end if				

                s_sql = " UPDATE t_UPLOAD_FILE" & _
						    " SET st_confirmation_ok = 1" & _
						    " WHERE guid = '" & uploaded_file_guid & "' " 
				    cn.Execute(s_sql)
				    if Err <> 0 then
					    msg_erro=Cstr(Err) & ": " & Err.Description
					    end if				
            
                for i=lbound(v_item1) to ubound(v_item1)
                    with v_item1(i)
			            s_sql = " INSERT INTO T_ESTOQUE_XML_ITEM " & _
					         " (id_estoque_xml, xml_prod_cProd, xml_prod_cEAN, xml_prod__NCM, xml_prod__CFOP, " & _
                             " xml_prod__qCom, xml_prod__vUnCom, xml_prod__vProd, xml_prod__vFrete, xml_imposto__pICMS, " & _
                             " xml_imposto__pIPI, xml_imposto__vIPI  " & _
                             "  ) VALUES " & _
                             " ("  & _
	                         CStr(id_estoque_xml) & ", " & _   
    				         " '" & .xml_prod_cProd & "', " & _
                             " '" & .xml_prod_cEAN & "', " & _
                             " '" & .xml_prod__NCM & "', " & _
                             " '" & .xml_prod__CFOP  & "', " & _
                             " '" & .xml_prod__qCom  & "', " & _
                             " '" & .xml_prod__vUnCom & "', " & _
                             " '" & .xml_prod__vProd & "', " & _
                             " '" & .xml_prod__vFrete & "', " & _
                             " '" & .xml_imposto__pICMS & "', " & _
                             " '" & .xml_imposto__pIPI & "', " & _
                             " '" & .xml_imposto__vIPI & "' " & _
                             " )" 
			                 cn.Execute(s_sql)
                             end with
			             if Err <> 0 then
                             msg_erro= "Problema na inclusão dos itens XML" & vbCrLf
				             msg_erro= msg_erro & Cstr(Err) & ": " & Err.Description
				             end if				
                    next
            
		     End if
        if msg_erro <> "" then
		 '	~~~~~~~~~~~~~~~~
			 cn.RollbackTrans
		 '	~~~~~~~~~~~~~~~~
			 Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_MOVIMENTO_ESTOQUE)
             end if

        '---- segundo xml, se houver
	     If uploaded_file_guid2 <> "" Then

     		'	GERA O NSU PARA O NOVO REGISTRO
			if Not fin_gera_nsu(T_ESTOQUE_XML, id_estoque_xml, msg_erro) then
				alerta = "FALHA AO GERAR NSU PARA O NOVO REGISTRO (" & msg_erro & ")"
			else
				if id_estoque_xml <= 0 then
					alerta = "NSU GERADO É INVÁLIDO (" & id_estoque_xml & ")"
					end if
				end if
       
			    s_sql = " INSERT INTO T_ESTOQUE_XML " & _
					 " (id, id_estoque, xml_data_import, xml_hora_import, xml_prioridade, xml_usuario, xml_ide__cNF, " & _
                     "  xml_ide__serie, xml_ide__nNF, xml_emit__CNPJ, xml_emit__xNome, xml_dest__CNPJ, xml_dest__xNome, " & _
                     "  xml_data_emissao_NF_entrada, xml_hora_emissao_NF_entrada, xml_conteudo " & _   
                     "  ) VALUES " & _
                     " ("  & _
	                 CStr(id_estoque_xml) & ", " & _   
    				 " '" & r_estoque.id_estoque & "', " & _
                     bd_formata_data(Date) & ", " & _
                     " '" & formata_hora_hhnnss(Now-converte_min_to_dec(10)) & "', " & _
                     " " & "2" & ", " & _
                     " '" & usuario & "', " & _
                     " '" & c_xml_ide__cNF_2  & "', " & _
                     " '" & c_xml_ide__serie_2  & "', " & _
                     " '" & c_xml_ide__nNF_2 & "', " & _
                     " '" & c_xml_emit__CNPJ_2 & "', " & _
                     " '" & QuotedStr(c_xml_emit__xNome_2) & "', " & _
                     " '" & c_xml_dest__CNPJ_2 & "', " & _
                     " '" & QuotedStr(c_xml_dest__xNome_2) & "', " & _
                     bd_formata_data(StrToDate(s_nfe_dt_hr_emissao2)) & ", " & _
                     " '" & s_nfe_hr_emissao2 & "', " & _
                     "(select file_content_text from t_UPLOAD_FILE where guid = '" & uploaded_file_guid2 & "') " & _
                     " )" 
			     cn.Execute(s_sql)
			     if Err <> 0 then
                     msg_erro= "Problema na inclusão do XML" & vbCrLf
				     msg_erro= msg_erro & Cstr(Err) & ": " & Err.Description
				     end if				

                s_sql = " UPDATE t_UPLOAD_FILE" & _
						    " SET st_confirmation_ok = 1" & _
						    " WHERE guid = '" & uploaded_file_guid2 & "' " 
				    cn.Execute(s_sql)
				    if Err <> 0 then
					    msg_erro=Cstr(Err) & ": " & Err.Description
					    end if				
            
                for i=lbound(v_item2) to ubound(v_item2)
                    with v_item2(i)
			            s_sql = " INSERT INTO T_ESTOQUE_XML_ITEM " & _
					         " (id_estoque_xml, xml_prod_cProd, xml_prod_cEAN, xml_prod__NCM, xml_prod__CFOP, " & _
                             " xml_prod__qCom, xml_prod__vUnCom, xml_prod__vProd, xml_prod__vFrete, xml_imposto__pICMS, " & _
                             " xml_imposto__pIPI, xml_imposto__vIPI " & _
                             "  ) VALUES " & _
                             " ("  & _
	                         CStr(id_estoque_xml) & ", " & _   
    				         " '" & .xml_prod_cProd & "', " & _
                             " '" & .xml_prod_cEAN & "', " & _
                             " '" & .xml_prod__NCM & "', " & _
                             " '" & .xml_prod__CFOP  & "', " & _
                             " '" & .xml_prod__qCom  & "', " & _
                             " '" & .xml_prod__vUnCom & "', " & _
                             " '" & .xml_prod__vProd & "', " & _
                             " '" & .xml_prod__vFrete & "', " & _
                             " '" & .xml_imposto__pICMS & "', " & _
                             " '" & .xml_imposto__pIPI & "', " & _
                             " '" & .xml_imposto__vIPI & "' " & _
                             " )" 
			                 cn.Execute(s_sql)
                             end with
			             if Err <> 0 then
                             msg_erro= "Problema na inclusão dos itens XML" & vbCrLf
				             msg_erro= msg_erro & Cstr(Err) & ": " & Err.Description
				             end if				
                    next
            
		     End if
        if msg_erro <> "" then
		 '	~~~~~~~~~~~~~~~~
			 cn.RollbackTrans
		 '	~~~~~~~~~~~~~~~~
			 Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_MOVIMENTO_ESTOQUE)
             end if
    
   
        		
		s_log = s_log & "; registrado com nº " & r_estoque.id_estoque
		s_log = s_log & "; obs=" & formata_texto_log(r_estoque.obs)
		s_log = s_log & "; id_nfe_emitente=" & r_estoque.id_nfe_emitente
		if c_log_edicao <> "" then s_log = s_log & chr(13) & c_log_edicao
		grava_log usuario, "", "", "", OP_LOG_ESTOQUE_ENTRADA, s_log
		
	'	PROCESSA OS PRODUTOS VENDIDOS SEM PRESENÇA NO ESTOQUE
		if Not estoque_processa_produtos_vendidos_sem_presenca_v2(r_estoque.id_nfe_emitente, usuario, msg_erro) then
		'	~~~~~~~~~~~~~~~~
			cn.RollbackTrans
		'	~~~~~~~~~~~~~~~~
			Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_MOVIMENTO_ESTOQUE)
			end if
		
	'	~~~~~~~~~~~~~~
		cn.CommitTrans
	'	~~~~~~~~~~~~~~
		if Err=0 then
			Response.Redirect("estoqueconsultaxml.asp?estoque_selecionado=" & r_estoque.id_estoque  & "&url_back=X" & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")))
		else
			alerta=Cstr(Err) & ": " & Err.Description
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

<% if alerta <> "" then %>
<!-- ************************************************************ -->
<!-- **********  PÁGINA PARA EXIBIR MENSAGENS DE ERRO  ********** -->
<!-- ************************************************************ -->
<body onload="bVOLTAR.focus();">
<center>
<!-- ********** <input type="hidden" name="c_perc_agio" id="c_perc_agio" value="<%=c_perc_agio%>">	********** -->
<br>
<!--  T E L A  -->
<p class="T">A V I S O</p>
<div class="MtAlerta" style="width:600px;font-weight:bold;" align="center"><p style='margin:5px 2px 5px 2px;'><%=alerta%></p></div>
<br><br>
<p class="TracoBottom"></p>
<table cellSpacing="0">
<tr>
	<td align="center"><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()"><img src="..\botao\voltar.gif" width="176" height="55" border="0"></a></td>
</tr>
</table>
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