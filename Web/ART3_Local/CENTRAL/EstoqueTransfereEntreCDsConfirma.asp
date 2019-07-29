<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/estoque.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'      ========================================================================
'       E S T O Q U E T R A N S F E R E E N T R E C D S C O N F I R M A . A S P
'      ========================================================================
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

    class cl_ESTOQUE_TRANSFERENCIA_ITEM
        dim documento
        dim id_estoque_origem
        dim entrada_tipo
        dim fabricante
        dim produto
        dim descricao_html
        dim qtde
        dim preco_fabricante
        dim vl_custo2
        dim vl_BC_ICMS_ST
        dim vl_ICMS_ST
        dim ncm
        dim cst
        dim st_ncm_cst_herdado_tabela_produto
        dim ean
        dim aliq_ipi
        dim aliq_icms
        dim vl_ipi
        dim preco_origem
        dim produto_xml
        end class
	
    dim s, s_log, i, n, usuario, msg_erro, c_log_edicao
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 
	
	dim alerta
	alerta=""

    dim s_sql    

'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs
    dim c_nfe_emitente_origem, c_nfe_emitente_destino, c_documento_transf, c_obs

	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	
	c_log_edicao = Trim(Request.Form("c_log_edicao"))
    
   	c_nfe_emitente_origem = Trim(Request.Form("c_nfe_emitente_origem"))
	c_nfe_emitente_destino = Trim(Request.Form("c_nfe_emitente_destino"))
    c_documento_transf = Trim(Request.Form("c_documento_transf"))
    c_obs = Trim(Request.Form("c_obs"))

    dim v_item1, v_item2, v_item3
    dim s_cod_prod1, s_cod_prod2
    dim id_estoque_transferencia

    'Procedimento: 
    '- obter as informações da tela anterior e armazenar no vetor 1
    '- chamar novamente a rotina de montagem e armazenar no vetor 2
    '- fazer a comparação entre os vetores; só gravar se os valores baterem
    '(OBS - COLETAR NA TABELA DA TELA ANTERIOR O ID_ESTOQUE_ORIGEM)

    n = Request.Form("c_produto").Count

   	redim v_item1(0)
	set v_item1(0) = New cl_ESTOQUE_TRANSFERENCIA_ITEM    
    for i = 1 to n 
        if Trim(Request.Form("c_produto")(i)) <> "" then 
            if Trim(v_item1(ubound(v_item1)).produto) <> "" then
				redim preserve v_item1(ubound(v_item1)+1)
				set v_item1(ubound(v_item1)) = New cl_ESTOQUE_TRANSFERENCIA_ITEM
				end if
			with v_item1(ubound(v_item1))
                .documento = Trim(Request.Form("c_documento")(i))
                .id_estoque_origem = Trim(Request.Form("c_id_estoque_origem")(i))
                .fabricante = Trim(Request.Form("c_fabricante")(i))
                .produto = Trim(Request.Form("c_produto")(i))
                .qtde  = CInt(Trim(Request.Form("c_qtde")(i)))
                .vl_custo2 = Trim(Request.Form("c_vl_custo2")(i))                
                end with
            end if
        next

   	redim v_item2(0)
	set v_item2(0) = New cl_ESTOQUE_TRANSFERENCIA_ITEM
    
    if not estoque_produto_transf_consiste_quantidades(c_nfe_emitente_origem, _
                                                    c_nfe_emitente_destino, _
                                                    v_item1, _
                                                    v_item2, _
										            msg_erro) then
        alerta = msg_erro
        end if
	
	if alerta = "" then
	'	INFORMAÇÕES PARA O LOG
		s_log = ""
		for i = Lbound(v_item2) to Ubound(v_item2)
			with v_item2(i)
				if .produto <> "" then
					s_log = s_log & log_estoque_monta_incremento(.qtde, "", .produto) & _
							"(" & formata_moeda(.preco_fabricante) & "; " & formata_moeda(.vl_custo2) & _
							"; NCM: " & .ncm & "; " & _
							"; CST: " & .cst & "; " & _
							"; % IPI: " & .aliq_ipi & "; " & _
							"; % ICMS: " & .aliq_icms & "; " & _
							"; VL IPI: " & formata_moeda(.vl_ipi) & ")"
					end if
				end with
			next

		s = "Transferência entre estoques do CD=" & c_nfe_emitente_origem & "," & _
			" para o CD=" & c_nfe_emitente_destino & "," & _
			" documento=" & c_documento_transf
		s_log = s & ":" & s_log
		
	'	~~~~~~~~~~~~~
		cn.BeginTrans
	'	~~~~~~~~~~~~~
        
        redim v_item3(0)
	    set v_item3(0) = New cl_ESTOQUE_TRANSFERENCIA_ITEM    
    
        if Not estoque_produto_transf_consiste_quantidades(c_nfe_emitente_origem, _
                                                    c_nfe_emitente_destino, _
                                                    v_item2, _
                                                    v_item3, _
										            msg_erro) then
		'	~~~~~~~~~~~~~~~~
			cn.RollbackTrans
		'	~~~~~~~~~~~~~~~~
			Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_TRANSFERENCIA_CD_CONFERE)
            end if

    ' 	GRAVA OS DADOS NAS TABELAS t_ESTOQUE_TRANSFERENCIA E t_ESTOQUE_TRANSFERENCIA_ITEM
    '	GERA O NSU PARA O NOVO REGISTRO
        'lhgx criar constante t_estoque_transferencia
		if Not fin_gera_nsu("T_ESTOQUE_TRANSFERENCIA", id_estoque_transferencia, msg_erro) then
			alerta = "FALHA AO GERAR NSU PARA O NOVO REGISTRO (" & msg_erro & ")"
		else
			if id_estoque_transferencia <= 0 then
				alerta = "NSU GERADO É INVÁLIDO (" & id_estoque_transferencia & ")"
				end if
			end if
       
			s_sql = " INSERT INTO T_ESTOQUE_TRANSFERENCIA " & _
					" (id, id_nfe_emitente_origem, id_nfe_emitente_destino, documento, data, hora, usuario, obs" & _   
                    "  ) VALUES " & _
                    " ("  & _
	                CStr(id_estoque_transferencia) & ", " & _   
    				" " & c_nfe_emitente_origem & ", " & _
                    " " & c_nfe_emitente_destino & ", " & _
                    " '" & c_documento_transf  & "', " & _
                    bd_formata_data(Date) & ", " & _
                    bd_formata_data_hora(Date) & ", " & _
                    " '" & usuario & "', " & _
                    " '" & c_obs & "' " & _
                    " )" 
			    cn.Execute(s_sql)
			    if Err <> 0 then
                    msg_erro= "Problema na inclusão da transferência" & vbCrLf
				    msg_erro= msg_erro & Cstr(Err) & ": " & Err.Description
				    end if				

            if msg_erro = "" then
                for i=lbound(v_item2) to ubound(v_item2)
                    with v_item2(i)
			            s_sql = " INSERT INTO T_ESTOQUE_TRANSFERENCIA_ITEM " & _
					            " (id_estoque_transferencia, id_estoque_origem, entrada_tipo, documento, fabricante, produto, qtde, preco_fabricante, vl_custo2, vl_BC_ICMS_ST, vl_ICMS_ST,  " & _
                                " ncm, cst, st_ncm_cst_herdado_tabela_produto, ean, aliq_ipi, aliq_icms, vl_ipi, preco_origem, produto_xml " & _
                                "  ) VALUES " & _
                                " ("  & _
	                            CStr(id_estoque_transferencia) & ", " & _   
    				            .id_estoque_origem & ", " & _
                                " " & .entrada_tipo & ", " & _
                                " '" & .documento & "', " & _
                                " '" & .fabricante & "', " & _
                                " '" & .produto  & "', " & _
                                .qtde  & ", " & _
                                Iif(IsNull(.preco_fabricante), "NULL", bd_formata_numero(.preco_fabricante)) & ", " & _
                                Iif(IsNull(.vl_custo2), "NULL", bd_formata_numero(.vl_custo2)) & ", " & _
                                Iif(IsNull(.vl_BC_ICMS_ST), "NULL", bd_formata_numero(.vl_BC_ICMS_ST)) & ", " & _
                                Iif(IsNull(.vl_ICMS_ST), "NULL", bd_formata_numero(.vl_ICMS_ST)) & ", " & _
                                " '" & .ncm & "', " & _
                                " '" & .cst & "', " & _
                                .st_ncm_cst_herdado_tabela_produto & ", " & _
                                " '" & .ean & "', " & _
                                Iif(IsNull(.aliq_ipi), "NULL", bd_formata_numero(.aliq_ipi)) & ", " & _
                                Iif(IsNull(.aliq_icms), "NULL", bd_formata_numero(.aliq_icms)) & ", " & _
                                Iif(IsNull(.vl_ipi), "NULL", bd_formata_numero(.vl_ipi)) & ", " & _
                                " '" & .preco_origem & "', " & _
                                " '" & .produto_xml & "' " & _
                                " )" 
			            cn.Execute(s_sql)
                        end with
			        if Err <> 0 then
                        msg_erro= "Problema na inclusão dos itens da transferência" & vbCrLf
				        msg_erro= msg_erro & Cstr(Err) & ": " & Err.Description
				        end if				
                    next
                end if
                
        if msg_erro <> "" then
		 '	~~~~~~~~~~~~~~~~
			 cn.RollbackTrans
		 '	~~~~~~~~~~~~~~~~
			 Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_TRANSFERENCIA_CD_CONFERE)
             end if
        		
		s_log = s_log & "; registrada com id " & CStr(id_estoque_transferencia)
		if c_log_edicao <> "" then s_log = s_log & chr(13) & c_log_edicao
		grava_log usuario, "", "", "", "ESTOQUE TRANSF CDS", s_log
				
	'	~~~~~~~~~~~~~~
		cn.CommitTrans
	'	~~~~~~~~~~~~~~
		if Err=0 then
			'Response.Redirect("estoqueconsultaxml.asp?estoque_selecionado=" & CStr(id_estoque_transferencia)  & "&url_back=X" & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")))
            Response.Redirect("estoquetransfereentrecdsconsulta.asp?transf_selecionada=" & CStr(id_estoque_transferencia))
		else
			alerta=Cstr(Err) & ": " & Err.Description
			end if
		end if



' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________
' --------------------------------------------------------------------
'   ESTOQUE PRODUTO TRANSFERÊNCIA ENTRE CD (COMPARAÇÃO DE QUANTIDADES)
'   Retorno da função:
'      False - Ocorreu falha ao tentar alterar os dados do estoque.
'      True - Conseguiu alterar os dados do estoque.
'   Esta função verifica se a quantidade permanece a mesma antes de gravar
'   as informações.
'	
'	(melhorar esta descrição)
'	
function estoque_produto_transf_consiste_quantidades(ByVal id_nfe_emitente_origem, _
                                                    ByVal id_nfe_emitente_destino, _
                                                    ByRef v_1, _
                                                    ByRef v_2, _
										            ByRef msg_erro)
dim s
dim s_cod_prod1
dim s_cod_prod2
dim iv

	estoque_produto_transf_consiste_quantidades = False
	
    msg_erro = ""	
    s_cod_prod1 = ""
    s_cod_prod2 = ""

    for iv = LBound(v_1) to UBound(v_1)
        with v_1(iv)
            s_cod_prod1 = .produto
            if s_cod_prod1 <> s_cod_prod2 then                
                if not estoque_produto_transf_cd_monta(id_nfe_emitente_origem, _
                                                        v_1(iv).fabricante, _
                                                        v_1(iv).produto, _
                                                        v_1(iv).qtde, _
                                                        v_2, _
                                                        s) then
                    msg_erro = "Problemas na conferência da transferência: " & s
                    exit function
                    end if
                end if
            s_cod_prod2 = s_cod_prod1
            end with
        next

    if alerta = "" then
        'primeira comparação: se os vetores não tiverem o mesmo número de elementos, os estoques mudaram
        if UBound(v_1) <> UBound(v_2) then
            msg_erro = "Mudanças nas quantidades dos estoques, reiniciar o processo!!!"
        else
        'segunda comparação: se os conteúdos dos vetores forem diferentes, os estoques mudaram
            for iv = LBound(v_1) to UBound(v_2)
                if (v_1(iv).id_estoque_origem <> v_2(iv).id_estoque_origem) Or _
                    (v_1(iv).fabricante <> v_2(iv).fabricante) Or _
                    (v_1(iv).produto <> v_2(iv).produto) Or _
                    (v_1(iv).qtde <> v_2(iv).qtde) then
                    msg_erro = "Mudanças em características dos estoques, reiniciar o processo!!!"
                    exit function
                    end if
                next
            end if
        end if	
	
	estoque_produto_transf_consiste_quantidades = True
end function

'-----------------------------------------------------------------------------------------------------------------'

function estoque_produto_transf_cd_monta(ByVal id_nfe_emitente, _
										ByVal id_fabricante, ByVal id_produto, ByVal qtde_a_sair, _
										ByRef v_item, _
										ByRef msg_erro)
dim s
dim rs
dim qtde_disponivel
dim qtde_movimentada
dim v_estoque
dim v_documento
dim v_entrada_tipo
dim iv
dim descricao_html_aux
dim qtde_aux
dim qtde_utilizada_aux
dim preco_fabricante_aux
dim vl_custo2_aux
dim vl_BC_ICMS_ST_aux
dim vl_ICMS_ST_aux
dim ncm_aux
dim cst_aux
dim st_ncm_cst_herdado_tabela_produto_aux
dim ean_aux
dim aliq_ipi_aux
dim aliq_icms_aux
dim vl_ipi_aux
dim preco_origem_aux
dim produto_xml_aux
dim qtde_movto
dim s_chave

	estoque_produto_transf_cd_monta = False
	msg_erro = ""
	
'	NENHUMA UNIDADE SERÁ RETIRADA!!
	If (qtde_a_sair<=0) Or (Trim(id_produto)="") Then
		estoque_produto_transf_cd_monta = True
		Exit Function
		End If

	if Not cria_recordset_pessimista(rs, msg_erro) then exit function

'	OBTÉM OS "LOTES" DO PRODUTO DISPONÍVEIS NO ESTOQUE (POLÍTICA FIFO)
	s = "SELECT" & _
			" t_ESTOQUE.id_estoque, t_ESTOQUE.documento, t_ESTOQUE.entrada_tipo, (qtde-qtde_utilizada) AS saldo" & _
		" FROM t_ESTOQUE INNER JOIN t_ESTOQUE_ITEM ON (t_ESTOQUE.id_estoque=t_ESTOQUE_ITEM.id_estoque)" & _
		" WHERE" & _
			" (t_ESTOQUE.id_nfe_emitente = " & Trim("" & id_nfe_emitente) & ")" & _
			" AND (t_ESTOQUE_ITEM.fabricante='" & id_fabricante & "')" & _
			" AND (produto='" & id_produto & "')" & _
			" AND ((qtde-qtde_utilizada) > 0)" & _
		" ORDER BY" & _
			" data_entrada, t_ESTOQUE.id_estoque"
	rs.open s, cn

	qtde_disponivel = 0
	ReDim v_estoque(0)
	ReDim v_documento(0)
    ReDim v_entrada_tipo(0)
	v_estoque(UBound(v_estoque)) = ""
	v_documento(UBound(v_documento)) = ""
    v_entrada_tipo(UBound(v_entrada_tipo)) = 0

	do while Not rs.Eof
	'	ARMAZENA AS ENTRADAS NO ESTOQUE CANDIDATAS À SAÍDA DE PRODUTOS
		If v_estoque(UBound(v_estoque)) <> "" Then
			ReDim Preserve v_estoque(UBound(v_estoque) + 1)
			v_estoque(UBound(v_estoque)) = ""
			End If
		v_estoque(UBound(v_estoque)) = Trim("" & rs("id_estoque"))
		If v_documento(UBound(v_documento)) <> "" Then
			ReDim Preserve v_documento(UBound(v_documento) + 1)
			v_documento(UBound(v_documento)) = ""
			End If
		v_documento(UBound(v_documento)) = Trim("" & rs("documento"))
        If v_entrada_tipo(UBound(v_entrada_tipo)) <> "" Then
			ReDim Preserve v_entrada_tipo(UBound(v_entrada_tipo) + 1)
			v_entrada_tipo(UBound(v_entrada_tipo)) = ""
			End If
		v_entrada_tipo(UBound(v_entrada_tipo)) = rs("entrada_tipo")
		qtde_disponivel = qtde_disponivel + CLng(rs("saldo"))
		rs.movenext
		loop

'	NÃO HÁ PRODUTOS SUFICIENTES NO ESTOQUE!!
	If qtde_a_sair > qtde_disponivel Then
		msg_erro = "Produto " & id_produto & " do fabricante " & id_fabricante & ": faltam " & _
					formata_inteiro(qtde_a_sair-qtde_disponivel) & " unidades no estoque (" & obtem_apelido_empresa_NFe_emitente(id_nfe_emitente) & ")."
		Exit Function
		End If

'	SIMULA A SAÍDA DO ESTOQUE!!
	qtde_movimentada = 0
	For iv = LBound(v_estoque) To UBound(v_estoque)
	
		If Trim(v_estoque(iv)) <> "" Then
		
		'	A QUANTIDADE NECESSÁRIA JÁ FOI RETIRADA DO ESTOQUE!!
			If qtde_movimentada >= qtde_a_sair Then Exit For
			
		'	T_ESTOQUE_ITEM: SAÍDA DE PRODUTOS
			s = "SELECT " & _
					"ei.*, p.descricao_html, p.ean as ean_produto" & _
				" FROM t_ESTOQUE_ITEM ei" & _
                " INNER JOIN t_PRODUTO p ON (ei.fabricante = p.fabricante AND ei.produto = p.produto)" & _
				" WHERE" & _
					" (ei.id_estoque = '" & Trim(v_estoque(iv)) & "')" & _
					" AND (ei.fabricante = '" & id_fabricante & "')" & _
					" AND (ei.produto = '" & id_produto & "')"
			if rs.State <> 0 then rs.Close
			rs.open s, cn
			if Err <> 0 then
				msg_erro=Cstr(Err) & ": " & Err.Description
				exit function
				end if

			if rs.Eof then
				msg_erro = "Falha ao acessar o registro no estoque do produto " & id_produto & " do fabricante " & id_fabricante & " (id_estoque = '" & Trim(v_estoque(iv)) & "')"
				Exit Function
			else
				qtde_aux = rs("qtde")
				qtde_utilizada_aux = rs("qtde_utilizada")
				preco_fabricante_aux = rs("preco_fabricante")
				vl_custo2_aux = rs("vl_custo2")
                vl_BC_ICMS_ST_aux = rs("vl_BC_ICMS_ST")
                vl_ICMS_ST_aux = rs("vl_ICMS_ST")
				ncm_aux = Trim(rs("ncm"))
				cst_aux = Trim(rs("cst"))
                st_ncm_cst_herdado_tabela_produto_aux = rs("st_ncm_cst_herdado_tabela_produto")
				ean_aux = Trim(rs("ean"))
                aliq_ipi_aux = rs("aliq_ipi")
                aliq_icms_aux = rs("aliq_icms")
                vl_ipi_aux = rs("vl_ipi")
                preco_origem_aux = rs("preco_origem")
                produto_xml_aux = Trim(rs("produto_xml"))
                descricao_html_aux = Trim(rs("descricao_html"))
				End If
			
			If (qtde_a_sair - qtde_movimentada) > (qtde_aux - qtde_utilizada_aux) Then
			'	QUANTIDADE DE PRODUTOS DESTE ITEM DE ESTOQUE É INSUFICIENTE P/ ATENDER O PEDIDO
				qtde_movto = qtde_aux - qtde_utilizada_aux
			Else
			'	QUANTIDADE DE PRODUTOS DESTE ITEM SOZINHO É SUFICIENTE P/ ATENDER O PEDIDO
				qtde_movto = qtde_a_sair - qtde_movimentada
				End If

			If v_item(Ubound(v_item)).produto <> "" Then
				redim preserve v_item(ubound(v_item)+1)
				set v_item(ubound(v_item)) = New cl_ESTOQUE_TRANSFERENCIA_ITEM
				End if
			v_item(Ubound(v_item)).documento = Trim(v_documento(iv))
            v_item(Ubound(v_item)).entrada_tipo = v_entrada_tipo(iv)
			v_item(Ubound(v_item)).id_estoque_origem = Trim(v_estoque(iv))
			v_item(Ubound(v_item)).fabricante = id_fabricante
			v_item(Ubound(v_item)).produto = id_produto
            v_item(Ubound(v_item)).descricao_html = descricao_html_aux
			v_item(Ubound(v_item)).qtde = qtde_movto
            v_item(Ubound(v_item)).preco_fabricante = preco_fabricante_aux
			v_item(Ubound(v_item)).vl_custo2 = vl_custo2_aux
            v_item(Ubound(v_item)).vl_BC_ICMS_ST = vl_BC_ICMS_ST_aux
            v_item(Ubound(v_item)).vl_ICMS_ST = vl_ICMS_ST_aux
			v_item(Ubound(v_item)).ncm = ncm_aux
			v_item(Ubound(v_item)).cst = cst_aux
			v_item(Ubound(v_item)).ean = ean_aux
            v_item(Ubound(v_item)).st_ncm_cst_herdado_tabela_produto = st_ncm_cst_herdado_tabela_produto_aux
            v_item(Ubound(v_item)).aliq_ipi = aliq_ipi_aux
            v_item(Ubound(v_item)).aliq_icms = aliq_icms_aux
            v_item(Ubound(v_item)).vl_ipi = vl_ipi_aux
            v_item(Ubound(v_item)).preco_origem = preco_origem_aux
            v_item(Ubound(v_item)).produto_xml = produto_xml_aux
		
		'	CONTABILIZA QUANTIDADE MOVIMENTADA
			qtde_movimentada = qtde_movimentada + qtde_movto
		
		'	JÁ CONSEGUIU ALOCAR TUDO?
			If qtde_movimentada >= qtde_a_sair Then Exit For
			End If
		Next
	
	if rs.State <> 0 then rs.Close
	set rs=nothing
	
	
	estoque_produto_transf_cd_monta = True
end function


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