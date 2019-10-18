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
    If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	
	c_log_edicao = Trim(Request.Form("c_log_edicao"))
    
   	c_nfe_emitente_origem = Trim(Request.Form("c_nfe_emitente_origem"))
	c_nfe_emitente_destino = Trim(Request.Form("c_nfe_emitente_destino"))
    c_documento_transf = Trim(Request.Form("c_documento_transf"))
    c_obs = Trim(Request.Form("c_obs"))

    dim v_item1, v_item2, v_item3
    dim v_agrupa
    dim s_cod_prod1, s_cod_prod2
    dim id_estoque_transferencia
    dim id_estoque_transferencia_item

    'Procedimento: 
    '- obter as informações da tela anterior e armazenar no vetor 1
    '- chamar novamente a rotina de montagem e armazenar no vetor 2
    '- fazer a comparação entre os vetores; só gravar se os valores baterem
    '(OBS - COLETAR NA TABELA DA TELA ANTERIOR O ID_ESTOQUE_ORIGEM)

    n = Request.Form("c_produto").Count

   	redim v_item1(0)
	set v_item1(0) = New cl_ESTOQUE_TRANSFERENCIA_ITEM_SUB
    for i = 1 to n 
        if Trim(Request.Form("c_produto")(i)) <> "" then 
            if Trim(v_item1(ubound(v_item1)).produto) <> "" then
				redim preserve v_item1(ubound(v_item1)+1)
				set v_item1(ubound(v_item1)) = New cl_ESTOQUE_TRANSFERENCIA_ITEM_SUB
				end if
			with v_item1(ubound(v_item1))
                .documento = Trim(Request.Form("c_documento")(i))
                .id_estoque_origem = Trim(Request.Form("c_id_estoque_origem")(i))
                .fabricante = Trim(Request.Form("c_fabricante")(i))
                .produto = Trim(Request.Form("c_produto")(i))
                .qtde  = CInt(Trim(Request.Form("c_qtde")(i)))
                .vl_custo2 = Trim(Request.Form("c_vl_custo2")(i))                
                .aliq_ipi = Trim(Request.Form("c_aliq_ipi")(i))                
                .aliq_icms = Trim(Request.Form("c_aliq_icms")(i))                
                .vl_ipi = Trim(Request.Form("c_vl_ipi")(i))                
                .nfe_entrada_numero = Trim(Request.Form("c_nfe_entrada_numero")(i))                
                .nfe_entrada_serie = Trim(Request.Form("c_nfe_entrada_serie")(i))                
                end with
            end if
        next

   	redim v_item2(0)
	set v_item2(0) = New cl_ESTOQUE_TRANSFERENCIA_ITEM_SUB
    
    if not estoque_produto_transf_consiste_quantidades(c_nfe_emitente_origem, _
                                                    c_nfe_emitente_destino, _
                                                    v_item1, _
                                                    v_item2, _
										            msg_erro) then
        alerta = msg_erro
        end if

	if alerta = "" then

    '   PREENCHENDO VALORES EM TELA PARA VETOR 2
	    for i = Lbound(v_item2) to Ubound(v_item2)
		    with v_item2(i)
                .aliq_ipi = v_item1(i).aliq_ipi
                .aliq_icms = v_item1(i).aliq_icms
                .vl_ipi = v_item1(i).vl_ipi
                .nfe_entrada_numero = v_item1(i).nfe_entrada_numero
                .nfe_entrada_serie = v_item1(i).nfe_entrada_serie
                end with
            next

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
	    set v_item3(0) = New cl_ESTOQUE_TRANSFERENCIA_ITEM_SUB
    
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

    'REAGRUPAR
        redim v_agrupa(0)
	    set v_agrupa(0) = New cl_ESTOQUE_TRANSFERENCIA_ITEM
        v_agrupa(0).produto = ""
    
        if Not estoque_produto_transf_reagrupa_itens(v_item2, _
                                                    v_agrupa, _
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
					" (id, id_nfe_emitente_origem, id_nfe_emitente_destino, documento, data, data_hora, usuario, obs" & _   
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
                for i=lbound(v_agrupa) to ubound(v_agrupa)
                    with v_agrupa(i)
			            s_sql = " INSERT INTO T_ESTOQUE_TRANSFERENCIA_ITEM " & _
					            " (id_estoque_transferencia, fabricante, produto, qtde,  " & _
                                " aliq_ipi, aliq_icms, vl_ipi " & _
                                "  ) VALUES " & _
                                " ("  & _
	                            CStr(id_estoque_transferencia) & ", " & _   
                                " '" & .fabricante & "', " & _
                                " '" & .produto  & "', " & _
                                .qtde  & ", " & _
                                Iif(Trim(.aliq_ipi) = "", "NULL", bd_formata_numero(.aliq_ipi)) & ", " & _
                                Iif(Trim(.aliq_icms) = "", "NULL", bd_formata_numero(.aliq_icms)) & ", " & _
                                Iif(Trim(.vl_ipi) = "", "NULL", bd_formata_numero(.vl_ipi)) & " " & _
                                " )" 
			            cn.Execute(s_sql)
                        end with
			        if Err <> 0 then
                        msg_erro= "Problema na inclusão dos itens agrupados" & vbCrLf
				        msg_erro= msg_erro & Cstr(Err) & ": " & Err.Description
				        end if				
                    next
                end if
                
            if msg_erro = "" then
                for i=lbound(v_item2) to ubound(v_item2)
                    with v_item2(i)
                        'OBTEM O id DA TABELA t_ESTOQUE_TRANSFERENCIA_ITEM para cada produto
                        s_sql = " SELECT id FROM t_ESTOQUE_TRANSFERENCIA_ITEM " & _
                              " WHERE id_estoque_transferencia = " & CStr(id_estoque_transferencia)  & _   
                              " AND fabricante = '" & .fabricante & "' " & _
                              " AND produto = '" & .produto & "' "
                        if rs.State <> 0 then rs.Close
                        rs.open s_sql, cn
                        id_estoque_transferencia_item = rs("id")

			            s_sql = " INSERT INTO T_ESTOQUE_TRANSFERENCIA_ITEM_SUB " & _
					            " (id_estoque_transferencia, id_estoque_transferencia_item, id_estoque_origem, entrada_tipo, documento, fabricante, produto, qtde, preco_fabricante, vl_custo2, vl_BC_ICMS_ST, vl_ICMS_ST,  " & _
                                " ncm, cst, st_ncm_cst_herdado_tabela_produto, ean, aliq_ipi, aliq_icms, vl_ipi, preco_origem, produto_xml, nfe_entrada_numero, nfe_entrada_serie " & _
                                "  ) VALUES " & _
                                " ("  & _
	                            CStr(id_estoque_transferencia) & ", " & _   
	                            CStr(id_estoque_transferencia_item) & ", " & _   
    				            " '" & .id_estoque_origem & "', " & _
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
                                Iif(Trim(.aliq_ipi) = "", "NULL", bd_formata_numero(.aliq_ipi)) & ", " & _
                                Iif(Trim(.aliq_icms) = "", "NULL", bd_formata_numero(.aliq_icms)) & ", " & _
                                Iif(Trim(.vl_ipi) = "", "NULL", bd_formata_numero(.vl_ipi)) & ", " & _
                                " '" & .preco_origem & "', " & _
                                " '" & .produto_xml & "', " & _
                                " '" & .nfe_entrada_numero & "', " & _
                                " '" & .nfe_entrada_serie & "' " & _
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