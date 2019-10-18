<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/estoque.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'      ==============================================================================
'       E S T O Q U E T R A N S F E R E E N T R E C D S G R A V A D A D O S . A S P
'      ==============================================================================
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
    dim c_transf_selecionada, c_nfe_emitente_origem, c_nfe_emitente_destino, c_documento_transf, c_obs

	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

    c_log_edicao = Trim(Request.Form("c_log_edicao"))
    
    c_transf_selecionada = Trim(Request("transf_selecionada"))
   	c_nfe_emitente_origem = Trim(Request.Form("c_nfe_emitente_origem"))
	c_nfe_emitente_destino = Trim(Request.Form("c_nfe_emitente_destino"))
    c_documento_transf = Trim(Request.Form("c_documento_transf"))
    c_obs = Trim(Request.Form("c_obs"))

    dim v_item1, v_item2, v_item3
    dim s_cod_prod1, s_cod_prod2
    dim id_estoque_transferencia
    dim s_chave
    dim v_lista_id_estoque
    dim s_id_estoque_destino
    dim s_id_fabricante_destino
    dim iv, j, i_seq

    'Procedimento: 
    '- obter as informações da tela anterior e armazenar no vetor 1
    '- chamar novamente a rotina de montagem e armazenar no vetor 2
    '- fazer a comparação entre os vetores; só efetuar a movimentação de estoque se os valores baterem
 
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
                .ean = Trim(Request.Form("c_ean")(i))
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
    '   COMPLEMENTA NO VETOR 1 OS CAMPOS DO BD PRESENTES NO VETOR 2
        For i = LBound(v_item2) to UBound(v_item2)
            with v_item1(i)
                .entrada_tipo = v_item2(i).entrada_tipo
                .vl_BC_ICMS_ST = v_item2(i).vl_BC_ICMS_ST
                .vl_ICMS_ST = v_item2(i).vl_ICMS_ST
                .ncm = v_item2(i).ncm
                .cst = v_item2(i).cst
                .st_ncm_cst_herdado_tabela_produto = v_item2(i).st_ncm_cst_herdado_tabela_produto
                .preco_origem = v_item2(i).preco_origem
                .produto_xml = v_item2(i).produto_xml
                end with
            Next

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
        
        msg_erro = ""        

        if Not cria_recordset_pessimista(rs, msg_erro) then
		'	~~~~~~~~~~~~~~~~
			cn.RollbackTrans
		'	~~~~~~~~~~~~~~~~
			Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
			end if

        redim v_item3(0)
	    set v_item3(0) = New cl_ESTOQUE_TRANSFERENCIA_ITEM_SUB
    
        if Not estoque_produto_transf_consiste_quantidades(c_nfe_emitente_origem, _
                                                    c_nfe_emitente_destino, _
                                                    v_item1, _
                                                    v_item3, _
										            msg_erro) then
		    msg_erro = "Falha na opperação de transferência: " & mmsg_erro
            end if

'---------------- INCEPTION INICIO -------------------------


    '	REALIZA A SAÍDA DO ESTOQUE!!
   	    redim v_lista_id_estoque(0)
	    set v_lista_id_estoque(0) = New cl_QUATRO_COLUNAS
        v_lista_id_estoque(0).c1 = "" 
        s_id_estoque_destino  = ""
        s_id_fabricante_destino  = ""
        i_seq = 0
	    For iv = LBound(v_item1) To UBound(v_item1)
            With v_item1(iv)
    
                'response.Write " lhgx item " & cstr(iv)
                
                '   OS PRODUTOS QUE TIVEREM O MESMO ID DO ESTOQUE DE ORIGEM TERÃO O MESMO ID NO ESTOQUE DE DESTINO
                '   SE É A PRIMEIRA APARIÇÃO DO ESTOQUE DE ORIGEM, CRIAR UM NOVO ESTOQUE DE DESTINO
                '   (o vetor v_lista_id_estoque armazenará:
                '   - o id origem na primeira coluna
                '   - o id destino na segunda coluna
                '   - o fabricante na terceira coluna
                '   - a sequência a ser usada no estoque na quarta coluna)
		            If localiza_cl_quatro_colunas(v_lista_id_estoque, .id_estoque_origem, j) Then
                        s_id_estoque_destino = v_lista_id_estoque(j).c2
                        s_id_fabricante_destino = v_lista_id_estoque(j).c3
                        v_lista_id_estoque(j).c4 = v_lista_id_estoque(j).c4 + 1
                        i_seq = v_lista_id_estoque(j).c4
                    Else
                    '	GERA A CHAVE P/ A NOVA ENTRADA NO ESTOQUE
				        If Not gera_id_estoque(s_id_estoque_destino, msg_erro) Then
				            msg_erro = "Falha na criação da identificação do estoque de destino"
                            end if

                        s_id_fabricante_destino = .fabricante
                    
                        s = "INSERT INTO t_ESTOQUE (" & _
						        "id_estoque, data_entrada, hora_entrada, id_nfe_emitente, fabricante, documento," & _
						        " usuario, data_ult_movimento, kit" & _
					        ") VALUES (" & _
						        "'" & s_id_estoque_destino & "'" & _
						        "," & bd_formata_data(Date) & _
						        ",'" & retorna_so_digitos(formata_hora(Now)) & "'" & _
						        ", " & c_nfe_emitente_destino & _
						        ",'" & s_id_fabricante_destino & "'" & _
						        ",'" & c_documento_transf & "'" & _
						        ",'" & usuario & "'" & _
						        "," & bd_formata_data(Date) & _
						        ", " & "0" & _
					        ")"
				        cn.Execute(s)
				        if Err <> 0 then
                            msg_erro = "Falha no cadastramento do estoque de destino - " & Cstr(Err) & ": " & Err.Description
					        end if

                        If v_lista_id_estoque(ubound(v_lista_id_estoque)).c1 <> "" Then
   	                        redim preserve v_lista_id_estoque(ubound(v_lista_id_estoque)+1) 
    	                    set v_lista_id_estoque(ubound(v_lista_id_estoque)) = New cl_QUATRO_COLUNAS
                            End If
                        v_lista_id_estoque(ubound(v_lista_id_estoque)).c1 = .id_estoque_origem
                        v_lista_id_estoque(ubound(v_lista_id_estoque)).c2 = s_id_estoque_destino
                        v_lista_id_estoque(ubound(v_lista_id_estoque)).c3 = s_id_fabricante_destino
                        v_lista_id_estoque(ubound(v_lista_id_estoque)).c4 = 1
                        i_seq = 1
                        ordena_cl_quatro_colunas v_lista_id_estoque, 1, UBound(v_lista_id_estoque)

                        End If

                '   T_ESTOQUE_ITEM: ENTRADA DE PRODUTOS NO ESTOQUE DESTINO
					s = "INSERT INTO T_ESTOQUE_ITEM" & _
						" (id_estoque, fabricante, produto, qtde, preco_fabricante, vl_custo2," & _
						" vl_BC_ICMS_ST, vl_ICMS_ST," & _
						" ncm, cst," & _
						" data_ult_movimento, sequencia, " & _
                        " ean, aliq_ipi, vl_ipi, aliq_icms, preco_origem, produto_xml " & _
						") VALUES (" & _
						"'" & s_id_estoque_destino & "'" & _
						",'" & .fabricante & "'" & _
						",'" & .produto & "'" & _
						"," & CStr(.qtde) & _
						"," & Iif(Trim(.preco_fabricante)="", "0.00", bd_formata_numero(.preco_fabricante)) & _
						"," & Iif(Trim(.vl_custo2)="", "0.00", bd_formata_numero(.vl_custo2)) & _
						"," & Iif(Trim(.vl_BC_ICMS_ST)="", "0.00", bd_formata_numero(.vl_BC_ICMS_ST)) & _
						"," & Iif(Trim(.vl_ICMS_ST)="", "0.00", bd_formata_numero(.vl_ICMS_ST)) & _
						",'" & Trim(.ncm) & "'" & _
						",'" & Trim(.cst) & "'" & _
						"," & bd_formata_data(Date) & _
						"," & CStr(i_seq) & _
						",'" & Trim(.ean) & "'" & _
						"," & Iif(Trim(.aliq_ipi)="", "NULL", bd_formata_numero(.aliq_ipi)) & _
						"," & Iif(Trim(.vl_ipi)="", "NULL", bd_formata_numero(.vl_ipi)) & _
						"," & Iif(Trim(.aliq_icms)="", "NULL", bd_formata_numero(.aliq_icms)) & _
						",'" & .preco_origem & "'" & _
						",'" & .produto_xml & "'" & _
						")"
                    'response.Write s
					cn.Execute(s)
					if Err <> 0 then
						msg_erro="Falha no cadastramento do item do estoque de destino - " & Cstr(Err) & ": " & Err.Description
						end if

                    'response.Write " lhgx atualizaremos estoque item " & cstr(iv)

		        '	T_ESTOQUE_ITEM: SAÍDA DE PRODUTOS DO ESTOQUE ORIGEM
                '   (obs: o fabricante de destino é o mesmo de origem)
			        s = "SELECT " & _
					        "*" & _
				        " FROM t_ESTOQUE_ITEM" & _
				        " WHERE" & _
					        " (id_estoque = '" & .id_estoque_origem & "')" & _
					        " AND (fabricante = '" & .fabricante & "')" & _
					        " AND (produto = '" & .produto & "')"
			        if rs.State <> 0 then rs.Close
			        rs.open s, cn
			        if Err <> 0 then
				        msg_erro = "Falha na seleção do item do estoque de origem - " & Cstr(Err) & ": " & Err.Description
				        end if

			        if rs.Eof then
				        msg_erro = "Falha ao acessar o registro no estoque do produto " & .produto & " do fabricante " & .fabricante & " (id_estoque = '" & id_estoque_origem & "')"
				        end if

			        rs("qtde_utilizada") = rs("qtde_utilizada") + .qtde
			        rs("data_ult_movimento") = Date
			        rs.Update
			        if Err <> 0 then
				        msg_erro="Falha na atualização do estoque de origem - " & Cstr(Err) & ": " & Err.Description
				        end if
		
		        '	T_ESTOQUE_MOVIMENTO: REGISTRA O MOVIMENTO DE SAÍDA DO ESTOQUE
			        If Not gera_id_estoque_movto(s_chave, msg_erro) Then
				        msg_erro = "Falha ao tentar obter um nº de identificação único para este registro de movimentação no estoque!!" & _
							        chr(13) & msg_erro
				        End If
			
                    'response.Write " lhgx inseriremos estoque item " & cstr(iv)
                    
                    s = "INSERT INTO t_ESTOQUE_MOVIMENTO" & _
				        " (id_movimento, data, hora, operacao, estoque, usuario, pedido, loja," & _
				        " fabricante, produto, id_estoque, qtde, kit) VALUES" & _
				        " ('" & s_chave & "'" & _
				        "," & bd_formata_data(Date) & _
				        ",'" & retorna_so_digitos(formata_hora(Now)) & "'" & _
				        ",'" & OP_ESTOQUE_TRANSFERENCIA & "'" & _
				        ",'" & ID_ESTOQUE_VENDA & "'" & _
				        ",'" & usuario & "'" & _
				        ",'" & "" & "'" & _
				        ",'" & "" & "'" & _
				        ",'" & .fabricante & "'" & _
				        ",'" & .produto & "'" & _
				        ",'" & s_id_estoque_destino & "'" & _
				        "," & CStr(.qtde) & _
				        "," & "0"  & ")"
			        cn.Execute(s)
			        if Err <> 0 then
				        msg_erro="Falha no cadastramento da movimentação do estoque de destino - " & Cstr(Err) & ": " & Err.Description
				        end if

                    'response.Write " lhgx atualizaremos estoque " & cstr(iv)

		        '	T_ESTOQUE: ATUALIZA DATA DO ÚLTIMO MOVIMENTO
			        s = "SELECT " & _
					        "*" & _
				        " FROM t_ESTOQUE" & _
				        " WHERE" & _
					        " (id_estoque = '" & .id_estoque_origem & "')"
			        if rs.State <> 0 then rs.Close
			        rs.open s, cn
			        if Err <> 0 then
				        msg_erro="Falha no seleção do estoque de origem - " & Cstr(Err) & ": " & Err.Description
				        end if
			
			        if rs.Eof then
				        msg_erro = "Falha ao acessar o registro principal no estoque do produto " & id_produto & " do fabricante " & id_fabricante
			        else
				        rs("data_ult_movimento") = Date
				        rs.Update
				        if Err <> 0 then
					        msg_erro="Falha na atualização do estoque de origem - " & Cstr(Err) & ": " & Err.Description
					        end if
				        End If

            ' 	ATUALIZA t_ESTOQUE_TRANSFERENCIA INDICANDO QUE A TRANSFERENCIA FOI CONFIRMADA

                'response.Write " lhgx atualizaremos transferência " & cstr(iv)

		        s_sql = " UPDATE T_ESTOQUE_TRANSFERENCIA_ITEM_SUB SET" & _
                        " id_estoque_destino = '" & s_id_estoque_destino & "'" & _
                        " WHERE (id_estoque_transferencia = '" & c_transf_selecionada & "') "
		        cn.Execute(s_sql)
		        if Err <> 0 then
                    msg_erro= "Problema na atualização da transferência" & vbCrLf
			        msg_erro= msg_erro & Cstr(Err) & ": " & Err.Description
			        end if				
                
                'response.Write " lhgx gravaremos log " & cstr(iv)

            '   Log de movimentação do estoque
	            if Not grava_log_estoque_v2(usuario, c_nfe_emitente_origem, .fabricante, .produto, .qtde, .qtde, OP_ESTOQUE_TRANSFERENCIA, _
                                            ID_ESTOQUE_VENDA, ID_ESTOQUE_VENDA, "", "", "", "", c_documento_transf, _
                                            "Transferência do estoque " & .id_estoque_origem & " (CD " & c_nfe_emitente_origem & _
                                            ") para o estoque " & s_id_estoque_destino & " (CD " & c_nfe_emitente_destino & ")", "") then
		            msg_erro="FALHA AO GRAVAR O LOG DA MOVIMENTAÇÃO NO ESTOQUE"
		            end if

                'response.Write " lhgx mensagem de erro " & msg_erro

                if msg_erro <> "" then
                '	~~~~~~~~~~~~~~~~
			        cn.RollbackTrans
		        '	~~~~~~~~~~~~~~~~
			        Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_TRANSFERENCIA_CD_CONFERE)
                    end if
                    
                End With

	        Next        
	
	    if rs.State <> 0 then rs.Close
	    set rs=nothing
	
'---------------- INCEPTION FIM -------------------------

    ' 	ATUALIZA t_ESTOQUE_TRANSFERENCIA INDICANDO QUE A TRANSFERENCIA FOI CONFIRMADA

		s_sql = " UPDATE T_ESTOQUE_TRANSFERENCIA SET" & _
                " st_confirmada = " & "1"  & ", " & _
                " dt_confirma = " & bd_formata_data(Date) & ", " & _
                " dt_hr_confirma = " & bd_formata_data_hora(Date) & ", " & _
                " usuario_confirma = '" & usuario & "' " & _
                " WHERE (id = " & c_transf_selecionada & ") "
		cn.Execute(s_sql)
		if Err <> 0 then
            msg_erro= "Problema na atualização da transferência" & vbCrLf
			msg_erro= msg_erro & Cstr(Err) & ": " & Err.Description
			end if				


        if msg_erro <> "" then
		 '	~~~~~~~~~~~~~~~~
			 cn.RollbackTrans
		 '	~~~~~~~~~~~~~~~~
			 alerta = msg_erro
        else        		
	    '	~~~~~~~~~~~~~~
		    cn.CommitTrans
	    '	~~~~~~~~~~~~~~
		    if Err<>0 then
		        alerta=Cstr(Err) & ": " & Err.Description
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



<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>


<script language="JavaScript" type="text/javascript">

    function fESTOQRetorna() {
        f.action = "estoquetransfereentrecds.asp";
        dREMOVE.style.visibility = "hidden";
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
	<td align="center"><a name="bVOLTAR" id="bVOLTAR" href="estoquetransfereentrecds.asp"><img src="..\botao\voltar.gif" width="176" height="55" border="0"></a></td>
</tr>
</table>
</center>
</body>
<% else %>
<!-- ************************************************************ -->
<!-- **************  PÁGINA PARA EXIBIR RESULTADO  ************** -->
<!-- ************************************************************ -->
<body onload="bVOLTAR.focus();">
<center>
<br>
<!--  T E L A  -->
<p class="T">A V I S O</p>
<div class='MtAviso' style="width:649px;font-weight:bold;border:1pt solid black;" align="center"><p style='margin:5px 2px 5px 2px;'>Transferência '<%=c_transf_selecionada%>' realizada com sucesso!!!</p></div>
<br><br>
<p class="TracoBottom"></p>
<table cellSpacing="0">
<tr>
	<td align="center"><a name="bVOLTAR" id="bVOLTAR" href="estoquetransfereentrecdsfiltro.asp"><img src="..\botao\voltar.gif" width="176" height="55" border="0"></a></td>
</tr>
</table>
</center>
</body>

<% end if %>

</html>


<%
'	if rs.State <> 0 then rs.Close
'	set rs = nothing

'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>