<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/estoque.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =========================================================
'	  E S T O Q U E E N T R A D A C O N F I R M A E A N . A S P
'     =========================================================
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

	class cl_ITEM_ESTOQUE_CADASTRAMENTO_AMPLIADO
		dim id_estoque
		dim fabricante
		dim produto
		dim qtde
		dim qtde_utilizada
		dim preco_fabricante
		dim data_ult_movimento
		dim sequencia
		dim vl_custo2
		dim vl_BC_ICMS_ST
		dim vl_ICMS_ST
		dim ncm
		dim ncm_redigite
		dim cst
		dim cst_redigite
        dim aliq_ipi
        dim aliq_icms
        dim vl_ipi
		end class
    
    class cl_ITEM_ESTOQUE_EAN
		dim fabricante
		dim produto
        dim ean
        dim ean_original
        end class
	
    dim s, s_log, i, n, usuario, msg_erro, c_log_edicao
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 
	
	dim alerta
	alerta=""

	dim r_estoque, v_item, v_item_ean
    dim c_perc_agio
    dim s_sql    

'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	
	c_log_edicao = Trim(Request.Form("c_log_edicao"))
	c_perc_agio = Trim(Request.Form("c_perc_agio"))												
	
	set r_estoque = New cl_ESTOQUE_AGIO
	with r_estoque
		.data_entrada = Date
		.hora_entrada = retorna_so_digitos(formata_hora(Now))
		.data_ult_movimento = Date
		.fabricante = normaliza_codigo(retorna_so_digitos(Request.Form("c_fabricante")), TAM_MIN_FABRICANTE)
		.documento = Trim(Request.Form("c_documento"))
        .perc_agio = converte_numero(c_perc_agio)
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
    redim v_item_ean(0)
	set v_item(0) = New cl_ITEM_ESTOQUE_CADASTRAMENTO_AMPLIADO
    set v_item_ean(0) = New cl_ITEM_ESTOQUE_EAN
	n = Request.Form("c_produto").Count
	for i = 1 to n
		s=Trim(Request.Form("c_produto")(i))
		if s <> "" then
			if Trim(v_item(ubound(v_item)).produto) <> "" then
				redim preserve v_item(ubound(v_item)+1)
				set v_item(ubound(v_item)) = New cl_ITEM_ESTOQUE_CADASTRAMENTO_AMPLIADO
				end if
			if Trim(v_item_ean(ubound(v_item_ean)).produto) <> "" then
                redim preserve v_item_ean(ubound(v_item_ean)+1)
               set v_item_ean(ubound(v_item_ean)) = New cl_ITEM_ESTOQUE_EAN
			 	end if
			with v_item(ubound(v_item))
				.fabricante=r_estoque.fabricante
				.produto=Ucase(Trim(Request.Form("c_produto")(i)))
				s = Trim(Request.Form("c_qtde")(i))
				if IsNumeric(s) then .qtde = CLng(s) else .qtde = 0
			'	PREÇO FABRICANTE
				s = Trim(Request.Form("c_vl_unitario")(i))
				.preco_fabricante = converte_numero(s)
				if .preco_fabricante < 0 then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Produto " & .produto & " está com valor inválido: " & formata_moeda(.preco_fabricante)
					end if
			'	CUSTO 2
				s = Trim(Request.Form("c_vl_custo2")(i))
				.vl_custo2 = converte_numero(s)
				if .vl_custo2 < 0 then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Produto " & .produto & " está com Custo II inválido: " & formata_moeda(.vl_custo2)
					end if
			'	BASE CÁLCULO ICMS ST
				s = Trim(Request.Form("c_vl_BC_ICMS_ST")(i))
				.vl_BC_ICMS_ST = converte_numero(s)
				if .vl_BC_ICMS_ST < 0 then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Produto " & .produto & " está com valor de base de cálculo do ICMS ST inválido: " & formata_moeda(.vl_BC_ICMS_ST)
					end if
			'	VALOR DO ICMS ST
				s = Trim(Request.Form("c_vl_ICMS_ST")(i))
				.vl_ICMS_ST = converte_numero(s)
				if .vl_ICMS_ST < 0 then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Produto " & .produto & " está com valor do ICMS ST inválido: " & formata_moeda(.vl_ICMS_ST)
					end if
			'	NCM
				.ncm = Trim(Request.Form("c_ncm")(i))
				.ncm_redigite = Trim(Request.Form("c_ncm_redigite")(i))
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
				.cst = Trim(Request.Form("c_cst")(i))
				.cst_redigite = Trim(Request.Form("c_cst_redigite")(i))
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
			'	ALIQUOTA DO IPI
				s = Trim(Request.Form("c_aliq_ipi")(i))
				.aliq_ipi = converte_numero(s)
				if .aliq_ipi < 0 then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Produto " & .produto & " está com a alíquota do IPI inválida: " & formata_numero(.aliq_ipi)
					end if
			'	ALIQUOTA DO ICMS
				s = Trim(Request.Form("c_aliq_icms")(i))
				.aliq_icms = converte_numero(s)
				if .aliq_icms < 0 then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Produto " & .produto & " está com a alíquota do ICMS inválida: " & formata_numero(.aliq_icms)
					end if
			'	VALOR DO IPI
				s = Trim(Request.Form("c_vl_ipi")(i))
				.vl_ipi = converte_numero(s)
				if .vl_ipi < 0 then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Produto " & .produto & " está com o valor do IPI inválido: " & formata_numero(.vl_ipi)
					end if
				end with
        '   INFORMAÇÕES EAN
            with v_item_ean(ubound(v_item_ean))
				.fabricante=r_estoque.fabricante
				.produto=Ucase(Trim(Request.Form("c_produto")(i)))
            '   EAN e EAN ORIGINAL
				.ean = Trim(Request.Form("c_ean")(i))
                .ean_original = Trim(Request.Form("c_ean_original")(i))
                end with
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

		s = "Entrada no estoque de mercadorias do fabricante=" & Trim(r_estoque.fabricante) & "," & _
			" documento=" & Trim(r_estoque.documento)
		if r_estoque.entrada_especial <> 0 then s = s & ", registrado como entrada especial"
		s_log = s & ":" & s_log
		
	'	~~~~~~~~~~~~~
		cn.BeginTrans
	'	~~~~~~~~~~~~~
        if Not estoque_nova_entrada_mercadorias_agio(r_estoque, v_item, msg_erro) then 
		'	~~~~~~~~~~~~~~~~
			cn.RollbackTrans
		'	~~~~~~~~~~~~~~~~
			Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_MOVIMENTO_ESTOQUE)
			end if

    ' 	GRAVA O EAN NA TABELA DE PRODUTOS
        msg_erro = ""
	    For i = LBound(v_item_ean) To UBound(v_item_ean)
		    With v_item_ean(i)
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
			Response.Redirect("estoqueconsultaean.asp?estoque_selecionado=" & r_estoque.id_estoque  & "&url_back=X" & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")))
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
	<td align="center"><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()"><img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
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