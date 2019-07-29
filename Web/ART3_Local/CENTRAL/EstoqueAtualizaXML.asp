<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/estoque.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ===================================================
'	  E S T O Q U E A T U A L I Z A X M L . A S P
'     ===================================================
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

	dim s, s_log, usuario, msg_erro
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

	dim estoque_selecionado
	estoque_selecionado = Trim(request("estoque_selecionado"))
	if (estoque_selecionado = "") then Response.Redirect("aviso.asp?id=" & ERR_ESTOQUE_NAO_ESPECIFICADO)

    dim c_perc_agio 
	
'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs, tEI
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(tEI, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	if Not operacao_permitida(OP_CEN_EDITA_ENTRADA_ESTOQUE, s_lista_operacoes_permitidas) then
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	dim alerta, mensagem
	alerta = ""
	mensagem = ""
	
	dim i, n
	dim r_estoque, v_item
	if Not le_estoque_agio(estoque_selecionado, r_estoque, msg_erro) then
		alerta = "Falha ao tentar consultar o registro de entrada de mercadorias no estoque nº " & _
				 estoque_selecionado & ": " & msg_erro
		end if

	if alerta = "" then
		r_estoque.documento = Trim(Request.Form("c_documento"))
		s = Trim(Request.Form("ckb_especial"))
		if s <> "" then
			r_estoque.entrada_especial = 1
		else
			r_estoque.entrada_especial = 0
			end if
		
		r_estoque.obs = Trim(Request.Form("c_obs"))
        'c_perc_agio = Replace(Trim(Request.Form("c_perc_agio")), ",", ".")
        c_perc_agio = Trim(Request.Form("c_perc_agio"))
        r_estoque.perc_agio = converte_numero(c_perc_agio)
		
		redim v_item(0)
		set v_item(0) = New cl_ITEM_ESTOQUE_ENTRADA_XML
		n = Request.Form("c_produto").Count
		for i = 1 to n
			s=Trim(Request.Form("c_produto")(i))
			if s <> "" then
				if Trim(v_item(ubound(v_item)).produto) <> "" then
					redim preserve v_item(ubound(v_item)+1)
					set v_item(ubound(v_item)) = New cl_ITEM_ESTOQUE_ENTRADA_XML
					end if
				with v_item(ubound(v_item))
					.fabricante = r_estoque.fabricante
					.produto = Ucase(Trim(Request.Form("c_produto")(i)))
					s = Trim(Request.Form("c_qtde")(i))
					if IsNumeric(s) then .qtde = CLng(s) else .qtde = 0
					.preco_fabricante = converte_numero(Trim(Request.Form("c_vl_unitario")(i)))
					.vl_custo2 = converte_numero(Trim(Request.Form("c_vl_custo2")(i)))
					.ncm = Trim(Request.Form("c_ncm")(i))
					.cst = Trim(Request.Form("c_cst")(i))
					.ean = Trim(Request.Form("c_ean")(i))
					.aliq_ipi = Trim(Request.Form("c_aliq_ipi")(i))
					.vl_ipi = Trim(Request.Form("c_vl_ipi")(i))
					.aliq_icms = Trim(Request.Form("c_aliq_icms")(i))
					end with
				end if
			next
		end if

	dim blnValorEditavel, sAnoMesEstoque, sAnoMesHoje, s_valor_readonly
	blnValorEditavel = False
	sAnoMesEstoque = Left(formata_data_yyyymmdd(r_estoque.data_entrada), 6)
	sAnoMesHoje = Left(formata_data_yyyymmdd(Now), 6)
	if sAnoMesEstoque = sAnoMesHoje then blnValorEditavel = True

	if alerta = "" then
		'Se o valor não for editável, assegura que continuarão idênticos aos valores já cadastrados
		if Not blnValorEditavel then
			for i=LBound(v_item) to UBound(v_item)
				if Trim("" & v_item(i).produto) <> "" then
					s = "SELECT " & _
							"*" & _
						" FROM t_ESTOQUE_ITEM" & _
						" WHERE" & _
							" (id_estoque = '" & estoque_selecionado & "')" & _
							" AND (fabricante = '" & Trim(v_item(i).fabricante) & "')" & _
							" AND (produto = '" & Trim(v_item(i).produto) & "')"
					if tEI.State <> 0 then tEI.Close
					tEI.open s, cn
					if Not tEI.Eof then
						v_item(i).preco_fabricante = tEI("preco_fabricante")
						v_item(i).vl_custo2 = tEI("vl_custo2")
						end if
					end if
				next
			end if
		end if

	if alerta = "" then
	'	~~~~~~~~~~~~~
		cn.BeginTrans
	'	~~~~~~~~~~~~~
		if estoque_atualiza_xml(usuario, r_estoque, v_item, s_log, msg_erro) then
			if s_log <> "" then grava_log usuario, "", "", "", OP_LOG_ESTOQUE_ALTERACAO, s_log

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
				s = "SELECT id_estoque FROM t_ESTOQUE WHERE (id_estoque='" & r_estoque.id_estoque & "')"
				set rs = cn.execute(s)
				if Not rs.Eof then
					Response.Redirect("estoqueconsultaxml.asp?estoque_selecionado=" & r_estoque.id_estoque & "&url_back=X" & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")))
				else
					mensagem = "Lote de mercadorias foi excluído do estoque."
					end if
			else
				alerta=Cstr(Err) & ": " & Err.Description
				end if
		else
		'	~~~~~~~~~~~~~~~~
			cn.RollbackTrans
		'	~~~~~~~~~~~~~~~~
			alerta = "Falha ao tentar alterar este registro de entrada de mercadorias no estoque:"
			alerta = texto_add_br(alerta) & msg_erro
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



<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>



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

<% if alerta <> "" then %>
<!-- ************************************************************ -->
<!-- **********  PÁGINA PARA EXIBIR MENSAGENS DE ERRO  ********** -->
<!-- ************************************************************ -->
<body onload="bVOLTAR.focus();">
<center>
<br>
<!--  T E L A  -->
<p class="T">A V I S O</p>
<div class="MtAlerta" style="width:600px;FONT-WEIGHT:bold;" align="center"><P style='margin:5px 2px 5px 2px;'><%=alerta%></p></div>
<BR><BR>
<p class="TracoBottom"></p>
<table cellSpacing="0">
<tr>
	<td align="CENTER"><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()"><img src="..\botao\voltar.gif" width="176" height="55" border="0"></a></td>
</tr>
</table>
</center>
</body>


<% else %>
<!-- *************************************************************** -->
<!-- **********  PÁGINA PARA EXIBIR MENSAGENS DE SUCESSO  ********** -->
<!-- *************************************************************** -->
<body>
<center>
<br>
<!--  T E L A  -->
<p class="T">A V I S O</p>
<div class="MtAviso" style="width:600px;FONT-WEIGHT:bold;" align="center"><P style='margin:5px 2px 5px 2px;'><%=mensagem%></p></div>
<BR><BR>
<p class="TracoBottom"></p>
<table cellSpacing="0">
<tr>
	<td align="CENTER"><a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>"><img src="..\botao\voltar.gif" width="176" height="55" border="0"></a></td>
</tr>
</table>
</center>
</body>

<% end if %>

</html>


<%
	if tEI.State <> 0 then tEI.Close
	set tEI = nothing

'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>