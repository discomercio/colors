<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ========================================================
'	  PedidoAnotaFreteConfirma.asp
'     ========================================================
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

	dim s, s_log, usuario
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

	dim alerta
	alerta=""

'	OBTÉM DADOS DO FORMULÁRIO
	dim intCounter, intQtdeItens
	dim v_item
    dim emissor_cnpj, id_nfe_emitente, serie_NF, numero_NF
    dim intNsuNovoPedidoFrete
    dim url_origem
    url_origem = Trim(Request.Form("c_url_origem"))
	redim v_item(0)
	set v_item(0) = New cl_PEDIDO_ANOTA_FRETE
	intQtdeItens = Request.Form("c_pedido").Count
	for intCounter = 1 to intQtdeItens
		s = Trim(Request.Form("c_pedido")(intCounter))
		if s <> "" then
			if Trim(v_item(Ubound(v_item)).pedido) <> "" then
				redim preserve v_item(Ubound(v_item)+1)
				set v_item(Ubound(v_item)) = New cl_PEDIDO_ANOTA_FRETE
				end if
			with v_item(Ubound(v_item))
				.pedido = UCase(Trim(Request.Form("c_pedido")(intCounter)))
				s = Trim(Request.Form("c_valor_frete")(intCounter))
				.vl_frete = converte_numero(s)
                .tipo_frete = Request.Form("c_tipo_frete")(intCounter)
                .transportadora_id = Trim(Request.Form("c_transportadora_id")(intCounter))
                .transportadora_cnpj = Trim(Request.Form("c_transportadora_cnpj")(intCounter))
                .num_NF = Trim(Request.Form("c_NF")(intCounter))
                .serie_NF = Trim(Request.Form("c_serie_NF")(intCounter))
                .emitente_NF = Trim(Request.Form("c_emitente_NF")(intCounter))
                .vl_NF_devolucao = Trim(Request.Form("c_valor_nf")(intCounter))
                .vl_total_NF = Trim(Request.Form("c_valor_total_nf")(intCounter))
				end with
			end if
		next

	for intCounter = Lbound(v_item) to Ubound(v_item)
		with v_item(intCounter)
			if Trim("" & .pedido) <> "" then
				if (.vl_frete < 0) then
					alerta = "Valor do frete do pedido " & Trim(.pedido) & " é inválido (" & formata_moeda(.vl_frete) & ")."
					end if
				end if
			end with
		next

'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs, msg_erro
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	if alerta = "" then
	'	~~~~~~~~~~~~~
		cn.BeginTrans
	'	~~~~~~~~~~~~~
		If Not cria_recordset_pessimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

    ' Captura os dados da transportadora selecionada


	'	INFORMAÇÕES PARA O LOG
		s_log = ""
		for intCounter = Lbound(v_item) to Ubound(v_item)
			with v_item(intCounter)
                s = "SELECT tNFE.NFe_numero_NF AS numero_NF" & _
	                    " ,tNFE.NFe_serie_NF AS serie_NF" & _
	                    " ,tEMT.cnpj AS emissor_cnpj" & _
	                    " ,tEMT.id AS id_nfe_emitente" & _
                    " FROM t_NFe_EMISSAO tNFE" & _
                    " LEFT JOIN t_PEDIDO tPED ON (tNFE.pedido = tPED.pedido)" & _
                    " LEFT JOIN t_NFe_EMITENTE tEMT ON (tNFE.id_nfe_emitente = tEMT.id)" & _
                    " WHERE tNFE.pedido = '" & .pedido & "'"

                if rs.State <> 0 then rs.Close
				rs.open s, cn

                if Not rs.Eof then
                    emissor_cnpj = Trim("" & rs("emissor_cnpj"))
                    'No caso dos CDs ES e ES-02, grava o id_nfe_emitente real do pedido, mesmo que na consulta tenha sido informado o outro CD (lembrando que a operação faz a validação pelo CNPJ do CD e não mais pelo id_nfe_emitente).
                    if .emitente_NF <> "-1" then id_nfe_emitente = Trim("" & rs("id_nfe_emitente"))
                    if .num_NF = "" then numero_NF = Trim("" & rs("numero_NF"))
                    if .serie_NF = "" then serie_NF = Trim("" & rs("serie_NF"))
                end if               
                if .emitente_NF = "-1" then id_nfe_emitente = .emitente_NF
                if .serie_NF <> "" then serie_NF = .serie_NF
                if .num_NF <> "" then numero_NF = .num_NF
				if Trim("" & .pedido) <> "" then
					s = "SELECT * FROM t_PEDIDO_FRETE WHERE (pedido='-1')"
					if rs.State <> 0 then rs.Close

                    if Not fin_gera_nsu(T_PEDIDO_FRETE, intNsuNovoPedidoFrete, msg_erro) then 
			            alerta = "FALHA AO GERAR NSU PARA O NOVO REGISTRO (" & msg_erro & ")"
		            else
			            if intNsuNovoPedidoFrete <= 0 then
				            alerta = "NSU GERADO É INVÁLIDO (" & intNsuNovoPedidoFrete & ")"
				            end if
			            end if
                    
					rs.open s, cn
					if Err <> 0 then
					'	~~~~~~~~~~~~~~~~
						cn.RollbackTrans
					'	~~~~~~~~~~~~~~~~
						Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
						end if

                        rs.AddNew
						rs("id") = intNsuNovoPedidoFrete
						rs("pedido") = .pedido
						rs("codigo_tipo_frete") = .tipo_frete
						rs("vl_frete") = .vl_frete
                        rs("transportadora_id") = .transportadora_id
                        rs("transportadora_cnpj") = .transportadora_cnpj
                        rs("id_nfe_emitente") = id_nfe_emitente
                        rs("serie_NF") = serie_NF
                        rs("numero_NF") = numero_NF
                        rs("tipo_preenchimento") = 1
                        rs("usuario_cadastro") = usuario
                        rs("usuario_ult_atualizacao") = usuario

                        if .vl_NF_devolucao <> "" then 
                            rs("vl_NF") = converte_numero(.vl_NF_devolucao)
                        else
                            rs("vl_NF") = converte_numero(.vl_total_NF)
                        end if
                        
                        if id_nfe_emitente = "-1" then 
                            rs("emissor_cnpj") = ""
                        else
                            rs("emissor_cnpj") = emissor_cnpj
                        end if
                        
						rs.Update
						
					if Err <> 0 then
					'	~~~~~~~~~~~~~~~~
						cn.RollbackTrans
					'	~~~~~~~~~~~~~~~~
						Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
						end if
					
				'	INFORMAÇÕES PARA O LOG
					if s_log <> "" then s_log = s_log & "; "
					s_log = s_log & _
							Trim(.pedido) & "=" & _
							formata_moeda(.vl_frete) & "|" & _
                            "Tipo Frete=" & .tipo_frete & "|" & _
                            "Transportadora=" & .transportadora_id
					end if  'if (blnTemItem)
				end with
			next

	'	INFORMAÇÕES PARA O LOG
		s_log = "Anotar frete no pedido: " & s_log
		grava_log usuario, "", "", "", OP_LOG_ANOTA_FRETE_PEDIDO, s_log
		
	'	~~~~~~~~~~~~~~
		cn.CommitTrans
	'	~~~~~~~~~~~~~~
		if Err=0 then 
			s = "Valor do frete cadastrado com sucesso no(s) pedido(s)"
			Session(SESSION_CLIPBOARD) = s
			Response.Redirect("mensagem.asp?url_back=" & server.URLEncode(url_origem) & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")))
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
<div class="MtAlerta" style="width:600px;font-weight:bold;" align="center"><P style='margin:5px 2px 5px 2px;'><%=alerta%></p></div>
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