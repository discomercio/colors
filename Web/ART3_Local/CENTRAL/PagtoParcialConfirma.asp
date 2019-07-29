<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ========================================================
'	  P A G T O P A R C I A L C O N F I R M A . A S P
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

	dim s, usuario, msg_erro, s_log
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

	dim pedido_selecionado, s_valor, m_valor, s_chave
	dim vl_TotalFamiliaPrecoVenda, vl_TotalFamiliaPrecoNF, vl_TotalFamiliaPago, vl_TotalFamiliaDevolucaoPrecoVenda, vl_TotalFamiliaDevolucaoPrecoNF, st_pagto, id_pedido_base
'	OBTÉM DADOS DO FORMULÁRIO
	pedido_selecionado = ucase(Trim(request("pedido_selecionado")))
	if (pedido_selecionado = "") then Response.Redirect("aviso.asp?id=" & ERR_PEDIDO_NAO_ESPECIFICADO)
	pedido_selecionado=normaliza_num_pedido(pedido_selecionado)

	id_pedido_base = retorna_num_pedido_base(pedido_selecionado)

	s_valor = Trim(Request.Form("c_valor"))
	m_valor = converte_numero(s_valor)

	dim intNumParcelasPagtoCartao
	dim alerta
	alerta=""

'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	if alerta = "" then
		if Not calcula_pagamentos(pedido_selecionado, vl_TotalFamiliaPrecoVenda, vl_TotalFamiliaPrecoNF, vl_TotalFamiliaPago, vl_TotalFamiliaDevolucaoPrecoVenda, vl_TotalFamiliaDevolucaoPrecoNF, st_pagto, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
		if Not gera_nsu(NSU_PEDIDO_PAGAMENTO, s_chave, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_GERAR_NSU)

	'	~~~~~~~~~~~~~
		cn.BeginTrans
	'	~~~~~~~~~~~~~
		If Not cria_recordset_pessimista(rs, msg_erro) then 
		'	~~~~~~~~~~~~~~~~
			cn.RollbackTrans
		'	~~~~~~~~~~~~~~~~
			Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
			end if

		s = "INSERT INTO t_PEDIDO_PAGAMENTO" & _
			" (id, pedido, data, hora, valor, usuario, tipo_pagto)" & _
			" VALUES (" & _
			" '" & s_chave & "'" & _
			", '" & pedido_selecionado & "'" & _
			", " & bd_formata_data(Date) & _
			", '" & retorna_so_digitos(formata_hora(Now)) & "'" & _
			", " & bd_formata_numero(m_valor) & _
			", '" & usuario & "'" & _
			", '" & COD_PAGTO_PARCIAL & "'" & _
			")"
		cn.Execute(s)
		if Err <> 0 then
			alerta=Cstr(Err) & ": " & Err.Description
		else
			s = "SELECT * FROM t_PEDIDO WHERE (pedido = '" & id_pedido_base & "')"
			if rs.State <> 0 then rs.Close
			rs.Open s, cn
			if rs.Eof then
				alerta = "Pedido-base " & id_pedido_base & " não foi encontrado."
			else
			'	CALCULA QUANTIDADE DE PARCELAS EM CARTÃO DE CRÉDITO
				intNumParcelasPagtoCartao = 0
				if Trim("" & rs("tipo_parcelamento")) = Trim("" & COD_FORMA_PAGTO_A_VISTA) then
					if Trim("" & rs("av_forma_pagto")) = Trim("" & ID_FORMA_PAGTO_CARTAO) then intNumParcelasPagtoCartao = 1
				elseif Trim("" & rs("tipo_parcelamento")) = Trim("" & COD_FORMA_PAGTO_PARCELA_UNICA) then
					if Trim("" & rs("pu_forma_pagto")) = Trim("" & ID_FORMA_PAGTO_CARTAO) then intNumParcelasPagtoCartao = 1
				elseif Trim("" & rs("tipo_parcelamento")) = Trim("" & COD_FORMA_PAGTO_PARCELADO_CARTAO) then
					intNumParcelasPagtoCartao = rs("pc_qtde_parcelas")
				elseif Trim("" & rs("tipo_parcelamento")) = Trim("" & COD_FORMA_PAGTO_PARCELADO_CARTAO_MAQUINETA) then
					'NOP
				elseif Trim("" & rs("tipo_parcelamento")) = Trim("" & COD_FORMA_PAGTO_PARCELADO_COM_ENTRADA) then
				'	ENTRADA + PRESTAÇÕES
					if Trim("" & rs("pce_forma_pagto_entrada")) = Trim("" & ID_FORMA_PAGTO_CARTAO) then intNumParcelasPagtoCartao = intNumParcelasPagtoCartao + 1
					if Trim("" & rs("pce_forma_pagto_prestacao")) = Trim("" & ID_FORMA_PAGTO_CARTAO) then intNumParcelasPagtoCartao = intNumParcelasPagtoCartao + rs("pce_prestacao_qtde")
				elseif Trim("" & rs("tipo_parcelamento")) = Trim("" & COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA) then
				'	1ª PRESTAÇÃO + DEMAIS PRESTAÇÕES
					if Trim("" & rs("pse_forma_pagto_prim_prest")) = Trim("" & ID_FORMA_PAGTO_CARTAO) then intNumParcelasPagtoCartao = intNumParcelasPagtoCartao + 1
					if Trim("" & rs("pse_forma_pagto_demais_prest")) = Trim("" & ID_FORMA_PAGTO_CARTAO) then intNumParcelasPagtoCartao = intNumParcelasPagtoCartao + rs("pse_demais_prest_qtde")
					end if
				
			'	PAGO (QUITADO)
			'	~~~~~~~~~~~~~~
				if (vl_TotalFamiliaDevolucaoPrecoNF + vl_TotalFamiliaPago + m_valor) >= (vl_TotalFamiliaPrecoNF - MAX_VALOR_MARGEM_ERRO_PAGAMENTO) then
					if Trim("" & rs("st_pagto")) <> ST_PAGTO_PAGO then
						rs("dt_st_pagto") = Date
						rs("dt_hr_st_pagto") = Now
						rs("usuario_st_pagto") = usuario
						end if
					rs("st_pagto") = ST_PAGTO_PAGO
					s_log = "quitado"
					if (vl_TotalFamiliaDevolucaoPrecoNF + vl_TotalFamiliaPago + m_valor) > vl_TotalFamiliaPrecoNF then 
						s_log = s_log & " (excedeu " & SIMBOLO_MONETARIO & " " & _
								formata_moeda((vl_TotalFamiliaDevolucaoPrecoNF+vl_TotalFamiliaPago+m_valor)-vl_TotalFamiliaPrecoNF) & ")"
					elseif (vl_TotalFamiliaDevolucaoPrecoNF + vl_TotalFamiliaPago + m_valor) < vl_TotalFamiliaPrecoNF then
						s_log = s_log & " (faltou " & SIMBOLO_MONETARIO & " " & _
								formata_moeda(vl_TotalFamiliaPrecoNF-(vl_TotalFamiliaDevolucaoPrecoNF+vl_TotalFamiliaPago+m_valor)) & ")"
						end if
				'	ANÁLISE DE CRÉDITO
					if Trim("" & rs("loja")) = NUMERO_LOJA_ECOMMERCE_AR_CLUBE then
						if (CLng(rs("analise_credito")) = CLng(COD_AN_CREDITO_ST_INICIAL)) then
							if s_log <> "" then s_log = s_log & "; "
							s_log = s_log & " Análise de crédito: " & descricao_analise_credito(rs("analise_credito")) & " => " & descricao_analise_credito(COD_AN_CREDITO_OK)
							rs("analise_credito") = CLng(COD_AN_CREDITO_OK)
							rs("analise_credito_data")=Now
							rs("analise_credito_usuario")=ID_USUARIO_SISTEMA
							end if
					else
					'	TRATAMENTO P/ O CASO EM QUE A CIELO INFORMA O STATUS "1 - EM ANDAMENTO" QUANDO ATIVA A PÁGINA DE RETORNO, MAS ACABA
					'	AUTORIZANDO A TRANSAÇÃO. NESTE CASO, O REGISTRO DO PAGAMENTO PRECISA SER FEITO MANUALMENTE NO PEDIDO.
						if (CLng(rs("analise_credito")) = CLng(COD_AN_CREDITO_ST_INICIAL)) And CLng(rs("st_forma_pagto_somente_cartao")) = 1 then
						'	SE HOUVER PARCELA PAGA EM CARTÃO, A EQUIPE DA ANÁLISE DE CRÉDITO DEVE VERIFICAR A NECESSIDADE DE PEDIR DOCUMENTAÇÃO COMPLEMENTAR AO CLIENTE
							If (intNumParcelasPagtoCartao = 0) then
								if s_log <> "" then s_log = s_log & "; "
								s_log = s_log & " Análise de crédito: " & descricao_analise_credito(rs("analise_credito")) & " => " & descricao_analise_credito(COD_AN_CREDITO_OK)
								rs("analise_credito") = CLng(COD_AN_CREDITO_OK)
								rs("analise_credito_data")=Now
								rs("analise_credito_usuario")=ID_USUARIO_SISTEMA
							else
								if s_log <> "" then s_log = s_log & "; "
								s_log = s_log & " Análise de crédito: " & descricao_analise_credito(rs("analise_credito")) & " => " & descricao_analise_credito(COD_AN_CREDITO_PENDENTE_VENDAS)
								rs("analise_credito") = CLng(COD_AN_CREDITO_PENDENTE_VENDAS)
								rs("analise_credito_data")=Now
								rs("analise_credito_usuario")=ID_USUARIO_SISTEMA
								end if
							end if
						end if
			'	PAGAMENTO PARCIAL
			'	~~~~~~~~~~~~~~~~~
				elseif (vl_TotalFamiliaPago + m_valor) > 0 then
					if Trim("" & rs("st_pagto")) <> ST_PAGTO_PARCIAL then
						rs("dt_st_pagto") = Date
						rs("dt_hr_st_pagto") = Now
						rs("usuario_st_pagto") = usuario
						end if
					rs("st_pagto") = ST_PAGTO_PARCIAL
					s_log = "pago parcial"
			'	NÃO PAGO
			'	~~~~~~~~
				else
					if Trim("" & rs("st_pagto")) <> ST_PAGTO_NAO_PAGO then
						rs("dt_st_pagto") = Date
						rs("dt_hr_st_pagto") = Now
						rs("usuario_st_pagto") = usuario
						end if
					rs("st_pagto") = ST_PAGTO_NAO_PAGO
					s_log = "não-pago"
					end if
				
				rs("vl_pago_familia") = vl_TotalFamiliaPago + m_valor
				s_log = "; status=" & s_log
				rs.Update
				if Err <> 0 then
					alerta=Cstr(Err) & ": " & Err.Description
					end if
				end if
			end if
		
		if alerta = "" then
			s_log = SIMBOLO_MONETARIO & " " & formata_moeda(m_valor) & s_log
			grava_log usuario, "", pedido_selecionado, "", OP_LOG_PEDIDO_PAGTO_PARCIAL, s_log
		'	~~~~~~~~~~~~~~
			cn.CommitTrans
		'	~~~~~~~~~~~~~~
			if Err=0 then 
				Response.Redirect("resumo.asp" & "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")))
			else
				alerta=Cstr(Err) & ": " & Err.Description
				end if
		else
		'	~~~~~~~~~~~~~~~~
			cn.RollbackTrans
		'	~~~~~~~~~~~~~~~~
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
<link href="<%=URL_FILE__EPRINTER_CSS%>" Rel="stylesheet" Type="text/css" media="print">

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
<table class="notPrint" cellSpacing="0">
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