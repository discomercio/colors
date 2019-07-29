<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =================================================================
'	  R E L D E V O L U C A O E X E C . A S P
'     =================================================================
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

	Server.ScriptTimeout = MAX_SERVER_SCRIPT_TIMEOUT_EM_SEG
	
	dim usuario, loja
	usuario = Trim(Session("usuario_atual"))
	loja = Trim(Session("loja_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 
	If (loja = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	CONECTA COM O BANCO DE DADOS
	dim cn, rs, msg_erro
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	if Not operacao_permitida(OP_LJA_REL_DEVOLUCAO_PRODUTOS, s_lista_operacoes_permitidas) then 
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	dim alerta
	dim s, s_aux, s_filtro, flag_ok
	dim ckb_periodo_devolucao, c_dt_devolucao_inicio, c_dt_devolucao_termino
	dim ckb_periodo_cadastro, c_dt_cadastro_inicio, c_dt_cadastro_termino
	dim ckb_produto, c_fabricante, c_produto
	dim c_loja

	alerta = ""

	ckb_periodo_devolucao = Trim(Request.Form("ckb_periodo_devolucao"))
	c_dt_devolucao_inicio = Trim(Request.Form("c_dt_devolucao_inicio"))
	c_dt_devolucao_termino = Trim(Request.Form("c_dt_devolucao_termino"))
	ckb_periodo_cadastro = Trim(Request.Form("ckb_periodo_cadastro"))
	c_dt_cadastro_inicio = Trim(Request.Form("c_dt_cadastro_inicio"))
	c_dt_cadastro_termino = Trim(Request.Form("c_dt_cadastro_termino"))
	ckb_produto = Trim(Request.Form("ckb_produto"))
	c_fabricante = retorna_so_digitos(Trim(Request.Form("c_fabricante")))
	c_produto = Ucase(Trim(Request.Form("c_produto")))
	c_loja = Trim(Request.Form("c_loja"))
	
	if alerta = "" then
		if c_loja = "" then
			alerta=texto_add_br(alerta)
			alerta = "NÃO FOI INFORMADO O Nº DA LOJA."
			end if
		end if
		
	if alerta = "" then
		if c_fabricante <> "" then
			s = "SELECT fabricante FROM t_FABRICANTE WHERE (fabricante='" & c_fabricante & "')"
			if rs.State <> 0 then rs.Close
			rs.open s, cn
			if rs.Eof then 
				alerta=texto_add_br(alerta)
				alerta = "FABRICANTE " & c_fabricante & " NÃO ESTÁ CADASTRADO."
				end if
			end if
		end if
	
	if alerta = "" then
		if c_produto <> "" then
			if (Not IsEAN(c_produto)) And (c_fabricante="") then
				alerta=texto_add_br(alerta)
				alerta=alerta & "NÃO FOI ESPECIFICADO O FABRICANTE DO PRODUTO A SER CONSULTADO."
			else
				s = "SELECT * FROM t_PRODUTO WHERE"
				if IsEAN(c_produto) then
					s = s & " (ean='" & c_produto & "')"
				else
					s = s & " (fabricante='" & c_fabricante & "') AND (produto='" & c_produto & "')"
					end if
				if rs.State <> 0 then rs.Close
				rs.open s, cn
				if Not rs.Eof then
					flag_ok = True
					if IsEAN(c_produto) And (c_fabricante<>"") then
						if (c_fabricante<>Trim("" & rs("fabricante"))) then
							flag_ok = False
							alerta=texto_add_br(alerta)
							alerta=alerta & "Produto a ser consultado " & c_produto & " NÃO pertence ao fabricante " & c_fabricante & "."
							end if
						end if
					if flag_ok then
					'	CARREGA CÓDIGO INTERNO DO PRODUTO
						c_fabricante = Trim("" & rs("fabricante"))
						c_produto = Trim("" & rs("produto"))
						end if
					end if
				end if
			end if
		end if
		

'	Período de consulta está restrito por perfil de acesso?
	dim dtMinDtInicialFiltroPeriodo, intMaxDiasDtInicialFiltroPeriodo
	dim strMinDtInicialFiltroPeriodoYYYYMMDD, strMinDtInicialFiltroPeriodoDDMMYYYY
	dim strDtRefDDMMYYYY
	if operacao_permitida(OP_LJA_RESTRINGE_DT_INICIAL_FILTRO_PERIODO, s_lista_operacoes_permitidas) then
		intMaxDiasDtInicialFiltroPeriodo = obtem_max_dias_dt_inicial_filtro_periodo()
		dtMinDtInicialFiltroPeriodo = Date - intMaxDiasDtInicialFiltroPeriodo
		strMinDtInicialFiltroPeriodoYYYYMMDD = formata_data_yyyymmdd(dtMinDtInicialFiltroPeriodo)
		strMinDtInicialFiltroPeriodoDDMMYYYY = formata_data(dtMinDtInicialFiltroPeriodo)
	'	PERÍODO DE DEVOLUÇÃO
		if alerta = "" then
			strDtRefDDMMYYYY = c_dt_devolucao_inicio
			if strDtRefDDMMYYYY <> "" then
				if StrToDate(strDtRefDDMMYYYY) < dtMinDtInicialFiltroPeriodo then
					alerta = "Data inválida para consulta: " & strDtRefDDMMYYYY & ".  O período de consulta não pode compreender datas anteriores a " & strMinDtInicialFiltroPeriodoDDMMYYYY
					end if
				end if
			end if

		if alerta = "" then
			strDtRefDDMMYYYY = c_dt_devolucao_termino
			if strDtRefDDMMYYYY <> "" then
				if StrToDate(strDtRefDDMMYYYY) < dtMinDtInicialFiltroPeriodo then
					alerta = "Data inválida para consulta: " & strDtRefDDMMYYYY & ".  O período de consulta não pode compreender datas anteriores a " & strMinDtInicialFiltroPeriodoDDMMYYYY
					end if
				end if
			end if

		if alerta = "" then
			if c_dt_devolucao_inicio = "" then c_dt_devolucao_inicio = strMinDtInicialFiltroPeriodoDDMMYYYY
			end if
			
	'	PERÍODO DE CADASTRO
		if alerta = "" then
			strDtRefDDMMYYYY = c_dt_cadastro_inicio
			if strDtRefDDMMYYYY <> "" then
				if StrToDate(strDtRefDDMMYYYY) < dtMinDtInicialFiltroPeriodo then
					alerta = "Data inválida para consulta: " & strDtRefDDMMYYYY & ".  O período de consulta não pode compreender datas anteriores a " & strMinDtInicialFiltroPeriodoDDMMYYYY
					end if
				end if
			end if

		if alerta = "" then
			strDtRefDDMMYYYY = c_dt_cadastro_termino
			if strDtRefDDMMYYYY <> "" then
				if StrToDate(strDtRefDDMMYYYY) < dtMinDtInicialFiltroPeriodo then
					alerta = "Data inválida para consulta: " & strDtRefDDMMYYYY & ".  O período de consulta não pode compreender datas anteriores a " & strMinDtInicialFiltroPeriodoDDMMYYYY
					end if
				end if
			end if

		if alerta = "" then
			if c_dt_cadastro_inicio = "" then c_dt_cadastro_inicio = strMinDtInicialFiltroPeriodoDDMMYYYY
			end if
	
	else
		strMinDtInicialFiltroPeriodoYYYYMMDD = ""
		strMinDtInicialFiltroPeriodoDDMMYYYY = ""
		end if

	




' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________

' _____________________________________
' CONSULTA EXECUTA
'
sub consulta_executa
dim r
dim s, s_aux, s_sql, cab_table, cab, n_reg, n_reg_total
dim s_where, s_from
dim vl_total_devolucao, vl_sub_total_devolucao
dim x, loja_a, qtde_lojas
dim qtde_total, qtde_sub_total
dim w_cliente, w_produto

'	MONTA CLÁUSULA WHERE
	s_where = ""

'	CRITÉRIO: PERÍODO DE DEVOLUÇÃO
	s = ""
	if c_dt_devolucao_inicio <> "" then
		if s <> "" then s = s & " AND"
		s = s & " (t_PEDIDO_ITEM_DEVOLVIDO.devolucao_data >= " & bd_formata_data(StrToDate(c_dt_devolucao_inicio)) & ")"
		end if
		
	if c_dt_devolucao_termino <> "" then
		if s <> "" then s = s & " AND"
		s = s & " (t_PEDIDO_ITEM_DEVOLVIDO.devolucao_data < " & bd_formata_data(StrToDate(c_dt_devolucao_termino)+1) & ")"
		end if
		
	if s <> "" then 
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (" & s & ")"
		end if
		
'	CRITÉRIO: PERÍODO DE CADASTRAMENTO DO PEDIDO
	s = ""
	if c_dt_cadastro_inicio <> "" then
		if s <> "" then s = s & " AND"
		s = s & " (t_PEDIDO.data >= " & bd_formata_data(StrToDate(c_dt_cadastro_inicio)) & ")"
		end if
		
	if c_dt_cadastro_termino <> "" then
		if s <> "" then s = s & " AND"
		s = s & " (t_PEDIDO.data < " & bd_formata_data(StrToDate(c_dt_cadastro_termino)+1) & ")"
		end if
		
	if s <> "" then 
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (" & s & ")"
		end if
		
'	CRITÉRIO: PRODUTO
	if ckb_produto <> "" then
		s = ""
		if c_fabricante <> "" then
			if s <> "" then s = s & " AND"
			s = s & " (t_PEDIDO_ITEM_DEVOLVIDO.fabricante = '" & c_fabricante & "')"
			end if
		
		if c_produto <> "" then
			if s <> "" then s = s & " AND"
			s = s & " (t_PEDIDO_ITEM_DEVOLVIDO.produto = '" & c_produto & "')"
			end if
		
		if s <> "" then 
			if s_where <> "" then s_where = s_where & " AND"
			s_where = s_where & " (" & s & ")"
			end if
		end if

'	CRITÉRIO: LOJA
	s = " (CONVERT(smallint, t_PEDIDO.loja) = " & c_loja & ")"
	if s_where <> "" then s_where = s_where & " AND"
	s_where = s_where & " (" & s & ")"
	
	if s_where <> "" then s_where = " WHERE" & s_where
	
	
'	MONTA CLÁUSULA FROM
	s_from = " FROM t_PEDIDO" & _
			 " INNER JOIN t_PEDIDO_ITEM_DEVOLVIDO ON (t_PEDIDO.pedido=t_PEDIDO_ITEM_DEVOLVIDO.pedido)" & _
			 " LEFT JOIN t_CLIENTE ON (t_PEDIDO.id_cliente=t_CLIENTE.id)"

	s_sql = "SELECT t_PEDIDO.loja, CONVERT(smallint,t_PEDIDO.loja) AS numero_loja," & _
			" t_PEDIDO.data, t_PEDIDO.pedido," & _
			" t_CLIENTE.nome_iniciais_em_maiusculas," & _
			" t_PEDIDO_ITEM_DEVOLVIDO.devolucao_data, t_PEDIDO_ITEM_DEVOLVIDO.fabricante," & _
			" t_PEDIDO_ITEM_DEVOLVIDO.produto," & _
			" t_PEDIDO_ITEM_DEVOLVIDO.descricao," & _
			" t_PEDIDO_ITEM_DEVOLVIDO.descricao_html," & _
			" Sum(t_PEDIDO_ITEM_DEVOLVIDO.qtde) AS qtde, Sum(t_PEDIDO_ITEM_DEVOLVIDO.qtde*t_PEDIDO_ITEM_DEVOLVIDO.preco_venda) AS valor_devolucao" & _
			s_from & _
			s_where
			
	s_sql = s_sql & " GROUP BY t_PEDIDO.loja, t_PEDIDO.data, t_PEDIDO.pedido, t_CLIENTE.nome_iniciais_em_maiusculas, t_PEDIDO_ITEM_DEVOLVIDO.devolucao_data, t_PEDIDO_ITEM_DEVOLVIDO.fabricante, t_PEDIDO_ITEM_DEVOLVIDO.produto, t_PEDIDO_ITEM_DEVOLVIDO.descricao, t_PEDIDO_ITEM_DEVOLVIDO.descricao_html" & _
					" ORDER BY numero_loja, t_PEDIDO.data, t_PEDIDO.pedido"

  ' CABEÇALHO
	w_cliente = 173
	w_produto = 170
	cab_table = "<TABLE cellSpacing=0>" & chr(13)
	cab = "	<TR style='background:azure'>" & chr(13) & _
		  "		<TD class='MDTE' style='width:70px' valign='bottom' NOWRAP><P class='R'>Nº Pedido</P></TD>" & chr(13) & _
		  "		<TD class='MTD' style='width:" & cstr(w_cliente) & "px' valign='bottom'><P class='R'>Cliente</P></TD>" & chr(13) & _
		  "		<TD class='MTD' style='width:64px' valign='bottom'><P class='R'>Data Devolução</P></TD>" & chr(13) & _
		  "		<TD class='MTD' style='width:35px' valign='bottom'><P class='R'>Fabr</P></TD>" & chr(13) & _
		  "		<TD class='MTD' style='width:" & cstr(w_produto) & "px' valign='bottom'><P class='R'>Produto</P></TD>" & chr(13) & _
		  "		<TD class='MTD' style='width:35px' align='right' valign='bottom'><P class='Rd' style='font-weight:bold;'>Qtde</P></TD>" & chr(13) & _
		  "		<TD class='MTD' style='width:80px' align='right' valign='bottom'><P class='Rd' style='font-weight:bold;'>Valor Devolução</P></TD>" & chr(13) & _
		  "	</TR>" & chr(13)
	
	x = ""
	n_reg = 0
	n_reg_total = 0
	qtde_lojas = 0
	vl_total_devolucao = 0
	qtde_total = 0
	
	loja_a = "XXXXX"
	
	set r = cn.execute(s_sql)
	do while Not r.Eof
	'	MUDOU DE LOJA?
		if Trim("" & r("loja"))<>loja_a then
			loja_a = Trim("" & r("loja"))
			qtde_lojas = qtde_lojas + 1
		  ' FECHA TABELA DA LOJA ANTERIOR
			if n_reg_total > 0 then 
				x = x & "	<TR style='background: #FFFFDD'>" & chr(13) & _
						"		<TD class='MTBE' COLSPAN='5' NOWRAP><p class='Cd'>" & _
						"TOTAL:</p></td>" & chr(13) & _
						"		<TD class='MTB'><p class='Cd'>" & formata_inteiro(qtde_sub_total) & "</p></td>" & chr(13) & _
						"		<TD class='MTBD'><p class='Cd'>" & formata_moeda(vl_sub_total_devolucao) & "</p></td>" & chr(13) & _
						"	</TR>" & chr(13) & _
						"</TABLE>" & chr(13)
				Response.Write x
				x="<BR>"
				end if

			n_reg = 0
			vl_sub_total_devolucao = 0
			qtde_sub_total = 0

			if n_reg_total > 0 then x = x & "<BR>" & chr(13)
			s = Trim("" & r("loja"))
			s_aux = x_loja(s)
			if (s<>"") And (s_aux<>"") then s = s & " - "
			s = s & s_aux
			x = x & cab_table
			if s <> "" then x = x & "	<TR>" & chr(13) & _
									"		<TD class='MDTE' COLSPAN='7' valign='bottom' class='MB' style='background:azure;'><p class='N'>&nbsp;" & s & "</p></td>" & chr(13) & _
									"	</tr>" & chr(13)
			x = x & cab
			end if

	 ' CONTAGEM
		n_reg = n_reg + 1
		n_reg_total = n_reg_total + 1

		x = x & "	<TR>" & chr(13)

	'> Nº PEDIDO
		x = x & "		<TD valign='top' class='MDTE'><P class='C'>&nbsp;<a href='javascript:fRELConcluir(" & _
				chr(34) & Trim("" & r("pedido")) & chr(34) & _
				")' title='clique para consultar o pedido'>" & Trim("" & r("pedido")) & "</a></P></TD>" & chr(13)

	'> CLIENTE
		s = Trim("" & r("nome_iniciais_em_maiusculas"))
		if s = "" then s = "&nbsp;"
		x = x & "		<TD valign='top' style='width:" & cstr(w_cliente) & "px;' class='MTD'><P class='Cn'>" & s & "</P></TD>" & chr(13)

	'> DATA DE DEVOLUÇÃO
		s = formata_data(r("devolucao_data"))
		if s = "" then s = "&nbsp;"
		x = x & "		<TD valign='top' class='MTD'><P class='Cnc'>" & s & "</P></TD>" & chr(13)
		
	'> FABRICANTE
		s = Trim("" & r("fabricante"))
		if s = "" then s = "&nbsp;"
		x = x & "		<TD valign='top' class='MTD'><P class='Cn'>" & s & "</P></TD>" & chr(13)

	'> PRODUTO
		s = Trim("" & r("produto"))
		s_aux = Trim("" & r("descricao_html"))
		if s_aux <> "" then s_aux = produto_formata_descricao_em_html(s_aux)
		if (s<>"") And (s_aux<>"") then s = s & " - "
		s = s & s_aux
		if s = "" then s = "&nbsp;"
		x = x & "		<TD valign='top' style='width:" & cstr(w_produto) & "px;' class='MTD'><P class='Cn'>" & s & "</P></TD>" & chr(13)

	'> QTDE
		s = formata_inteiro(r("qtde"))
		if s = "" then s = "&nbsp;"
		x = x & "		<TD valign='top' align='right' class='MTD'><P class='Cnd'>" & s & "</P></TD>" & chr(13)

	'> VALOR DA DEVOLUÇÃO
		s = formata_moeda(r("valor_devolucao"))
		if s = "" then s = "&nbsp;"
		x = x & "		<TD valign='top' align='right' class='MTD'><P class='Cnd'>" & s & "</P></TD>" & chr(13)

	'> TOTALIZAÇÃO DE VALORES
		vl_sub_total_devolucao = vl_sub_total_devolucao + r("valor_devolucao")
		qtde_sub_total = qtde_sub_total + r("qtde")
		
		vl_total_devolucao = vl_total_devolucao + r("valor_devolucao")
		qtde_total = qtde_total + r("qtde")
			
		x = x & "	</TR>" & chr(13)
	
		if (n_reg_total mod 100) = 0 then
			Response.Write x
			x = ""
			end if
		
		r.MoveNext
		loop
	
  ' MOSTRA TOTAL DA ÚLTIMA LOJA
	if n_reg <> 0 then 
		x = x & "	<TR style='background: #FFFFDD'>" & chr(13) & _
				"		<TD COLSPAN='5' class='MTBE' NOWRAP><p class='Cd'>" & _
				"TOTAL:</p></td>" & chr(13) & _
				"		<TD class='MTB'><p class='Cd'>" & formata_inteiro(qtde_sub_total) & "</p></td>" & chr(13) & _
				"		<TD class='MTBD'><p class='Cd'>" & formata_moeda(vl_sub_total_devolucao) & "</p></td>" & chr(13) & _
				"	</TR>" & chr(13)
		
	'>	TOTAL GERAL
		if qtde_lojas > 1 then
			x = x & "	<TR>" & chr(13) & _
					"		<TD COLSPAN='7' style='border-left:0px;border-right:0px;'>&nbsp;</td>" & chr(13) & _
					"	</tr>" & chr(13) & _
					"	<TR>" & chr(13) & _
					"		<TD COLSPAN='7' style='border-left:0px;border-right:0px;'>&nbsp;</td>" & chr(13) & _
					"	</tr>" & chr(13) & _
					"	<TR style='background:honeydew'>" & chr(13) & _
					"		<TD class='MTBE' COLSPAN='5' NOWRAP><p class='Cd'>" & _
					"TOTAL GERAL:</p></td>" & chr(13) & _
					"		<TD class='MTB'><p class='Cd'>" & formata_inteiro(qtde_total) & "</p></td>" & chr(13) & _
					"		<TD class='MTBD'><p class='Cd'>" & formata_moeda(vl_total_devolucao) & "</p></td>" & chr(13) & _
					"	</TR>" & chr(13)
			end if
		end if

  ' MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	if n_reg_total = 0 then
		x = cab_table & cab
		x = x & "	<TR>" & chr(13) & _
				"		<TD class='MT' colspan='7'><P class='ALERTA'>&nbsp;NENHUM PEDIDO ENCONTRADO&nbsp;</P></TD>" & chr(13) & _
				"	</TR>" & chr(13)
		end if

  ' FECHA TABELA DA ÚLTIMA LOJA
	x = x & "</TABLE>" & chr(13)
	
	Response.write x

	if r.State <> 0 then r.Close
	set r=nothing
	
end sub

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
	<title>LOJA</title>
	</head>



<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
window.status='Aguarde, executando a consulta ...';

function fRELConcluir( id_pedido ){
	window.status = "Aguarde ...";
	fREL.pedido_selecionado.value=id_pedido;
	fREL.action = "pedido.asp"
	fREL.submit(); 
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



<% else %>
<!-- ***************************************************** -->
<!-- **********  PÁGINA PARA EXIBIR RESULTADO   ********** -->
<!-- ***************************************************** -->
<body onload="window.status='Concluído';" link=#000000 alink=#000000 vlink=#000000>

<center>

<form id="fREL" name="fREL" method="post">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="pedido_selecionado" id="pedido_selecionado" value="">

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Devolução de Produtos</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>

<!-- FILTROS -->
<% 
	s_filtro = "<table width='649' cellPadding='0' CellSpacing='0' style='border-bottom:1px solid black' border='0'>"

	if (c_dt_devolucao_inicio <> "") Or (c_dt_devolucao_termino <> "") then
		s = ""
		s_aux = c_dt_devolucao_inicio
		if s_aux = "" then s_aux = "N.I."
		s = s & s_aux & " e "
		s_aux = c_dt_devolucao_termino
		if s_aux = "" then s_aux = "N.I."
		s = s & s_aux
		s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
				   "<p class='N'>Devolvido entre:&nbsp;</p></td><td valign='top' width='99%'>" & _
				   "<p class='N'>" & s & "</p></td></tr>"
		end if

	if (c_dt_cadastro_inicio <> "") Or (c_dt_cadastro_termino <> "") then
		s = ""
		s_aux = c_dt_cadastro_inicio
		if s_aux = "" then s_aux = "N.I."
		s = s & s_aux & " e "
		s_aux = c_dt_cadastro_termino
		if s_aux = "" then s_aux = "N.I."
		s = s & s_aux
		s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
				   "<p class='N'>Pedidos colocados entre:&nbsp;</p></td><td valign='top' width='99%'>" & _
				   "<p class='N'>" & s & "</p></td></tr>"
		end if

	if ckb_produto <> "" then 
		s_aux = c_fabricante
		if s_aux = "" then s_aux = "todos"
		s = "fabricante: " & s_aux
		s_aux = c_produto
		if s_aux = "" then s_aux = "todos"
		s = s & ",&nbsp;&nbsp;produto: " & s_aux
		s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
				   "<p class='N'>Somente pedidos que incluam:&nbsp;</p></td><td valign='top'>" & _
				   "<p class='N'>" & s & "</p></td></tr>"
		end if

	s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
			   "<p class='N'>Emissão:&nbsp;</p></td><td valign='top' width='99%'>" & _
			   "<p class='N'>" & formata_data_hora(Now) & "</p></td></tr>"

	s_filtro = s_filtro & "</table>"
	Response.Write s_filtro
%>

<!--  RELATÓRIO  -->
<br>
<% consulta_executa %>

<!-- ************   SEPARADOR   ************ -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>
<br>


<table class="notPrint" width="649" cellSpacing="0">
<tr>
	<td align="center"><a name="bVOLTA" id="bVOLTA" href="javascript:history.back()" title="volta para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
</tr>
</table>
</form>

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
