<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ======================================================
'	  R E L P R O D U T O S S E M P R E S E N C A . A S P
'     ======================================================
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

'	CONECTA COM O BANCO DE DADOS
	dim cn, rs, msg_erro
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	if Not operacao_permitida(OP_CEN_REL_PRODUTOS_PENDENTES, s_lista_operacoes_permitidas) then 
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	dim alerta
	dim s, rb_analise_credito, rb_etg_imediata, c_fabricante, c_produto, flag_ok, c_empresa,s_filtro

	alerta = ""

	rb_analise_credito = trim(Request.Form("rb_analise_credito"))
	rb_etg_imediata = trim(Request.Form("rb_etg_imediata"))
	c_empresa = Trim(Request.Form("c_empresa"))

	c_fabricante = retorna_so_digitos(Request.Form("c_fabricante"))
	if c_fabricante <> "" then 
		c_fabricante = normaliza_codigo(c_fabricante, TAM_MIN_FABRICANTE)
		s = "SELECT fabricante FROM t_FABRICANTE WHERE (fabricante='" & c_fabricante & "')"
		if rs.State <> 0 then rs.Close
		rs.open s, cn
		if rs.Eof then
			alerta = "FABRICANTE " & c_fabricante & " NÃO ESTÁ CADASTRADO."
			end if
		end if

	c_produto = Ucase(Trim(Request.Form("c_produto")))
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




' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________

' _____________________________________
' CONSULTA EXECUTA
'
sub consulta_executa
dim r
dim s, s_aux, s_sql, fabricante_a, produto_a, x, cab_table, cab
dim n_reg, n_reg_total, n_qtde, vl, msg_erro

'	LEMBRE-SE DE INCLUIR A RESTRIÇÃO "anulado_status=0" P/ SELECIONAR APENAS 
'	OS MOVIMENTOS VÁLIDOS, POIS "anulado_status<>0" INDICAM MOVIMENTOS QUE 
'	FORAM CANCELADOS E QUE ESTÃO NO BD APENAS POR QUESTÃO DE HISTÓRICO.
	s_sql = "SELECT" & _
				" t_ESTOQUE_MOVIMENTO.fabricante," & _
				" t_ESTOQUE_MOVIMENTO.produto," & _
				" descricao," & _
				" descricao_html," & _
				" t_ESTOQUE_MOVIMENTO.pedido," & _
				" qtde," & _
				" t_PEDIDO.data AS data_pedido," & _
				" t_PEDIDO__BASE.analise_credito," & _
				" t_PEDIDO.st_etg_imediata" & _
			" FROM t_ESTOQUE_MOVIMENTO" & _
				" LEFT JOIN t_PRODUTO" & _
					" ON ((t_ESTOQUE_MOVIMENTO.fabricante=t_PRODUTO.fabricante) AND (t_ESTOQUE_MOVIMENTO.produto=t_PRODUTO.produto))" & _
				" INNER JOIN t_PEDIDO" & _
					" ON (t_ESTOQUE_MOVIMENTO.pedido=t_PEDIDO.pedido)" & _
				" INNER JOIN t_PEDIDO AS t_PEDIDO__BASE ON (t_PEDIDO.pedido_base=t_PEDIDO__BASE.pedido)" & _
			" WHERE" & _
				" (anulado_status=0)" & _
				" AND (estoque='" & ID_ESTOQUE_SEM_PRESENCA & "')"	
	
	if c_fabricante <> "" then
		s_sql = s_sql & " AND (t_ESTOQUE_MOVIMENTO.fabricante='" & c_fabricante & "')" 
		end if

	if c_produto <> "" then
		if s_sql <> "" then s_sql = s_sql & " AND"
		s_sql = s_sql & " (t_ESTOQUE_MOVIMENTO.produto = '" & c_produto & "')"
		end if
	
	if (rb_analise_credito <> "TODOS") And (rb_analise_credito <> "") then
		if s_sql <> "" then s_sql = s_sql & " AND"
		s_sql = s_sql & " (t_PEDIDO__BASE.analise_credito = " & rb_analise_credito & ")"
		end if

	if (rb_etg_imediata <> "TODOS") And (rb_etg_imediata <> "") then
		if s_sql <> "" then s_sql = s_sql & " AND"
		s_sql = s_sql & " (t_PEDIDO.st_etg_imediata = " & rb_etg_imediata & ")"
		end if

    if c_empresa <> "" then
        if s_sql <> "" then s_sql = s_sql & " AND"
		s_sql = s_sql & " (t_PEDIDO.id_nfe_emitente = '" & c_empresa & "')"
		end if
		
	s_sql = s_sql & _
			" ORDER BY" & _
				" t_ESTOQUE_MOVIMENTO.fabricante," & _
				" t_ESTOQUE_MOVIMENTO.produto," & _
				" t_PEDIDO.data," & _
				" t_PEDIDO.hora," & _
				" t_PEDIDO.pedido"

  ' CABEÇALHO
	cab_table = "<TABLE class='Q' cellSpacing=0>" & chr(13)
	cab = "	<TR style='background: #FFF0E0' NOWRAP>" & chr(13) & _
		  "		<TD width='80' valign='bottom' NOWRAP class='MD MB'><P class='R'>PEDIDO</P></TD>" & chr(13) & _
		  "		<TD width='60' valign='bottom' NOWRAP class='MD MB'><P class='Rd' style='font-weight:bold;'>QTDE</P></TD>" & chr(13) & _
		  "		<TD width='85' valign='bottom' NOWRAP class='MD MB'><P class='Rd' style='font-weight:bold;'>DATA PEDIDO</P></TD>" & chr(13) & _
		  "		<TD width='200' valign='bottom' NOWRAP class='MD MB'><P class='R' style='font-weight:bold;'>ANÁLISE CRÉDITO</P></TD>" & chr(13) & _
		  "		<TD width='35' valign='bottom' NOWRAP class='MD MB'><P class='R' style='font-weight:bold;'>ENTR IMED</P></TD>" & chr(13) & _
		  "		<TD width='130' valign='bottom' NOWRAP class='MB'><P class='Rd' style='font-weight:bold;'>VL TOTAL PEDIDO</P></TD>" & chr(13) & _
		  "	</TR>" & chr(13)
	
	x = ""
	n_reg = 0
	n_reg_total = 0
	n_qtde = 0
	fabricante_a = "XXXXXX"
	produto_a = String(20, "X")
	
	set r = cn.execute(s_sql)
	do while Not r.Eof
	'	MUDOU PRODUTO?
		if (Trim("" & r("fabricante"))<>fabricante_a) Or (Trim("" & r("produto"))<>produto_a) then
			if n_reg_total > 0 then 
			  ' FECHA TABELA DO PRODUTO ANTERIOR
				x = x & _
					"	<TR NOWRAP style='background: #FFFFDD'>" & chr(13) & _
					"		<TD COLSPAN='2' NOWRAP><p class='Cd'>" & _
								"TOTAL:&nbsp;&nbsp;&nbsp;" & formata_inteiro(n_qtde) & "</p></TD>" & chr(13) & _
					"		<TD COLSPAN='4'>&nbsp;</td>" & chr(13) & _
					"	</TR>" & chr(13) & _
					"</TABLE>" & chr(13)
				Response.Write x
				x="<BR>" & chr(13)
				end if

		  ' MUDOU FABRICANTE?
			if Trim("" & r("fabricante"))<>fabricante_a then
				if n_reg_total > 0 then x = x & "<BR>" & chr(13)
				s = Trim("" & r("fabricante"))
				s_aux = x_fabricante(s)
				if (s<>"") And (s_aux<>"") then s = s & " - "
				s = s & s_aux
				if s <> "" then x = x & "<p class='STP' style='margin-bottom:4px;'>" & s & "</p>" & chr(13)
				end if
			
		  ' INICIA NOVA TABELA P/ O NOVO PRODUTO
			x = x & _
				cab_table & _
				"	<TR>" & chr(13) & _
				"		<TD COLSPAN='6' align='center' valign='bottom' class='MB' style='background:azure;'>" & _
				"<P class='F'>" & r("produto") & " - " & _
				produto_formata_descricao_em_html(Trim("" & r("descricao_html"))) & "</P></TD>" & chr(13) & _
				"	</TR>" & chr(13)
			
			x = x & cab
			n_reg = 0
			n_qtde = 0
			fabricante_a = Trim("" & r("fabricante"))
			produto_a = Trim("" & r("produto"))
			end if
			
	  ' CONTAGEM
		n_reg = n_reg + 1
		n_reg_total = n_reg_total + 1

		x = x & "	<TR NOWRAP>" & chr(13)

	 '> PEDIDO
		x = x & "		<TD class='MDB'><P class='C'><a href='javascript:fRELConcluir(" & _
				chr(34) & Trim("" & r("pedido")) & chr(34) & ")' title='clique para consultar o pedido'>" & _
				Trim("" & r("pedido")) & "</a></P></TD>" & chr(13)

	 '> QTDE
		x = x & "		<TD class='MDB'><P class='Cd'>&nbsp;" & formata_inteiro(r("qtde")) & "</P></TD>" & chr(13)

	 '> DATA DO PEDIDO
		x = x & "		<TD class='MDB' NOWRAP><P class='Cc' NOWRAP>&nbsp;" & formata_data(r("data_pedido")) & "</P></TD>" & chr(13)

	 '> STATUS DA ANÁLISE DE CRÉDITO
		x = x & "		<TD class='MDB' NOWRAP><P class='C' NOWRAP style='color:" & x_analise_credito_cor(r("analise_credito")) & ";'>&nbsp;" & x_analise_credito(r("analise_credito")) & "</P></TD>" & chr(13)

	 '> ENTREGA IMEDIATA
		s = ""
		if Trim("" & r("st_etg_imediata")) = Cstr(COD_ETG_IMEDIATA_NAO) then
			s = "Não"
		elseif Trim("" & r("st_etg_imediata")) = Cstr(COD_ETG_IMEDIATA_SIM) then
			s = "Sim"
			end if
		if s = "" then s = "&nbsp;"
		x = x & "		<TD class='MDB' NOWRAP><P class='C' NOWRAP>&nbsp;" & s & "</P></TD>" & chr(13)

	 '> VALOR TOTAL DO PEDIDO
		if Not calcula_valor_total_pedido(r("pedido"), vl, msg_erro) then vl=0
		x = x & "		<TD class='MB' NOWRAP><P class='Cd'>&nbsp;" & formata_moeda(vl) & "</P></TD>" & chr(13)

		n_qtde = n_qtde + r("qtde")
		
		x = x & "	</TR>" & chr(13)
			
		if (n_reg_total mod 100) = 0 then
			Response.Write x
			x = ""
			end if
			
		r.movenext
		loop
		
  ' MOSTRA TOTAL DO ÚLTIMO PRODUTO
	if n_reg <> 0 then 
		x = x & _
			"	<TR NOWRAP style='background: #FFFFDD'>" & chr(13) & _
			"		<TD COLSPAN='2' NOWRAP><P class='Cd'>" & "TOTAL:&nbsp;&nbsp;&nbsp;" & formata_inteiro(n_qtde) & "</P></TD>" & chr(13) & _
			"		<TD COLSPAN='4'>&nbsp;</td>" & chr(13) & _
			"	</TR>" & chr(13)
		end if

  ' MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	if n_reg_total = 0 then
		x = ""		
		x = x & _
			cab_table & _
			cab & _
            "	<br>" & chr(13) & _
			"	<TR NOWRAP>" & chr(13) & _
			"		<TD colspan='6'><P class='ALERTA'>&nbsp;NÃO HÁ PRODUTOS VENDIDOS SEM PRESENÇA NO ESTOQUE&nbsp;</P></TD>" & chr(13) & _
			"	</TR>" & chr(13)
		end if

  ' FECHA TABELA DO ÚLTIMO PRODUTO
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
	<title>CENTRAL</title>
	</head>



<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
window.status='Aguarde, executando a consulta ...';

function fRELConcluir( id_pedido ) {
	fREL.action = "pedido.asp";
	fREL.pedido_selecionado.value = id_pedido;
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



<% else %>
<!-- ***************************************************** -->
<!-- **********  PÁGINA PARA EXIBIR RESULTADO   ********** -->
<!-- ***************************************************** -->
<body onload="window.status='Concluído';" link=#000000 alink=#000000 vlink=#000000>

<center>

<form id="fREL" name="fREL" method="post">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="pedido_selecionado" id="pedido_selecionado" value=''>
<input type="hidden" name="c_fabricante" id="c_fabricante" value="<%=c_fabricante%>">
<input type="hidden" name="c_produto" id="c_produto" value="<%=c_produto%>">
<input type="hidden" name="rb_analise_credito" id="rb_analise_credito" value="<%=rb_analise_credito%>">


<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Produtos Pendentes</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span>
     </td>
</tr>
</table>

<!-- FILTROS -->
<%
    s_filtro = "<table width='649' cellpadding='0' cellspacing='0' style='border-bottom:1px solid black' border='0'>"

     s = c_empresa
	if s = "" then 
		s = "todas"
	else
		s =  obtem_apelido_empresa_NFe_emitente(c_empresa)
    end if
        s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>Empresa:&nbsp;</span></td><td align='left' valign='top'>" & _
			   "<span class='N'>" & s & "</span></td></tr>"    

    s = c_fabricante
	if s = "" then s = "todos"
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>Fabricante:&nbsp;</span></td><td align='left' valign='top'>" & _
			   "<span class='N'>" & s & "</span></td></tr>"
	
	s = c_produto
	if s = "" then s = "todos"
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>Produto:&nbsp;</span></td><td align='left' valign='top'>" & _
			   "<span class='N'>" & s & "</span></td></tr>"

    s = rb_analise_credito
    if s <> "TODOS" then s = descricao_analise_credito(s)

	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>Status da Análise de Crédito:&nbsp;</span></td><td align='left' valign='top'>" & _
			   "<span class='N'>" & LCase(s) & "</span></td></tr>"

    s = rb_etg_imediata    
    if s = "1" then s = "não"
    if s = "2" then s = "sim"	
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>Entrega Imediata:&nbsp;</span></td><td align='left' valign='top'>" & _
			   "<span class='N'>" & LCase(s) & "</span></td></tr>"


    s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>Emissão:&nbsp;</span></td><td align='left' valign='top' width='99%'>" & _
			   "<span class='N'>" & formata_data_hora(Now) & "</span></td></tr>"

    s_filtro = s_filtro & "</table>"
	Response.Write s_filtro
%>


<!--  RELATÓRIO  -->
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
