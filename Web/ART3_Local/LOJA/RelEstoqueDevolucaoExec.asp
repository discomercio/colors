<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=False %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ======================================================
'	  RelEstoqueDevolucaoExec.asp
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
	if Not operacao_permitida(OP_LJA_REL_PRODUTOS_ESTOQUE_DEVOLUCAO, s_lista_operacoes_permitidas) then 
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	dim alerta
	dim s, s_aux, s_filtro, flag_ok, s_filtro_operacao
	dim c_loja, c_fabricante, c_produto, c_pedido
	dim c_vendedor, c_indicador, c_captador, s_nome_vendedor

	alerta = ""

	c_loja = Trim(Request.Form("c_loja"))
	c_fabricante = retorna_so_digitos(Trim(Request.Form("c_fabricante")))
	c_produto = Ucase(Trim(Request.Form("c_produto")))
	c_pedido = Ucase(Trim(Request.Form("c_pedido")))
	c_vendedor = Ucase(Trim(Request.Form("c_vendedor")))
	c_indicador = Ucase(Trim(Request.Form("c_indicador")))
	c_captador = Ucase(Trim(Request.Form("c_captador")))

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


'	Pedido cadastrado?
	if alerta = "" then
		if c_pedido <> "" then
			s = "SELECT pedido FROM t_PEDIDO WHERE (pedido='" & c_pedido & "')"
			if rs.State <> 0 then rs.Close
			rs.open s, cn
			if rs.Eof then 
				alerta=texto_add_br(alerta)
				alerta = "PEDIDO " & c_pedido & " NÃO ESTÁ CADASTRADO."
				end if
			end if
		end if
	
'	Vendedor
	if alerta = "" then
		s_nome_vendedor = ""
		if c_vendedor <> "" then
			s = "SELECT nome_iniciais_em_maiusculas FROM t_USUARIO WHERE (usuario='" & c_vendedor & "')"
			if rs.State <> 0 then rs.Close
			rs.open s, cn
			if rs.Eof then 
				alerta=texto_add_br(alerta)
				alerta = "VENDEDOR " & c_vendedor & " NÃO ESTÁ CADASTRADO."
			else
				s_nome_vendedor = Trim("" & rs("nome_iniciais_em_maiusculas"))
				end if
			end if
		end if
		
		
	
	


' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________

function monta_link_pedido(byval id_pedido)
dim strLink
	monta_link_pedido = ""
	id_pedido = Trim("" & id_pedido)
	if id_pedido = "" then exit function
	strLink = "<a href='javascript:fRELConcluir(" & _
				chr(34) & id_pedido & chr(34) & _
				")' title='clique para consultar o pedido'>" & _
				id_pedido & "</a>"
	monta_link_pedido=strLink
end function


' _____________________________________
' CONSULTA EXECUTA
'
sub consulta_executa
dim r
dim cab, cab_table
dim n_reg, n_reg_total
dim x, s, s_aux, s_sql, msg_erro
dim s_where, s_where_loja
dim produto_a
dim intNumProdutos, intQtdeTotal, intQtdeSubTotal

'	CRITÉRIOS COMUNS
	s_where = ""

'	FILTROS
'	~~~~~~~
'	FABRICANTE
	if c_fabricante <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_ESTOQUE_MOVIMENTO.fabricante = '" & c_fabricante & "')"
		end if

'	PRODUTO
	if c_produto <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_ESTOQUE_MOVIMENTO.produto = '" & c_produto & "')"
		end if

'	PEDIDO
	if c_pedido <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_ESTOQUE_MOVIMENTO.pedido = '" & c_pedido & "')"
		end if

'	VENDEDOR
	if c_vendedor <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_PEDIDO.vendedor = '" & c_vendedor & "')"
		end if

'	INDICADOR
	if c_indicador <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_PEDIDO__BASE.indicador = '" & c_indicador & "')"
		end if

'	CAPTADOR
	if c_captador <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_ORCAMENTISTA_E_INDICADOR.captador = '" & c_captador & "')"
		end if

'	ESTA LOJA
	s_where_loja = " (t_PEDIDO.numero_loja = " & c_loja & ")"
	if s_where <> "" then s_where = s_where & " AND"
	s_where = s_where & " (" & s_where_loja & ")"

'	CRITÉRIOS FIXOS
	if s_where <> "" then s_where = s_where & " AND"
	s_where = s_where & " (anulado_status=0)"
	
	if s_where <> "" then s_where = s_where & " AND"
	s_where = s_where & " (estoque='" & ID_ESTOQUE_DEVOLUCAO & "')"
	
'	MONTA A CONSULTA
	if s_where <> "" then s_where = " WHERE" & s_where
	
	s_sql = "SELECT" & _
				" t_PEDIDO_ITEM_DEVOLVIDO.devolucao_data," & _
				" t_ESTOQUE_MOVIMENTO.loja," & _
				" CONVERT(smallint,t_ESTOQUE_MOVIMENTO.loja) AS numero_loja," & _
				" t_ESTOQUE_MOVIMENTO.fabricante," & _
				" t_ESTOQUE_MOVIMENTO.produto," & _
				" t_PRODUTO.descricao," & _
				" t_PRODUTO.descricao_html," & _
				" t_PEDIDO.data," & _
				" t_ESTOQUE_MOVIMENTO.pedido," & _
				" t_PEDIDO.obs_2," & _
				" t_PEDIDO__BASE.vendedor," & _
				" t_PEDIDO__BASE.indicador," & _
				" t_CLIENTE.nome_iniciais_em_maiusculas AS nome_cliente," & _
				" t_PEDIDO_ITEM_DEVOLVIDO.motivo," & _
				" Coalesce(Sum(t_ESTOQUE_MOVIMENTO.qtde),0) AS saldo" & _
			" FROM t_ESTOQUE_MOVIMENTO" & _
				" LEFT JOIN t_PRODUTO ON ((t_ESTOQUE_MOVIMENTO.fabricante=t_PRODUTO.fabricante) AND (t_ESTOQUE_MOVIMENTO.produto=t_PRODUTO.produto))" & _
				" INNER JOIN t_PEDIDO ON (t_ESTOQUE_MOVIMENTO.pedido=t_PEDIDO.pedido)" & _
				" INNER JOIN t_PEDIDO AS t_PEDIDO__BASE" & _
					" ON (t_PEDIDO.pedido_base=t_PEDIDO__BASE.pedido)" & _
				" INNER JOIN t_ESTOQUE ON (t_ESTOQUE.id_estoque=t_ESTOQUE_MOVIMENTO.id_estoque)" & _
				" LEFT JOIN t_PEDIDO_ITEM_DEVOLVIDO ON (t_PEDIDO_ITEM_DEVOLVIDO.id=t_ESTOQUE.devolucao_id_item_devolvido)" & _
				" LEFT JOIN t_ORCAMENTISTA_E_INDICADOR" & _
					" ON (t_PEDIDO__BASE.indicador=t_ORCAMENTISTA_E_INDICADOR.apelido)" & _
				" INNER JOIN t_CLIENTE ON (t_PEDIDO.id_cliente=t_CLIENTE.id)" & _
			s_where & _
			" GROUP BY" & _
				" t_ESTOQUE_MOVIMENTO.fabricante," & _
				" t_ESTOQUE_MOVIMENTO.produto," & _
				" t_PRODUTO.descricao," & _
				" t_PRODUTO.descricao_html," & _
				" t_PEDIDO_ITEM_DEVOLVIDO.devolucao_data," & _
				" t_ESTOQUE_MOVIMENTO.loja," & _
				" t_PEDIDO.data," & _
				" t_ESTOQUE_MOVIMENTO.pedido," & _
				" t_PEDIDO.obs_2," & _
				" t_PEDIDO__BASE.vendedor," & _
				" t_PEDIDO__BASE.indicador," & _
				" t_CLIENTE.nome_iniciais_em_maiusculas," & _
				" t_PEDIDO_ITEM_DEVOLVIDO.motivo" & _
			" ORDER BY" & _
				" t_ESTOQUE_MOVIMENTO.fabricante," & _
				" t_ESTOQUE_MOVIMENTO.produto," & _
				" t_PRODUTO.descricao," & _
				" t_PRODUTO.descricao_html," & _
				" t_PEDIDO_ITEM_DEVOLVIDO.devolucao_data," & _
				" numero_loja," & _
				" t_PEDIDO.data," & _
				" t_ESTOQUE_MOVIMENTO.pedido," & _
				" t_PEDIDO.obs_2"
	
  ' CABEÇALHO
	cab_table = "<TABLE CellSpacing=0 CellPadding=0>" & chr(13)
	cab = "	<TR style='background:azure' NOWRAP>" & chr(13) & _
		  "		<TD class='MDTE tdDataDevolucao' style='vertical-align:bottom'><P class='Rc'>Data Devol</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdLoja' style='vertical-align:bottom'><P class='Rc'>Loja</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdPedido' style='vertical-align:bottom'><P class='R'>Pedido</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdObs2' style='vertical-align:bottom'><P class='R'>Obs II</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdQtd' style='vertical-align:bottom'><P class='Rd' style='font-weight:bold;'>Qtde</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdCliente' style='vertical-align:bottom'><P class='R'>Cliente</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdIndicador' style='vertical-align:bottom'><P class='R'>Parceiro</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdVendedor' style='vertical-align:bottom'><P class='R'>Vendedor</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdMotivo' style='vertical-align:bottom'><P class='R'>Motivo</P></TD>" & chr(13) & _
		  "	</TR>" & chr(13)

	produto_a = "XXXXXXXXXXXXXX"
	intNumProdutos = 0
	
	n_reg = 0
	n_reg_total = 0

	x = ""

	If Not cria_recordset_otimista(r, msg_erro) then 
		Response.Write msg_erro
		exit sub
		end if

	r.open s_sql, cn
	do while Not r.Eof
	
	'	MUDOU DE PRODUTO?
		s = "|" & Trim("" & r("fabricante")) & "|" & Trim("" & r("produto")) & "|"
		if produto_a <> s then
			produto_a = s
			intNumProdutos = intNumProdutos + 1
		'	FECHA TABELA DO PRODUTO ANTERIOR
			if n_reg_total > 0 then
				x = x & "	<TR NOWRAP style='background: #FFFFDD'>" & chr(13) & _
						"		<TD class='MTBE' colspan='4' NOWRAP><p class='Cd'>TOTAL:</p></td>" & chr(13) & _
						"		<TD class='MTB' NOWRAP><p class='Cd'>" & formata_inteiro(intQtdeSubTotal) & "</p></td>" & chr(13) & _
						"		<TD class='MTBD' colspan='4'>&nbsp;</TD>" & chr(13) & _
						"	</TR>" & chr(13) & _
						"</TABLE>" & chr(13)
				Response.Write x
				x="<BR>"
				end if
			
			intQtdeSubTotal = 0
			n_reg = 0
			if n_reg_total > 0 then x = x & "<BR>"
		'	DESCRIÇÃO DO PRODUTO
			s_aux = Trim("" & r("fabricante"))
			if s_aux <> "" then s_aux = "(" & s_aux & ") "
			s = Trim("" & r("produto"))
			s = s_aux & s
			s_aux = produto_formata_descricao_em_html(Trim("" & r("descricao_html")))
			if s_aux <> "" then s_aux = " - " & s_aux
			s = s & s_aux
			if s = "" then s = "&nbsp;"
			x = x & cab_table
			if s <> "" then x = x & "<TR><TD class='MDTE' COLSPAN='9' valign='bottom' class='MB' style='background:azure;'><p class='N'>&nbsp;" & s & "</p></td></tr>" & chr(13)
			x = x & cab
			end if

	 ' CONTAGEM
		n_reg = n_reg + 1
		n_reg_total = n_reg_total + 1

		x = x & "	<TR NOWRAP>" & chr(13)

	'> DATA DA DEVOLUÇÃO
		s = formata_data(r("devolucao_data"))
		x = x & "		<TD class='MDTE tdDataDevolucao'><P class='Cnc'>" & s & "</P></TD>" & chr(13)

	'> LOJA
		s = Trim("" & r("loja"))
		x = x & "		<TD class='MTD tdLoja'><P class='Cnc'>" & s & "</P></TD>" & chr(13)

	'> PEDIDO
		s = monta_link_pedido(Trim("" & r("pedido")))
		x = x & "		<TD class='MTD tdPedido'><P class='Cn'>" & s & "</P></TD>" & chr(13)

	'> OBS II
		s = Trim("" & r("obs_2"))
		if s = "" then s = "&nbsp;"
		x = x & "		<TD class='MTD tdObs2'><P class='Cn'>" & s & "</P></TD>" & chr(13)

	'> QUANTIDADE
		s = formata_inteiro(r("saldo"))
		x = x & "		<TD class='MTD tdQtd'><P class='Cd'>" & s & "</P></TD>" & chr(13)

		intQtdeSubTotal = intQtdeSubTotal + r("saldo")
		intQtdeTotal = intQtdeTotal + r("saldo")
		
	'> CLIENTE
		s = Trim("" & r("nome_cliente"))
		if s = "" then s = "&nbsp;"
		x = x & "		<TD class='MTD tdCliente'><P class='Cn'>" & s & "</P></TD>" & chr(13)

	'> INDICADOR
		s = iniciais_em_maiusculas(Trim("" & r("indicador")))
		if s = "" then s = "&nbsp;"
		x = x & "		<TD class='MTD tdIndicador'><P class='Cn'>" & s & "</P></TD>" & chr(13)

	'> VENDEDOR
		s = iniciais_em_maiusculas(Trim("" & r("vendedor")))
		if s = "" then s = "&nbsp;"
		x = x & "		<TD class='MTD tdVendedor'><P class='Cn'>" & s & "</P></TD>" & chr(13)

	'> MOTIVO
		s = Trim("" & r("motivo"))
		if s = "" then s = "&nbsp;"
		x = x & "		<TD class='MTD tdMotivo'><P class='Cn'>" & s & "</P></TD>" & chr(13)

		x = x & "	</TR>" & chr(13)

		if (n_reg mod 100) = 0 then
			Response.Write x
			x = ""
			end if
			
		r.MoveNext
		loop

'	MOSTRA TOTAL DO ÚLTIMO PRODUTO
	if n_reg <> 0 then
		x = x & "	<TR NOWRAP style='background: #FFFFDD'>" & chr(13) & _
				"		<TD class='MTBE' colspan='4' NOWRAP><p class='Cd'>TOTAL:</p></td>" & chr(13) & _
				"		<TD class='MTB' NOWRAP><p class='Cd'>" & formata_inteiro(intQtdeSubTotal) & "</p></td>" & chr(13) & _
				"		<TD class='MTBD' colspan='4'>&nbsp;</TD>" & chr(13) & _
				"	</TR>" & chr(13)
				
	'	TOTAL GERAL
		if intNumProdutos > 1 then
			x = x & "<TR><TD COLSPAN='9' style='border-left:0px;border-right:0px;'>&nbsp;</td></tr>" & chr(13) & _
					"<TR><TD COLSPAN='9' style='border-left:0px;border-right:0px;'>&nbsp;</td></tr>" & chr(13) & _
					"	<TR NOWRAP style='background:honeydew'>" & chr(13) & _
					"		<TD class='MTBE' colspan='4' NOWRAP><p class='Cd'>TOTAL GERAL:</p></td>" & chr(13) & _
					"		<TD class='MTB' NOWRAP><p class='Cd'>" & formata_inteiro(intQtdeTotal) & "</p></td>" & chr(13) & _
					"		<TD class='MTBD' colspan='4'>&nbsp;</TD>" & chr(13) & _
					"	</TR>" & chr(13)
			
			end if
		
		end if
	
  ' MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	if n_reg_total = 0 then
		x = cab_table & cab
		x = x & "	<TR NOWRAP>" & chr(13) & _
				"		<TD class='MT' colspan='9'><P class='ALERTA'>&nbsp;NENHUM REGISTRO ENCONTRADO&nbsp;</P></TD>" & chr(13) & _
				"	</TR>" & chr(13)
		end if

  ' FECHA TABELA
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

<style type="text/css">
.tdDataDevolucao{
	vertical-align: top;
	width: 65px;
	}
.tdLoja{
	vertical-align: top;
	width: 30px;
	}
.tdPedido{
	vertical-align: top;
	font-weight: bold;
	width: 65px;
	}
.tdObs2{
	vertical-align: top;
	width: 60px;
	}
.tdQtd{
	vertical-align: top;
	width: 35px;
	}
.tdCliente{
	vertical-align: top;
	width: 130px;
	}
.tdIndicador{
	vertical-align: top;
	width: 65px;
	}
.tdVendedor{
	vertical-align: top;
	width: 65px;
	}
.tdMotivo{
	vertical-align: top;
	width: 160px;
	}
</style>


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
	<td align="CENTER"><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()"><img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
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
<input type="hidden" name="c_fabricante" id="c_fabricante" value="<%=c_fabricante%>">
<input type="hidden" name="c_produto" id="c_produto" value="<%=c_produto%>">
<input type="hidden" name="c_pedido" id="c_pedido" value="<%=c_pedido%>">
<input type="hidden" name="c_vendedor" id="c_vendedor" value="<%=c_vendedor%>">
<input type="hidden" name="c_indicador" id="c_indicador" value="<%=c_indicador%>">
<input type="hidden" name="c_captador" id="c_captador" value="<%=c_captador%>">
<input type="hidden" name="c_loja" id="c_loja" value="<%=c_loja%>">

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="685" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Produtos no Estoque de Devolução</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>

<!-- FILTROS -->
<% 
	s_filtro = "<table width='685' cellPadding='0' CellSpacing='0' style='border-bottom:1px solid black' border='0'>" & chr(13)
   
	s = c_fabricante
	if s <> "" then
		s_aux = x_fabricante(s)
		if (s<>"") And (s_aux<>"") then s = s & " - "
		s = s & s_aux
		end if
	if s = "" then s = "todos"
	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' NOWRAP>" & _
				"<p class='N'>Fabricante:&nbsp;</p></td><td valign='top'>" & _
				"<p class='N'>" & s & "</p></td>" & chr(13) & _
				"	</tr>" & chr(13)
	
	s = c_produto
	if s <> "" then
		s_aux = produto_formata_descricao_em_html(produto_descricao_html(c_fabricante, s))
		if (s<>"") And (s_aux<>"") then s = s & " - "
		s = s & s_aux
		end if
	if s = "" then s = "todos"
	s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='baseline' NOWRAP>" & _
					"<p class='N'>Produto:&nbsp;</p></td><td valign='baseline'>" & _
					"<p class='N'>" & s & "</p></td>" & chr(13) & _
					"	</tr>" & chr(13)

	s = c_pedido
	if s = "" then s = "todos"
	s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='top' NOWRAP>" & _
					"<p class='N'>Pedido:&nbsp;</p></td><td valign='top'>" & _
					"<p class='N'>" & s & "</p></td>" & chr(13) & _
					"	</tr>" & chr(13)

	s = c_vendedor
	if s = "" then 
		s = "todos"
	else
		if (s_nome_vendedor <> "") And (s_nome_vendedor <> c_vendedor) then s = s & " (" & s_nome_vendedor & ")"
		end if
	s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
			   "<p class='N'>Vendedor:&nbsp;</p></td><td valign='top'>" & _
			   "<p class='N'>" & s & "</p></td></tr>"

	s = c_indicador
	if s = "" then 
		s = "todos"
	else
		s = s & " (" & x_orcamentista_e_indicador(c_indicador) & ")"
		end if
	s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
			   "<p class='N'>Indicador:&nbsp;</p></td><td valign='top'>" & _
			   "<p class='N'>" & s & "</p></td></tr>"

	s = c_captador
	if s = "" then 
		s = "todos"
	else
		s = s & " (" & x_usuario(c_captador) & ")"
		end if
	s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
			   "<p class='N'>Captador:&nbsp;</p></td><td valign='top'>" & _
			   "<p class='N'>" & s & "</p></td></tr>"
	
	s = c_loja
	if s = "" then s = "todas"
	s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
			   "<p class='N'>Loja:&nbsp;</p></td><td valign='top'>" & _
			   "<p class='N'>" & s & "</p></td></tr>"

	s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='top' NOWRAP>" & _
					"<p class='N'>Emissão:&nbsp;</p></td><td valign='top' width='99%'>" & _
					"<p class='N'>" & formata_data_hora(Now) & "</p></td>" & chr(13) & _
					"	</tr>" & chr(13)
	
	s_filtro = s_filtro & "</table>" & chr(13)
	Response.Write s_filtro
%>

<!--  RELATÓRIO  -->
<br>
<% consulta_executa %>

<!-- ************   SEPARADOR   ************ -->
<table width="685" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>
<br>


<table class="notPrint" width="685" cellSpacing="0">
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
