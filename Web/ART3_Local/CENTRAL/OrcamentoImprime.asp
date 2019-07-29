<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ===========================================
'	  O R C A M E N T O I M P R I M E . A S P
'     ===========================================
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

	dim s, usuario, orcamento_selecionado
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

	orcamento_selecionado = ucase(Trim(request("orcamento_selecionado")))
	if (orcamento_selecionado = "") then Response.Redirect("aviso.asp?id=" & ERR_ORCAMENTO_NAO_ESPECIFICADO)
	s = normaliza_num_orcamento(orcamento_selecionado)
	if s <> "" then orcamento_selecionado = s
	
	dim i, n, s_fabricante, s_produto, s_descricao, s_descricao_html, s_obs, s_qtde, s_preco_lista, s_desc_dado, s_vl_unitario
	dim s_vl_TotalItem, m_TotalItem, m_TotalItemComRA, m_TotalDestePedido, m_TotalDestePedidoComRA
	dim s_preco_NF, m_total_NF, m_total_RA
	
'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	dim r_orcamento, v_item, alerta, msg_erro
	alerta=""
	if Not le_orcamento(orcamento_selecionado, r_orcamento, msg_erro) then 
		alerta = msg_erro
	else
		if Not le_orcamento_item(orcamento_selecionado, v_item, msg_erro) then alerta = msg_erro
		end if

	dim r_pedido
	if alerta = "" then
		if r_orcamento.st_orc_virou_pedido = 1 then
			if Not le_pedido(r_orcamento.pedido, r_pedido, msg_erro) then alerta = msg_erro
			end if
		end if

	dim s_aux, s2, s3, r_loja, r_cliente, s_script

	if alerta = "" then
		set r_loja = New cl_LOJA
		if Not x_loja_bd(r_orcamento.loja, r_loja) then Response.Redirect("aviso.asp?id=" & ERR_LOJA_NAO_CADASTRADA)
		
		set r_cliente = New cl_CLIENTE
		if Not x_cliente_bd(r_orcamento.id_cliente, r_cliente) then Response.Redirect("aviso.asp?id=" & ERR_CLIENTE_NAO_CADASTRADO)
		end if

	if alerta = "" then
		if Not orcamento_calcula_total_NF_e_RA(orcamento_selecionado, m_total_NF, m_total_RA, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
		end if

	dim strTextoIndicador
	dim r_orcamentista_e_indicador
	if alerta = "" then
		call le_orcamentista_e_indicador(r_orcamento.orcamentista, r_orcamentista_e_indicador, msg_erro)
		end if

	dim blnTemRA
	blnTemRA = False
	if alerta = "" then
		for i=Lbound(v_item) to Ubound(v_item)
			if Trim("" & v_item(i).produto) <> "" then
				if v_item(i).preco_NF <> v_item(i).preco_venda then
					blnTemRA = True
					exit for
					end if
				end if
			next
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
<!-- #include file = "../global/printerx.txt"    -->
	<title>CENTRAL</title>
</head>


<script src="<%=URL_FILE__CONST_JS%>" Language="JavaScript" Type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
var ja_carregou=false;

function fPEDConcluir(s_pedido){
	window.status = "Aguarde ...";
	fPED.pedido_selecionado.value=s_pedido;
	fPED.submit(); 
}

function impressora_carregada() {
var s;
	if (!( "object" == typeof(printer) && "string" == typeof(printer.module_id))) {
		s = "Componente necessário para impressão do orçamento não foi carregado corretamente!!";
		alert(s);
		return false;
		}
	return true;
}

function fORCImpressora ( f ) {
	if (!ja_carregou) return;
	if (!impressora_carregada()) return;
	printer.Initialize;
	if (printer.printing) printer.EndDoc();
	printer.seleciona_impressora();
}

function fORCMargens( f ) {
	if (!ja_carregou) return;
	if (!impressora_carregada()) return;
	printer.configura_margens();
}
</script>

<% 
s_script = "<script language='JavaScript'>" & chr(13) & _
"function fORCImprime( f ) {" & chr(13) & _
"var cx, cy, cw, offx, offy, margemx, margemy, campo;" & chr(13) & _
	"if (!ja_carregou) return;" & chr(13) & _
	"if (!impressora_carregada(printer)) return;" & chr(13) & _
	"printer.Initialize;" & chr(13) & _
	"if (printer.printing) printer.EndDoc();" & chr(13) & _
	"printer.landscape=false;" & chr(13) & _
	"printer.setpapersizeletter();" & chr(13) & _
	"printer.job_title='ORÇAMENTO " & orcamento_selecionado & "';" & chr(13) & _
	"printer.brushstyle='bsClear';" & chr(13) & _
	"printer.fontcolor=0;" & chr(13) & _
	"printer.fontname='Arial';" & chr(13) & _
	"printer.fontsize=9;" & chr(13) & _
	"printer.fontnormal=true;" & chr(13) & _
	"printer.fontbold=true;" & chr(13) & _
	"printer.penmode='pmBlack';" & chr(13) & _
	"printer.penstyle='psSolid';" & chr(13) & _
	"printer.pencolor=0;" & chr(13) & _
	"printer.penwidth=2;" & chr(13) & _
	"printer.BeginDoc();" & chr(13) & _
	"margemx=1;" & chr(13) & _
	"margemy=1;" & chr(13) & _
	"offx=2;" & chr(13) & _
	"offy=2;" & chr(13) 

'>	Nº ORÇAMENTO
	s = orcamento_selecionado
	s_script = s_script & _
	"printer.fontsize=12;" & chr(13) & _
	"cx=margemx+170+(30-printer.texto_largura('" & s & "'))/2;" & chr(13) & _
	"cy=margemy+3+1;" & chr(13) & _
	"printer.imprime(cx, cy, '" & s & "');" & chr(13)

'	RÓTULO "ORÇAMENTO"
	s = "ORÇAMENTO"
	s_script = s_script & _
	"printer.fontsize=10;" & chr(13) & _
	"cx=margemx+170+(30-printer.texto_largura('" & s & "'))/2;" & chr(13) & _
	"cy=cy+printer.texto_altura('X')+1;" & chr(13) & _
	"printer.imprime(cx, cy, '" & s & "');" & chr(13)
	
'>	DATA
	s = formata_data(r_orcamento.data)
	s_script = s_script & _
	"printer.fontsize=12;" & chr(13) & _
	"cx=margemx+170+(30-printer.texto_largura('" & s & "'))/2;" & chr(13) & _
	"cy=margemy+27-offy-printer.texto_altura('X');" & chr(13) & _
	"printer.imprime(cx, cy, '" & s & "');" & chr(13)
	
'>	LOJA - NOME
	s = ""
	with r_loja
		if Trim(.razao_social) <> "" then
			s = iniciais_em_maiusculas(Trim(.razao_social))
		else
			s = iniciais_em_maiusculas(Trim(.nome))
			end if
		end with
	s_script = s_script & _
	"printer.fontsize=9;" & chr(13) & _
	"cx=margemx+offx;" & chr(13) & _
	"cy=margemy+40-offy-printer.texto_altura('X');" & chr(13) & _
	"cw=margemx+126-cx-1;" & chr(13) & _
	"printer.imprime_campo(cx, cy, cw, '" & s & "');" & chr(13)
	
'	LOJA - TELEFONE
	s = ""
	with r_loja
		if Trim(.telefone) <> "" then
			s = telefone_formata(Trim(.telefone))
			s_aux=Trim(.ddd)
			if s_aux<>"" then s = "(" & s_aux & ") " & s
			end if
		end with
	s_script = s_script & _
	"cx=margemx+offx+126;" & chr(13) & _
	"cw=margemx+170-cx-1;" & chr(13) & _
	"printer.imprime_campo(cx, cy, cw, '" & s & "');" & chr(13)

'	VENDEDOR
	s_script = s_script & _
	"cx=margemx+offx+170;" & chr(13) & _
	"cw=margemx+200-cx-1;" & chr(13) & _
	"printer.imprime_campo(cx, cy, cw, '" & r_orcamento.vendedor & "');" & chr(13)

'>	CLIENTE - NOME
	s = iniciais_em_maiusculas(r_cliente.nome)
	s_script = s_script & _
	"cx=margemx+offx;" & chr(13) & _
	"cy=margemy+53-offy-printer.texto_altura('X');" & chr(13) & _
	"cw=margemx+148-cx-1;" & chr(13) & _
	"printer.imprime_campo(cx, cy, cw, '" & s & "');" & chr(13)

'	CLIENTE - CNPJ/CPF
	s = cnpj_cpf_formata(r_cliente.cnpj_cpf)
	s_script = s_script & _
	"cx=margemx+offx+148;" & chr(13) & _
	"cw=margemx+200-cx-1;" & chr(13) & _
	"printer.imprime_campo(cx, cy, cw, '" & s & "');" & chr(13)

'>	CLIENTE - ENDEREÇO
	s = iniciais_em_maiusculas(r_cliente.endereco)
	if r_cliente.endereco_numero <> "" then s = s & ", " & r_cliente.endereco_numero
	if r_cliente.endereco_complemento <> "" then s = s & " " & r_cliente.endereco_complemento
	s_script = s_script & _
	"cx=margemx+offx;" & chr(13) & _
	"cy=margemy+66-offy-printer.texto_altura('X');" & chr(13) & _
	"cw=margemx+126-cx-1;" & chr(13) & _
	"printer.imprime_campo(cx, cy, cw, '" & s & "');" & chr(13)

'	CLIENTE - BAIRRO
	s = iniciais_em_maiusculas(r_cliente.bairro)
	s_script = s_script & _
	"cx=margemx+offx+126;" & chr(13) & _
	"cw=margemx+170-cx-1;" & chr(13) & _
	"printer.imprime_campo(cx, cy, cw, '" & s & "');" & chr(13)

'	CLIENTE - CEP
	s = cep_formata(r_cliente.cep)
	s_script = s_script & _
	"cx=margemx+170+(30-printer.texto_largura('" & s & "'))/2;" & chr(13) & _
	"printer.imprime(cx, cy, '" & s & "');" & chr(13)

'>	CLIENTE - MUNICÍPIO
	s = iniciais_em_maiusculas(r_cliente.cidade)
	s_script = s_script & _
	"cx=margemx+offx;" & chr(13) & _
	"cy=margemy+78-offy-printer.texto_altura('X');" & chr(13) & _
	"cw=margemx+67-cx-1;" & chr(13) & _
	"printer.imprime_campo(cx, cy, cw, '" & s & "');" & chr(13)

'	CLIENTE - TELEFONE
	s = ""
	with r_cliente
		if Trim(.tel_res) <> "" then
			s = telefone_formata(Trim(.tel_res))
			s_aux=Trim(.ddd_res)
			if s_aux<>"" then s = "(" & s_aux & ") " & s
			end if
		
		if s = "" then
			if Trim(.tel_com) <> "" then
				s = telefone_formata(Trim(.tel_com))
				s_aux = Trim(.ddd_com)
				if s_aux<>"" then s = "(" & s_aux & ") " & s
				s_aux = Trim(.ramal_com)
				if s_aux<>"" then s = s & "  (R. " & s_aux & ")"
				end if
			end if
		end with
		
	s_script = s_script & _
	"cx=margemx+offx+67;" & chr(13) & _
	"cw=margemx+126-cx-1;" & chr(13) & _
	"printer.imprime_campo(cx, cy, cw, '" & s & "');" & chr(13)
	
'	CLIENTE - UF
	s = Ucase(r_cliente.uf)
	s_script = s_script & _
	"cx=margemx+126+(22-printer.texto_largura('" & s & "'))/2;" & chr(13) & _
	"printer.imprime(cx, cy, '" & s & "');" & chr(13)

'	CLIENTE - RG/IE
	if r_cliente.tipo = ID_PF then
		s = r_cliente.rg
	else
		s = r_cliente.ie
		end if
	s_script = s_script & _
	"cx=margemx+offx+148;" & chr(13) & _
	"cw=margemx+200-cx-1;" & chr(13) & _
	"printer.imprime_campo(cx, cy, cw, '" & s & "');" & chr(13)

'>	PRODUTOS
	s_script = s_script & _
	"cy=margemy+83+2;" & chr(13)

	m_TotalDestePedido=0
	m_TotalDestePedidoComRA=0
	n = Lbound(v_item)-1
	for i=1 to MAX_ITENS 
		n = n+1
		if n <= Ubound(v_item) then
			with v_item(n)
				s_fabricante=.fabricante
				s_produto=.produto
				s_descricao=.descricao
				s_descricao_html=produto_formata_descricao_em_html(.descricao_html)
				s_obs=.obs
				s_qtde=formata_inteiro(.qtde)
				s_vl_unitario=formata_moeda(.preco_NF)
				m_TotalItem=.qtde * .preco_venda
				m_TotalItemComRA=.qtde * .preco_NF
				s_vl_TotalItem=formata_moeda(m_TotalItem)
				m_TotalDestePedido=m_TotalDestePedido + m_TotalItem
				m_TotalDestePedidoComRA=m_TotalDestePedidoComRA + m_TotalItemComRA
				end with

		'>	QTDE
			s=s_qtde
			s_script = s_script & _
			"cx=margemx+15-offx-1-printer.texto_largura('" & s & "');" & chr(13) & _
			"printer.imprime(cx, cy, '" & s & "');" & chr(13)

		'	CÓDIGO
			s = s_produto
			s_script = s_script & _
			"cx=margemx+offx+15;" & chr(13) & _
			"printer.imprime(cx, cy, '" & s & "');" & chr(13)

		'	DESCRIÇÃO/OBSERVAÇÕES
			s = s_obs
			if (s_descricao<>"") And (s<>"") then s = " (" & s & ")"
			s = s_descricao & s
			s = filtra_texto_js(s, "'")
			s_script = s_script & _
			"cx=margemx+offx+41;" & chr(13) & _
			"cw=margemx+141-cx-1;" & chr(13) & _
			"campo='" & s & "';" & chr(13) & _
			"if (printer.texto_largura(campo)>cw) campo=iniciais_em_maiusculas(campo);" & chr(13) & _
			"printer.imprime_campo(cx, cy, cw, campo);" & chr(13)

		'	VALOR UNITÁRIO
			s=s_vl_unitario
			s_script = s_script & _
			"cx=margemx+170-offx-printer.texto_largura('" & s & "');" & chr(13) & _
			"printer.imprime(cx, cy, '" & s & "');" & chr(13)

		'	VALOR TOTAL
			s=formata_moeda(m_TotalItemComRA)
			s_script = s_script & _
			"cx=margemx+200-offx-printer.texto_largura('" & s & "');" & chr(13) & _
			"printer.imprime(cx, cy, '" & s & "');" & chr(13) & _
			"cy=cy+printer.texto_altura('X')+1;" & chr(13) 
			end if
		next

'>	VALOR TOTAL DO ORÇAMENTO
	s=formata_moeda(m_TotalDestePedidoComRA)
	s_script = s_script & _
	"cx=margemx+200-offx-printer.texto_largura('" & s & "');" & chr(13) & _
	"cy=margemy+169-1-printer.texto_altura('X');" & chr(13) & _
	"printer.imprime(cx, cy, '" & s & "');" & chr(13)

'>	FORMA DE PAGAMENTO
	s=filtra_texto_js(r_orcamento.forma_pagto, "'")
	s_script = s_script & _
	"cx=margemx;" & chr(13) & _
	"cy=margemy+185;" & chr(13) & _
	"campo='" & s & "';" & chr(13) & _
	"printer.imprime_texto(campo, cx+offx, cy+0.5, cx+(200-offx), cy+15, DT_NOPREFIX+DT_WORDBREAK);" & chr(13)

'>	OBSERVAÇÃO
	s=filtra_texto_js(r_orcamento.obs_1, "'")
	s_script = s_script & _
	"cx=margemx;" & chr(13) & _
	"cy=margemy+207;" & chr(13) & _
	"campo='" & s & "';" & chr(13) & _
	"printer.imprime_texto(campo, cx+offx, cy+0.5, cx+(200-offx), cy+15, DT_NOPREFIX+DT_WORDBREAK);" & chr(13)

s_script = s_script & _
	"printer.EndDoc();" & chr(13) & _
	"alert('Orçamento " & orcamento_selecionado & " foi impresso!!');" & chr(13) & _
	"}" & chr(13) & _
	"</script>" & chr(13)

	Response.Write s_script
%>


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
<link href="<%=URL_FILE__ESCREEN_CSS%>" Rel="stylesheet" Type="text/css" media="screen">

<style type="text/css">
#rb_etg_imediata {
	margin: 0pt 2pt 1pt 15pt;
	vertical-align: top;
	}
#rb_status {
	margin: 0pt 2pt 1pt 15pt;
	vertical-align: top;
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
<div class="MtAlerta" style="width:600px;font-weight:bold;" align="center"><P style='margin:5px 2px 5px 2px;'><%=alerta%></p></div>
<br><br>
<p class="TracoBottom"></p>
<table cellspacing="0">
<tr>
	<td align="center"><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()"><img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
</tr>
</table>
</center>
</body>




<% else %>
<!-- ************************************************************* -->
<!-- **********  PÁGINA PARA EXIBIR O ORÇAMENTO  ***************** -->
<!-- ************************************************************* -->
<body onload="ja_carregou=true;">
<center>

<form method="post" action="Pedido.asp" id="fPED" name="fPED">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="pedido_selecionado" id="pedido_selecionado" value=''>
</form>

<form id="fORC" name="fORC" method="post">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="orcamento_selecionado" id="orcamento_selecionado" value='<%=orcamento_selecionado%>'>

<!--  I D E N T I F I C A Ç Ã O   D O   O R Ç A M E N T O -->
<%=MontaHeaderIdentificacaoOrcamento(orcamento_selecionado, r_orcamento, 649)%>
<br>


<!--  L O J A   -->
<table width="649" class="Q" cellspacing="0">
	<tr>
<%	s = ""
	with r_loja
		if Trim(.razao_social) <> "" then
			s = Trim(.razao_social)
		else
			s = Trim(.nome)
			end if
		end with
	strTextoIndicador = ""
	if r_orcamento.orcamentista <> "" then
		strTextoIndicador = r_orcamento.orcamentista
		if r_orcamentista_e_indicador.desempenho_nota <> "" then
			strTextoIndicador = strTextoIndicador & " (" & r_orcamentista_e_indicador.desempenho_nota & ")"
			end if
		end if
%>
	<td class="MD"><p class="Rf">LOJA</p><p class="C"><%=s%>&nbsp;</p></td>
	<td width="145" class="MD"><p class="Rf">ORÇAMENTISTA</p><p class="C"><%=strTextoIndicador%>&nbsp;</p></td>
	<td width="145"><p class="Rf">VENDEDOR</p><p class="C"><%=r_orcamento.vendedor%>&nbsp;</p></td>
	</tr>
	</table>
	
<br>

<!--  CLIENTE   -->
<table width="649" class="Q" cellspacing="0">
	<tr>
<%	s = ""
	with r_cliente
		if Trim(.nome) <> "" then
			s = Trim(.nome)
			end if
		end with
	
	if r_cliente.tipo = ID_PF then s_aux="NOME DO CLIENTE" else s_aux="RAZÃO SOCIAL DO CLIENTE"
%>
	<td class="MD"><p class="Rf"><%=s_aux%></p><p class="C"><%=s%>&nbsp;</p></td>
		
		
<%	if r_cliente.tipo = ID_PF then s_aux="CPF" else s_aux="CNPJ"
	s = cnpj_cpf_formata(r_cliente.cnpj_cpf) 
%>
		<td width="145"><p class="Rf"><%=s_aux%></p><p class="C"><%=s%>&nbsp;</p></td>
	</tr>
	</table>

<!--  ENDEREÇO DO CLIENTE  -->
<table width="649" class="QS" cellspacing="0">
	<tr>
<%	with r_cliente
		s = formata_endereco(.endereco, .endereco_numero, .endereco_complemento, .bairro, .cidade, .uf, .cep)
		end with
%>		
		<td><p class="Rf">ENDEREÇO</p><p class="C"><%=s%>&nbsp;</p></td>
	</tr>
</table>

<!--  TELEFONE DO CLIENTE  -->
<table width="649" class="QS" cellspacing="0">
	<tr>
<%	s = ""
	with r_cliente
		if Trim(.tel_res) <> "" then
			s = telefone_formata(Trim(.tel_res))
			s_aux=Trim(.ddd_res)
			if s_aux<>"" then s = "(" & s_aux & ") " & s
			end if
		end with
	
	s2 = ""
	with r_cliente
		if Trim(.tel_com) <> "" then
			s2 = telefone_formata(Trim(.tel_com))
			s_aux = Trim(.ddd_com)
			if s_aux<>"" then s2 = "(" & s_aux & ") " & s2
			s_aux = Trim(.ramal_com)
			if s_aux<>"" then s2 = s2 & "  (R. " & s_aux & ")"
			end if
		end with

	s3 = ""
	with r_cliente
		if .tipo = ID_PF then s3 = Trim(.rg) else s3 = Trim(.ie)
		end with
%>

<% if r_cliente.tipo = ID_PF then %>
	<td class="MD" width="33%"><p class="Rf">TELEFONE RESIDENCIAL</p><p class="C"><%=s%>&nbsp;</p></td>
	<td class="MD" width="33%"><p class="Rf">TELEFONE COMERCIAL</p><p class="C"><%=s2%>&nbsp;</p></td>
	<td><p class="Rf">RG</p><p class="C"><%=s3%>&nbsp;</p></td>
<% else %>
	<td class="MD" width="50%"><p class="Rf">TELEFONE</p><p class="C"><%=s2%>&nbsp;</p></td>
	<td><p class="Rf">IE</p><p class="C"><%=s3%>&nbsp;</p></td>
<% end if %>

	</tr>
</table>

<!--  E-MAIL DO CLIENTE  -->
<table width="649" class="QS" cellspacing="0">
	<tr>
		<td><p class="Rf">E-MAIL</p><p class="C"><%=Trim(r_cliente.email)%>&nbsp;</p></td>
	</tr>
</table>

<!--  ENDEREÇO DE ENTREGA  -->
<%	with r_orcamento
		s = formata_endereco(.EndEtg_endereco, .EndEtg_endereco_numero, .EndEtg_endereco_complemento, .EndEtg_bairro, .EndEtg_cidade, .EndEtg_uf, .EndEtg_cep)
		end with
%>		
<table width="649" class="QS" cellspacing="0" style="table-layout:fixed">
	<tr>
		<td align="left"><p class="Rf">ENDEREÇO DE ENTREGA</p><p class="C"><%=s%>&nbsp;</p></td>
	</tr>
    <%	if r_orcamento.EndEtg_cod_justificativa <> "" then %>	
    <tr>
		<td align="left" style="word-wrap:break-word"><p class="C" ><%=obtem_descricao_tabela_t_codigo_descricao(GRUPO_T_CODIGO_DESCRICAO__ENDETG_JUSTIFICATIVA,r_orcamento.EndEtg_cod_justificativa)%>&nbsp;</p></td>
	</tr>
    <%end if %>
</table>



<!--  R E L A Ç Ã O   D E   P R O D U T O S  -->
<br>
<br>
<table class="Qx" cellspacing="0">
	<tr bgColor="#FFFFFF">
	<td class="MB" valign="bottom"><p class="PLTe">Fabr</p></td>
	<td class="MB" valign="bottom"><p class="PLTe">Produto</p></td>
	<td class="MB" valign="bottom"><p class="PLTe" style="width:287px;">Descrição/Observações</p></td>
	<td class="MB" valign="bottom"><p class="PLTd">Qtd</p></td>
	<% if (r_orcamento.permite_RA_status = 1) Or blnTemRA then %>
	<td class="MB" valign="bottom"><p class="PLTd">Preço</p></td>
	<% end if %>
	<td class="MB" valign="bottom"><p class="PLTd">VL Lista</p></td>
	<td class="MB" valign="bottom"><p class="PLTd">Desc</p></td>
	<td class="MB" valign="bottom"><p class="PLTd">VL Venda</p></td>
	<td class="MB" valign="bottom"><p class="PLTd">VL Total</p></td>
	</tr>

<% m_TotalDestePedido=0
   m_TotalDestePedidoComRA=0
   n = Lbound(v_item)-1
   for i=1 to MAX_ITENS 
	 n = n+1
	 if n <= Ubound(v_item) then
		with v_item(n)
			s_fabricante=.fabricante
			s_produto=.produto
			s_descricao=.descricao
			s_descricao_html=produto_formata_descricao_em_html(.descricao_html)
			s_obs=.obs
			if (s_descricao_html<>"") And (s_obs<>"") then s_obs=" (" & s_obs & ")"
			s_qtde=.qtde
			s_preco_lista=formata_moeda(.preco_lista)
			if .desc_dado=0 then s_desc_dado="" else s_desc_dado=formata_perc_desc(.desc_dado)
			s_vl_unitario=formata_moeda(.preco_venda)
			s_preco_NF=formata_moeda(.preco_NF)
			m_TotalItem=.qtde * .preco_venda
			m_TotalItemComRA=.qtde * .preco_NF
			s_vl_TotalItem=formata_moeda(m_TotalItem)
			m_TotalDestePedido=m_TotalDestePedido + m_TotalItem
			m_TotalDestePedidoComRA=m_TotalDestePedidoComRA + m_TotalItemComRA
			end with
	 else
		s_fabricante=""
		s_produto=""
		s_descricao=""
		s_descricao_html=""
		s_obs=""
		s_qtde=""
		s_preco_lista=""
		s_desc_dado=""
		s_vl_unitario=""
		s_preco_NF=""
		s_vl_TotalItem=""
		end if

'	A VERSÃO 5.0 DO IE NÃO DESENHA AS MARGENS SE O SPAN NÃO POSSUIR CONTEÚDO
	if s_descricao = "" then s_descricao = "&nbsp;"
	if s_descricao_html = "" then s_descricao_html = "&nbsp;"
	if s_obs = "" then s_obs = "&nbsp;"

%>
	<% if (i > MIN_LINHAS_ITENS_IMPRESSAO_ORCAMENTO) And (s_produto = "") then %>
	<tr class="notPrint">
	<% else %>
	<tr>
	<% end if %>
	<td class="MDBE"><input name="c_fabricante" id="c_fabricante" class="PLLe" style="width:25px;"
		value='<%=s_fabricante%>' readonly tabindex=-1></td>
	<td class="MDB"><input name="c_produto" id="c_produto" class="PLLe" style="width:54px;"
		value='<%=s_produto%>' readonly tabindex=-1></td>
	<td class="MDB"><span name="c_descricao" id="c_descricao" class="PLLe" style="margin-left:2px;"><%=s_descricao_html%></span>
					<span name="c_obs" id="c_obs" class="PLLe" style="color:navy;"><%=s_obs%></span></td>
	<td class="MDB" align="right"><input name="c_qtde" id="c_qtde" class="PLLd" style="width:21px;"
		value='<%=s_qtde%>' readonly tabindex=-1></td>
	<% if (r_orcamento.permite_RA_status = 1) Or blnTemRA then %>
	<td class="MDB" align="right"><input name="c_vl_NF" id="c_vl_NF" class="PLLd" style="width:62px;"
		value='<%=s_preco_NF%>' readonly tabindex=-1></td>
	<% end if %>
	<td class="MDB" align="right"><input name="c_preco_lista" id="c_preco_lista" class="PLLd" style="width:62px;"
		value='<%=s_preco_lista%>' readonly tabindex=-1></td>
	<td class="MDB" align="right"><input name="c_desc" id="c_desc" class="PLLd" style="width:28px;"
		value='<%=s_desc_dado%>' readonly tabindex=-1></td>
	<td class="MDB" align="right"><input name="c_vl_unitario" id="c_vl_unitario" class="PLLd" style="width:62px;"
		value='<%=s_vl_unitario%>' readonly tabindex=-1></td>
	<td class="MDB" align="right"><input name="c_vl_total" id="c_vl_total" class="PLLd" style="width:70px;" 
		value='<%=s_vl_TotalItem%>' readonly tabindex=-1></td>
	</tr>
<% next %>

	<tr>
	<td colspan="3">
		<table cellspacing="0" cellpadding="0" width='100%' style="margin-top:4px;">
			<tr>
			<td width="60%">&nbsp;</td>
			<% if (r_orcamento.permite_RA_status = 1) Or blnTemRA then %>
			<td align="right">
				<table cellspacing="0" cellpadding="0" style="margin-right:2px;">
					<tr>
						<td class="MTBE"><p class="PLTe">&nbsp;RA</p></td>
						<td class="MTBD" align="right"><input name="c_total_RA" id="c_total_RA" class="PLLd" style="width:70px;color:<%if m_total_RA >=0 then Response.Write " green" else Response.Write " red"%>;" 
							value='<%=formata_moeda(m_total_RA)%>' readonly tabindex=-1></td>
					</tr>
				</table>
			</td>
			<% end if %>
			<td align="right">
				<table cellspacing="0" cellpadding="0" style="margin-right:2px;">
					<tr>
						<td class="MTBE" nowrap><p class="PLTe">&nbsp;COM(%)</p></td>
						<td class="MTBD" align="left"><input name="c_perc_RT" id="c_perc_RT" class="PLLd" style="width:30px;color:blue;" 
							value='<%=formata_perc_RT(r_orcamento.perc_RT)%>' readonly tabindex=-1></td>
					</tr>
				</table>
			</td>
			</tr>
		</table>
	</td>
	<% if (r_orcamento.permite_RA_status = 1) Or blnTemRA then %>
	<td class="MD">&nbsp;</td>
	<td class="MDB" align="right">
		<input name="c_total_NF" id="c_total_NF" class="PLLd" style="width:70px;color:blue;" 
				value='<%=formata_moeda(m_TotalDestePedidoComRA)%>' readonly tabindex=-1>
	</td>
	<td colspan="3" class="MD">&nbsp;</td>
	<% else %>
	<td colspan="4" class="MD">&nbsp;</td>
	<% end if %>
	<td class="MDB" align="right"><input name="c_total_geral" id="c_total_geral" class="PLLd" style="width:70px;color:blue;" 
		value='<%=formata_moeda(m_TotalDestePedido)%>' readonly tabindex=-1></td>
	</tr>
</table>

<% if r_orcamento.tipo_parcelamento = 0 then %>
<!--  TRATA VERSÃO ANTIGA DA FORMA DE PAGAMENTO   -->
<br>
<table class="Q" cellspacing="0" style="width:649px;">
	<tr>
		<td class="MB" colspan="5"><p class="Rf">Observações I</p>
			<textarea name="c_obs1" id="c_obs1" class="PLLe notPrint" rows="<%=Cstr(MAX_LINHAS_OBS1)%>" 
				style="width:642px;margin-left:2pt;" 
				readonly tabindex=-1><%=r_orcamento.obs_1%></textarea>
			<span class="PLLe notVisible"><%
				s = substitui_caracteres(r_orcamento.obs_1,chr(13),"<br>")
				if s = "" then s = "&nbsp;"
				Response.Write s %></span>
		</td>
	</tr>
	<tr>
		<td class="MB" colspan="5"><p class="Rf">Observações II</p>
			<input name="c_obs2" id="c_obs2" class="PLLe" style="width:85px;margin-left:2pt;" 
				readonly tabindex=-1 value='<%=r_orcamento.obs_2%>'>
		</td>
	</tr>
	<tr>
		<td class="MDB" nowrap width="10%"><p class="Rf">Parcelas</p>
			<input name="c_qtde_parcelas" id="c_qtde_parcelas" class="PLLc" style="width:60px;"
				readonly tabindex=-1 value='<%if (r_orcamento.qtde_parcelas<>0) Or (r_orcamento.forma_pagto<>"") then Response.write Cstr(r_orcamento.qtde_parcelas)%>'>
		</td>
		<td class="MDB" nowrap valign="top"><p class="Rf">Entrega Imediata</p>
		<% 	if Cstr(r_orcamento.st_etg_imediata) = Cstr(COD_ETG_IMEDIATA_NAO) then
				s = "NÃO"
			elseif Cstr(r_orcamento.st_etg_imediata) = Cstr(COD_ETG_IMEDIATA_SIM) then
				s = "SIM"
			else
				s = ""
				end if
			
			if s <> "" then
				s_aux=formata_data_e_talvez_hora_hhmm(r_orcamento.etg_imediata_data)
				if s_aux <> "" then s = s & " &nbsp; (" & r_orcamento.etg_imediata_usuario & " em " & s_aux & ")"
				end if	
			if s="" then s="&nbsp;"
		%>
		<p class="C" style="margin-top:3px;"><%=s%></p>
		</td>
		<td class="MDB" nowrap valign="top"><p class="Rf">Bem de Uso/Consumo</p>
		<% 	if Cstr(r_orcamento.StBemUsoConsumo) = Cstr(COD_ST_BEM_USO_CONSUMO_NAO) then
				s = "NÃO"
			elseif Cstr(r_orcamento.StBemUsoConsumo) = Cstr(COD_ST_BEM_USO_CONSUMO_SIM) then
				s = "SIM"
			else
				s = ""
				end if
		
			if s="" then s="&nbsp;"
		%>
		<p class="C" style="margin-top:3px;"><%=s%></p>
		</td>
		<td class="MDB" nowrap valign="top"><p class="Rf">Instalador Instala</p>
		<% 	if Cstr(r_orcamento.InstaladorInstalaStatus) = Cstr(COD_INSTALADOR_INSTALA_NAO) then
				s = "NÃO"
			elseif Cstr(r_orcamento.InstaladorInstalaStatus) = Cstr(COD_INSTALADOR_INSTALA_SIM) then
				s = "SIM"
			else
				s = ""
				end if
		
			if s="" then s="&nbsp;"
		%>
		<p class="C" style="margin-top:3px;"><%=s%></p>
		</td>
		<td class="MB" nowrap valign="top"><p class="Rf">Garantia Indicador</p>
		<% 	if Cstr(r_orcamento.GarantiaIndicadorStatus) = Cstr(COD_GARANTIA_INDICADOR_STATUS__NAO) then
				s = "NÃO"
			elseif Cstr(r_orcamento.GarantiaIndicadorStatus) = Cstr(COD_GARANTIA_INDICADOR_STATUS__SIM) then
				s = "SIM"
			else
				s = ""
				end if
		
			if s="" then s="&nbsp;"
		%>
		<p class="C" style="margin-top:3px;"><%=s%></p>
		</td>
	</tr>
	<tr>
		<td colspan="5"><p class="Rf">Forma de Pagamento</p>
			<textarea name="c_forma_pagto" id="c_forma_pagto" class="PLLe notPrint" rows="<%=Cstr(MAX_LINHAS_FORMA_PAGTO)%>" 
				style="width:642px;margin-left:2pt;"
				readonly tabindex=-1><%=r_orcamento.forma_pagto%></textarea>
		<span class="PLLe notVisible"><%
			s = substitui_caracteres(r_orcamento.forma_pagto,chr(13),"<br>")
			if s = "" then s = "&nbsp;"
			Response.Write s %></span>
		</td>
	</tr>
</table>
<% else %>
<!--  TRATA NOVA VERSÃO DA FORMA DE PAGAMENTO   -->
<br>
<table class="Q" cellspacing="0" style="width:649px;">
	<tr>
		<td class="MB" colspan="5"><p class="Rf">Observações I</p>
			<textarea name="c_obs1" id="c_obs1" class="PLLe notPrint" rows="<%=Cstr(MAX_LINHAS_OBS1)%>" 
				style="width:642px;margin-left:2pt;" 
				readonly tabindex=-1><%=r_orcamento.obs_1%></textarea>
			<span class="PLLe notVisible"><%
				s = substitui_caracteres(r_orcamento.obs_1,chr(13),"<br>")
				if s = "" then s = "&nbsp;"
				Response.Write s %></span>
		</td>
	</tr>
	<tr>
		<td class="MD" nowrap><p class="Rf">Observações II</p>
			<input name="c_obs2" id="c_obs2" class="PLLe" style="width:85px;margin-left:2pt;" 
				readonly tabindex=-1 value='<%=r_orcamento.obs_2%>'>
		</td>
		<td class="MD" nowrap valign="top"><p class="Rf">Entrega Imediata</p>
		<% 	if Cstr(r_orcamento.st_etg_imediata) = Cstr(COD_ETG_IMEDIATA_NAO) then
				s = "NÃO"
			elseif Cstr(r_orcamento.st_etg_imediata) = Cstr(COD_ETG_IMEDIATA_SIM) then
				s = "SIM"
			else
				s = ""
				end if
			
			if s <> "" then
				s_aux=formata_data_e_talvez_hora_hhmm(r_orcamento.etg_imediata_data)
				if s_aux <> "" then s = s & " &nbsp; (" & r_orcamento.etg_imediata_usuario & " em " & s_aux & ")"
				end if
			if s="" then s="&nbsp;"
		%>
		<p class="C" style="margin-top:3px;"><%=s%></p>
		</td>
		<td class="MD" nowrap valign="top"><p class="Rf">Bem de Uso/Consumo</p>
		<% 	if Cstr(r_orcamento.StBemUsoConsumo) = Cstr(COD_ST_BEM_USO_CONSUMO_NAO) then
				s = "NÃO"
			elseif Cstr(r_orcamento.StBemUsoConsumo) = Cstr(COD_ST_BEM_USO_CONSUMO_SIM) then
				s = "SIM"
			else
				s = ""
				end if
		
			if s="" then s="&nbsp;"
		%>
		<p class="C" style="margin-top:3px;"><%=s%></p>
		</td>
		<td class="MD" nowrap valign="top"><p class="Rf">Instalador Instala</p>
		<% 	if Cstr(r_orcamento.InstaladorInstalaStatus) = Cstr(COD_INSTALADOR_INSTALA_NAO) then
				s = "NÃO"
			elseif Cstr(r_orcamento.InstaladorInstalaStatus) = Cstr(COD_INSTALADOR_INSTALA_SIM) then
				s = "SIM"
			else
				s = ""
				end if
		
			if s="" then s="&nbsp;"
		%>
		<p class="C" style="margin-top:3px;"><%=s%></p>
		</td>
		<td nowrap valign="top"><p class="Rf">Garantia Indicador</p>
		<% 	if Cstr(r_orcamento.GarantiaIndicadorStatus) = Cstr(COD_GARANTIA_INDICADOR_STATUS__NAO) then
				s = "NÃO"
			elseif Cstr(r_orcamento.GarantiaIndicadorStatus) = Cstr(COD_GARANTIA_INDICADOR_STATUS__SIM) then
				s = "SIM"
			else
				s = ""
				end if
		
			if s="" then s="&nbsp;"
		%>
		<p class="C" style="margin-top:3px;"><%=s%></p>
		</td>
	</tr>
</table>
<br>
<table class="Q" style="width:649px;" cellspacing="0">
  <tr>
	<td><p class="Rf">Forma de Pagamento</p></td>
  </tr>
  <tr>
	<td>
	  <table width="100%" cellspacing="0" cellpadding="0" border="0">
		<% if Cstr(r_orcamento.tipo_parcelamento) = COD_FORMA_PAGTO_A_VISTA then %>
		<!--  À VISTA  -->
		<tr>
		  <td>
			<table cellspacing="0" cellpadding="0" border="0">
			  <tr>
				<td><span class="C">À Vista&nbsp&nbsp;(<%=x_opcao_forma_pagamento(r_orcamento.av_forma_pagto)%>)</span></td>
			  </tr>
			</table>
		  </td>
		</tr>
		<% elseif Cstr(r_orcamento.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELA_UNICA then %>
		<!--  PARCELA ÚNICA  -->
		<tr>
		  <td>
			<table cellspacing="0" cellpadding="0" border="0">
			  <tr>
				<td><span class="C">Parcela Única:&nbsp;&nbsp;<%=SIMBOLO_MONETARIO%>&nbsp;<%=formata_moeda(r_orcamento.pu_valor)%>&nbsp;&nbsp;&nbsp;(<%=x_opcao_forma_pagamento(r_orcamento.pu_forma_pagto)%>)&nbsp;&nbsp;vencendo após&nbsp;<%=formata_inteiro(r_orcamento.pu_vencto_apos)%>&nbsp;dias</span></td>
			  </tr>
			</table>
		  </td>
		</tr>
		<% elseif Cstr(r_orcamento.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_CARTAO then %>
		<!--  PARCELADO NO CARTÃO (INTERNET)  -->
		<tr>
		  <td>
			<table cellspacing="0" cellpadding="0" border="0">
			  <tr>
				<td><span class="C">Parcelado no Cartão (internet) em&nbsp;&nbsp;<%=Cstr(r_orcamento.pc_qtde_parcelas)%>&nbsp;x&nbsp;&nbsp;<%=SIMBOLO_MONETARIO & " " & formata_moeda(r_orcamento.pc_valor_parcela)%></span></td>
			  </tr>
			</table>
		  </td>
		</tr>
		<% elseif Cstr(r_orcamento.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_CARTAO_MAQUINETA then %>
		<!--  PARCELADO NO CARTÃO (MAQUINETA)  -->
		<tr>
		  <td>
			<table cellspacing="0" cellpadding="0" border="0">
			  <tr>
				<td><span class="C">Parcelado no Cartão (maquineta) em&nbsp;&nbsp;<%=Cstr(r_orcamento.pc_maquineta_qtde_parcelas)%>&nbsp;x&nbsp;&nbsp;<%=SIMBOLO_MONETARIO & " " & formata_moeda(r_orcamento.pc_maquineta_valor_parcela)%></span></td>
			  </tr>
			</table>
		  </td>
		</tr>
		<% elseif Cstr(r_orcamento.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_COM_ENTRADA then %>
		<!--  PARCELADO COM ENTRADA  -->
		<tr>
		  <td>
			<table cellspacing="0" cellpadding="0" border="0">
			  <tr>
				<td><span class="C">Entrada:&nbsp;&nbsp;<%=SIMBOLO_MONETARIO%>&nbsp;<%=formata_moeda(r_orcamento.pce_entrada_valor)%>&nbsp;&nbsp;&nbsp;(<%=x_opcao_forma_pagamento(r_orcamento.pce_forma_pagto_entrada)%>)</span></td>
			  </tr>
			  <tr>
				<td><span class="C">Prestações:&nbsp;&nbsp;<%=formata_inteiro(r_orcamento.pce_prestacao_qtde)%>&nbsp;x&nbsp;<%=SIMBOLO_MONETARIO%>&nbsp;<%=formata_moeda(r_orcamento.pce_prestacao_valor)%>&nbsp;&nbsp;&nbsp;(<%=x_opcao_forma_pagamento(r_orcamento.pce_forma_pagto_prestacao)%>)&nbsp;&nbsp;vencendo a cada&nbsp;<%=formata_inteiro(r_orcamento.pce_prestacao_periodo)%>&nbsp;dias</span></td>
			  </tr>
			</table>
		  </td>
		</tr>
		<% elseif Cstr(r_orcamento.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA then %>
		<!--  PARCELADO SEM ENTRADA  -->
		<tr>
		  <td>
			<table cellspacing="0" cellpadding="0" border="0">
			  <tr>
				<td><span class="C">1ª Prestação:&nbsp;&nbsp;<%=SIMBOLO_MONETARIO%>&nbsp;<%=formata_moeda(r_orcamento.pse_prim_prest_valor)%>&nbsp;&nbsp;&nbsp;(<%=x_opcao_forma_pagamento(r_orcamento.pse_forma_pagto_prim_prest)%>)&nbsp;&nbsp;vencendo após&nbsp;<%=formata_inteiro(r_orcamento.pse_prim_prest_apos)%>&nbsp;dias</span></td>
			  </tr>
			  <tr>
				<td><span class="C">Demais Prestações:&nbsp;&nbsp;<%=Cstr(r_orcamento.pse_demais_prest_qtde)%>&nbsp;x&nbsp;<%=SIMBOLO_MONETARIO%>&nbsp;<%=formata_moeda(r_orcamento.pse_demais_prest_valor)%>&nbsp;&nbsp;&nbsp;(<%=x_opcao_forma_pagamento(r_orcamento.pse_forma_pagto_demais_prest)%>)&nbsp;&nbsp;vencendo a cada&nbsp;<%=Cstr(r_orcamento.pse_demais_prest_periodo)%>&nbsp;dias</span></td>
			  </tr>
			</table>
		  </td>
		</tr>
		<% end if %>
	  </table>
	</td>
  </tr>
  <tr>
	<td class="MC"><p class="Rf">Descrição da Forma de Pagamento</p>
	  <textarea name="c_forma_pagto" id="c_forma_pagto" class="PLLe notPrint" rows="<%=Cstr(MAX_LINHAS_FORMA_PAGTO)%>"
				style="width:642px;margin-left:2pt;"
				readonly tabindex=-1><%=r_orcamento.forma_pagto%></textarea>
		<span class="PLLe notVisible"><%
			s = substitui_caracteres(r_orcamento.forma_pagto,chr(13),"<br>")
			if s = "" then s = "&nbsp;"
			Response.Write s %></span>
	</td>
  </tr>
</table>
<% end if %>


<!-- ************   SEPARADOR   ************ -->
<table class="notPrint" width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>

<table class="notPrint" width="649" cellpadding="0" cellspacing="0">
<tr><td align="right"><span class="Rc">
	<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
	<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
	</span>
</td></tr>
</table>

<!-- ************   BOTÕES   ************ -->
<table class="notPrint" width="649" cellpadding="0" cellspacing="0">
<tr>
	<td align="left"><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()" title="volta para página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a>
	</td>
	<td align="center"><div name="dIMPRESSORA" id="dIMPRESSORA">
		<a name="bIMPRESSORA" id="bIMPRESSORA" href="javascript:fORCImpressora(fORC)" title="seleciona a impressora">
		<img src="../botao/impressora.gif" width="176" height="55" border="0"></a></div>
	</td>
	<td align="center"><div name="dMARGENS" id="dMARGENS">
		<a name="bMARGENS" id="bMARGENS" href="javascript:fORCMargens(fORC)" title="configura as margens de impressão">
		<img src="../botao/margens.gif" width="176" height="55" border="0"></a></div>
	</td>
	<td align="right"><div name="dIMPRIME" id="dIMPRIME">
		<a name="bIMPRIME" id="bIMPRIME" href="javascript:fORCImprime(fORC)" title="imprime o orçamento em formulário contínuo">
		<img src="../botao/imprimir.gif" width="176" height="55" border="0"></a></div>
	</td>
</tr>

</table>

</form>

</center>
</body>

<% end if %>

</html>


<%
'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>
