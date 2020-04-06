<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/estoque.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ===========================================
'	  P E D I D O I M P R I M E . A S P
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

	dim s, usuario, pedido_selecionado
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

	pedido_selecionado = ucase(Trim(request("pedido_selecionado")))
	if (pedido_selecionado = "") then Response.Redirect("aviso.asp?id=" & ERR_PEDIDO_NAO_ESPECIFICADO)
	s = normaliza_num_pedido(pedido_selecionado)
	if s <> "" then pedido_selecionado = s
	
	dim i, n, s_fabricante, s_produto, s_descricao, s_descricao_html, s_qtde, s_preco_lista, s_desc_dado
	dim s_vl_unitario, s_vl_TotalItem, m_TotalItem, m_TotalDestePedido, m_TotalItemComRA, m_TotalDestePedidoComRA
	dim s_preco_NF, m_TotalFamiliaParcelaRA
	
'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

	dim r_pedido, v_item, alerta, msg_erro
	alerta=""
	if Not le_pedido(pedido_selecionado, r_pedido, msg_erro) then 
		alerta = msg_erro
	else
		if Not le_pedido_item(pedido_selecionado, v_item, msg_erro) then alerta = msg_erro
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

	dim s_aux, s2, s3, r_loja, r_cliente, s_cor, s_falta, v_pedido, s_script
	dim v_disp
	dim vl_TotalFamiliaPrecoVenda, vl_TotalFamiliaPrecoNF, vl_TotalFamiliaPago, vl_TotalFamiliaDevolucaoPrecoVenda, vl_TotalFamiliaDevolucaoPrecoNF
	dim vl_saldo_a_pagar, s_vl_saldo_a_pagar, st_pagto
	dim v_item_devolvido, s_devolucoes
	dim v_pedido_perda, s_perdas, vl_total_perdas
    dim cliente__tipo, cliente__cnpj_cpf, cliente__rg, cliente__ie, cliente__nome
    dim cliente__endereco, cliente__endereco_numero, cliente__endereco_complemento, cliente__bairro, cliente__cidade, cliente__uf, cliente__cep
    dim cliente__tel_res, cliente__ddd_res, cliente__tel_com, cliente__ddd_com, cliente__ramal_com, cliente__tel_cel, cliente__ddd_cel
    dim cliente__tel_com_2, cliente__ddd_com_2, cliente__ramal_com_2, cliente__email

	s_devolucoes = ""
	s_perdas = ""
	vl_total_perdas = 0
	
	dim total_cubagem, total_volumes, total_peso
	total_cubagem = 0
	total_volumes = 0
	total_peso = 0
	if alerta = "" then
		for i=Lbound(v_item) to Ubound(v_item)
			with v_item(i)
				if Trim("" & .produto) <> "" then
					total_cubagem = total_cubagem + (.qtde * .cubagem)
					total_volumes = total_volumes + (.qtde * .qtde_volumes)
					total_peso = total_peso + (.qtde * .peso)
					end if
				end with
			next
		end if
	
	if alerta = "" then
		if Not IsPedidoEncerrado(r_pedido.st_entrega) then
			redim v_disp(Ubound(v_item))
			for i=Lbound(v_disp) to Ubound(v_disp)
				set v_disp(i) = New cl_ITEM_STATUS_ESTOQUE
				v_disp(i).pedido		= v_item(i).pedido
				v_disp(i).fabricante	= v_item(i).fabricante
				v_disp(i).produto		= v_item(i).produto
				v_disp(i).qtde			= v_item(i).qtde
				next
			
			if Not estoque_verifica_status_item(v_disp, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
			end if
			
	'	OBTÉM OS NÚMEROS DE PEDIDOS QUE COMPÕEM ESTA FAMÍLIA DE PEDIDOS
		if Not recupera_familia_pedido(pedido_selecionado, v_pedido, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
	'	OBTÉM OS VALORES A PAGAR, JÁ PAGO E O STATUS DE PAGAMENTO (PARA TODA A FAMÍLIA DE PEDIDOS)
		if Not calcula_pagamentos(pedido_selecionado, vl_TotalFamiliaPrecoVenda, vl_TotalFamiliaPrecoNF, vl_TotalFamiliaPago, vl_TotalFamiliaDevolucaoPrecoVenda, vl_TotalFamiliaDevolucaoPrecoNF, st_pagto, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
		m_TotalFamiliaParcelaRA = vl_TotalFamiliaPrecoNF - vl_TotalFamiliaPrecoVenda
		vl_saldo_a_pagar = vl_TotalFamiliaPrecoNF - vl_TotalFamiliaPago - vl_TotalFamiliaDevolucaoPrecoNF
		s_vl_saldo_a_pagar = formata_moeda(vl_saldo_a_pagar)
	'	VALORES NEGATIVOS REPRESENTAM O 'CRÉDITO' QUE O CLIENTE POSSUI EM CASO DE PEDIDOS CANCELADOS QUE HAVIAM SIDO PAGOS
		if (st_pagto = ST_PAGTO_PAGO) And (vl_saldo_a_pagar > 0) then s_vl_saldo_a_pagar = ""
		
	'	HÁ DEVOLUÇÕES?
		if Not le_pedido_item_devolvido(pedido_selecionado, v_item_devolvido, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
		for i=Lbound(v_item_devolvido) to Ubound(v_item_devolvido)
			with v_item_devolvido(i)
				if .produto <> "" then
					if .qtde = 1 then s = "" else s = "s"
					if s_devolucoes <> "" then s_devolucoes = s_devolucoes & chr(13) & "<br>" & chr(13)
					s_devolucoes = s_devolucoes & formata_data(.devolucao_data) & " " & _
								   formata_hhnnss_para_hh_nn(.devolucao_hora) & " - " & _
								   formata_inteiro(.qtde) & " unidade" & s & " do " & .produto & " - " & produto_formata_descricao_em_html(.descricao_html)
					if Trim(.motivo) <> "" then	s_devolucoes = s_devolucoes & " (" & .motivo & ")"
					if .NFe_numero_NF > 0 then s_devolucoes = s_devolucoes & " [NF: " & .NFe_numero_NF & "]"
					end if
				end with
			next
		
	'	HÁ PERDAS?
		if Not le_pedido_perda(pedido_selecionado, v_pedido_perda, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
		for i=Lbound(v_pedido_perda) to Ubound(v_pedido_perda)
			with v_pedido_perda(i)
				if .id <> "" then
					vl_total_perdas = vl_total_perdas + .valor
					if s_perdas <> "" then s_perdas = s_perdas & chr(13) & "<br>" & chr(13)
					s_perdas = s_perdas & formata_data(.data) & " " & _
							   formata_hhnnss_para_hh_nn_ss(.hora) & ": " & SIMBOLO_MONETARIO & " " & formata_moeda(.valor)
					if Trim(.obs) <> "" then s_perdas = s_perdas & " (" & .obs & ")"
					end if
				end with
			next
		
		set r_loja = New cl_LOJA
		if Not x_loja_bd(r_pedido.loja, r_loja) then Response.Redirect("aviso.asp?id=" & ERR_LOJA_NAO_CADASTRADA)
		
		set r_cliente = New cl_CLIENTE
		if Not x_cliente_bd(r_pedido.id_cliente, r_cliente) then Response.Redirect("aviso.asp?id=" & ERR_CLIENTE_NAO_CADASTRADO)

        'le as variáveis da origem certa: ou do pedido ou do cliente, todas comecam com cliente__
        cliente__tipo = r_cliente.tipo
        cliente__cnpj_cpf = r_cliente.cnpj_cpf
	    cliente__rg = r_cliente.rg
        cliente__ie = r_cliente.ie
        cliente__nome = r_cliente.nome
        cliente__endereco = r_cliente.endereco
        cliente__endereco_numero = r_cliente.endereco_numero
        cliente__endereco_complemento = r_cliente.endereco_complemento
        cliente__bairro = r_cliente.bairro
        cliente__cidade = r_cliente.cidade
        cliente__uf = r_cliente.uf
        cliente__cep = r_cliente.cep
        cliente__tel_res = r_cliente.tel_res
        cliente__ddd_res = r_cliente.ddd_res
        cliente__tel_com = r_cliente.tel_com
        cliente__ddd_com = r_cliente.ddd_com
        cliente__ramal_com = r_cliente.ramal_com
        cliente__tel_cel = r_cliente.tel_cel
        cliente__ddd_cel = r_cliente.ddd_cel
        cliente__tel_com_2 = r_cliente.tel_com_2
        cliente__ddd_com_2 = r_cliente.ddd_com_2
        cliente__ramal_com_2 = r_cliente.ramal_com_2
        cliente__email = r_cliente.email

        if isActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos and r_pedido.st_memorizacao_completa_enderecos <> 0 then 
            cliente__tipo = r_pedido.endereco_tipo_pessoa
            cliente__cnpj_cpf = r_pedido.endereco_cnpj_cpf
	        cliente__rg = r_pedido.endereco_rg
            cliente__ie = r_pedido.endereco_ie
            cliente__nome = r_pedido.endereco_nome
            cliente__endereco = r_pedido.endereco_logradouro
            cliente__endereco_numero = r_pedido.endereco_numero
            cliente__endereco_complemento = r_pedido.endereco_complemento
            cliente__bairro = r_pedido.endereco_bairro
            cliente__cidade = r_pedido.endereco_cidade
            cliente__uf = r_pedido.endereco_uf
            cliente__cep = r_pedido.endereco_cep
            cliente__tel_res = r_pedido.endereco_tel_res
            cliente__ddd_res = r_pedido.endereco_ddd_res
            cliente__tel_com = r_pedido.endereco_tel_com
            cliente__ddd_com = r_pedido.endereco_ddd_com
            cliente__ramal_com = r_pedido.endereco_ramal_com
            cliente__tel_cel = r_pedido.endereco_tel_cel
            cliente__ddd_cel = r_pedido.endereco_ddd_cel
            cliente__tel_com_2 = r_pedido.endereco_tel_com_2
            cliente__ddd_com_2 = r_pedido.endereco_ddd_com_2
            cliente__ramal_com_2 = r_pedido.endereco_ramal_com_2
            cliente__email = r_pedido.endereco_email
            end if
		end if

	dim strTextoIndicador
	dim r_orcamentista_e_indicador
	if alerta = "" then
		call le_orcamentista_e_indicador(r_pedido.indicador, r_orcamentista_e_indicador, msg_erro)
		end if




' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________

' ___________________________________
' EXIBE_FAMILIA_PEDIDO
'
function exibe_familia_pedido(byval pedido_selecionado, byref v_pedido)
const PEDIDOS_POR_LINHA = 8
dim i
dim n
dim x
	exibe_familia_pedido = ""
	if Ubound(v_pedido) = Lbound(v_pedido) then exit function

	x = "<table width='649' class='Q' cellSpacing='0'>" & chr(13) & _
		"<tr><td>" & chr(13) & _
		"<p class='Rf'>FAMÍLIA DE PEDIDOS</p><p class='C'>" & chr(13) & _
		"<table width='100%' class='QT' cellSpacing='0'>" & chr(13) & _
		"<tr>" & chr(13)
	
	n = 0
	for i = Lbound(v_pedido) to Ubound(v_pedido)
		if Trim(v_pedido(i))<>"" then
			n = n+1
			if n > PEDIDOS_POR_LINHA then 
				n = 1
				x = x & "</tr>" & chr(13) & "<tr>"
				end if
			x = x & "<td width='12.5%' class='L' style='text-align:left;color:black;'>"
			if v_pedido(i) <> pedido_selecionado then 
				x = x & "<a href='pedido.asp?pedido_selecionado=" & Trim(v_pedido(i)) & "&url_back=X" & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")) & _
						"' title='clique para consultar o pedido' class='L' style='color:purple;'>"
				end if
			x = x & Trim(v_pedido(i))
			if v_pedido(i) <> pedido_selecionado then x = x & "</a>"
			x = x & "</td>" & chr(13)
			end if
		next
	
	if (n Mod PEDIDOS_POR_LINHA)<> 0 then
		for i = ((n Mod PEDIDOS_POR_LINHA)+1) to PEDIDOS_POR_LINHA
			x = x & "<td>&nbsp;</td>" & chr(13)
			next
		end if
	
	x = x & "</tr></table>" & chr(13) & _
			"</td></tr></table>" & chr(13) & _
			"<br>"
	
	exibe_familia_pedido = x
end function

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
	<title>CENTRAL<%=MontaNumPedidoExibicaoTitleBrowser(pedido_selecionado)%></title>
</head>


<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
var ja_carregou=false;
function impressora_carregada() {
var s;
	if (!( "object" == typeof(printer) && "string" == typeof(printer.module_id))) {
		s = "Componente necessário para impressão do pedido não foi carregado corretamente!!";
		alert(s);
		return false;
		}
	return true;
}

function fPEDImpressora ( f ) {
	if (!ja_carregou) return;
	if (!impressora_carregada()) return;
	printer.Initialize;
	if (printer.printing) printer.EndDoc();
	printer.seleciona_impressora();
}

function fPEDMargens( f ) {
	if (!ja_carregou) return;
	if (!impressora_carregada()) return;
	printer.configura_margens();
}
</script>

<% 
s_script = "<script language='JavaScript' type='text/javascript'>" & chr(13) & _
"function fPEDImprime( f ) {" & chr(13) & _
"var cx, cy, cw, offx, offy, margemx, margemy,campo;" & chr(13) & _
	"if (!ja_carregou) return;" & chr(13) & _
	"if (!impressora_carregada(printer)) return;" & chr(13) & _
	"printer.Initialize;" & chr(13) & _
	"if (printer.printing) printer.EndDoc();" & chr(13) & _
	"printer.landscape=false;" & chr(13) & _
	"printer.setpapersizeletter();" & chr(13) & _
	"printer.job_title='PEDIDO " & pedido_selecionado & "';" & chr(13) & _
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

'>	Nº PEDIDO
	s = pedido_selecionado
	s_script = s_script & _
	"printer.fontsize=12;" & chr(13) & _
	"cx=margemx+170+(30-printer.texto_largura('" & s & "'))/2;" & chr(13) & _
	"cy=margemy+15-offy-2-printer.texto_altura('X');" & chr(13) & _
	"printer.imprime(cx, cy, '" & s & "');" & chr(13)
	
'>	DATA
	s = formata_data(r_pedido.data)
	s_script = s_script & _
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
	"printer.imprime_campo(cx, cy, cw, '" & r_pedido.vendedor & "');" & chr(13)

'>	CLIENTE - NOME
	s = iniciais_em_maiusculas(cliente__nome)
	s_script = s_script & _
	"cx=margemx+offx;" & chr(13) & _
	"cy=margemy+53-offy-printer.texto_altura('X');" & chr(13) & _
	"cw=margemx+148-cx-1;" & chr(13) & _
	"printer.imprime_campo(cx, cy, cw, '" & s & "');" & chr(13)

'	CLIENTE - CNPJ/CPF
	s = cnpj_cpf_formata(cliente__cnpj_cpf)
	s_script = s_script & _
	"cx=margemx+offx+148;" & chr(13) & _
	"cw=margemx+200-cx-1;" & chr(13) & _
	"printer.imprime_campo(cx, cy, cw, '" & s & "');" & chr(13)

'>	CLIENTE - ENDEREÇO
	s = iniciais_em_maiusculas(cliente__endereco)
	if cliente__endereco_numero <> "" then s = s & ", " & cliente__endereco_numero
	if cliente__endereco_complemento <> "" then s = s & " " & cliente__endereco_complemento
	s_script = s_script & _
	"cx=margemx+offx;" & chr(13) & _
	"cy=margemy+66-offy-printer.texto_altura('X');" & chr(13) & _
	"cw=margemx+126-cx-1;" & chr(13) & _
	"printer.imprime_campo(cx, cy, cw, '" & s & "');" & chr(13)

'	CLIENTE - BAIRRO
	s = iniciais_em_maiusculas(cliente__bairro)
	s_script = s_script & _
	"cx=margemx+offx+126;" & chr(13) & _
	"cw=margemx+170-cx-1;" & chr(13) & _
	"printer.imprime_campo(cx, cy, cw, '" & s & "');" & chr(13)

'	CLIENTE - CEP
	s = cep_formata(cliente__cep)
	s_script = s_script & _
	"cx=margemx+170+(30-printer.texto_largura('" & s & "'))/2;" & chr(13) & _
	"printer.imprime(cx, cy, '" & s & "');" & chr(13)

'>	CLIENTE - MUNICÍPIO
	s = iniciais_em_maiusculas(cliente__cidade)
	s_script = s_script & _
	"cx=margemx+offx;" & chr(13) & _
	"cy=margemy+78-offy-printer.texto_altura('X');" & chr(13) & _
	"cw=margemx+67-cx-1;" & chr(13) & _
	"printer.imprime_campo(cx, cy, cw, '" & s & "');" & chr(13)

'	CLIENTE - TELEFONE
	s = ""
	if Trim(cliente__tel_res) <> "" then
		s = telefone_formata(Trim(cliente__tel_res))
		s_aux=Trim(cliente__ddd_res)
		if s_aux<>"" then s = "(" & s_aux & ") " & s
		end if
		
	if s = "" then
		if Trim(cliente__tel_com) <> "" then
			s = telefone_formata(Trim(cliente__tel_com))
			s_aux = Trim(cliente__ddd_com)
			if s_aux<>"" then s = "(" & s_aux & ") " & s
			s_aux = Trim(cliente__ramal_com)
			if s_aux<>"" then s = s & "  (R. " & s_aux & ")"
			end if
		end if
	
	s_script = s_script & _
	"cx=margemx+offx+67;" & chr(13) & _
	"cw=margemx+126-cx-1;" & chr(13) & _
	"printer.imprime_campo(cx, cy, cw, '" & s & "');" & chr(13)
	
'	CLIENTE - UF
	s = Ucase(cliente__uf)
	s_script = s_script & _
	"cx=margemx+126+(22-printer.texto_largura('" & s & "'))/2;" & chr(13) & _
	"printer.imprime(cx, cy, '" & s & "');" & chr(13)

'	CLIENTE - RG/IE
	if cliente__tipo = ID_PF then
		s = cliente__rg
	else
		s = cliente__ie
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

		'	DESCRIÇÃO
			s=s_descricao
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

'>	VALOR TOTAL DO PEDIDO
	s=formata_moeda(m_TotalDestePedidoComRA)
	s_script = s_script & _
	"cx=margemx+200-offx-printer.texto_largura('" & s & "');" & chr(13) & _
	"cy=margemy+169-1-printer.texto_altura('X');" & chr(13) & _
	"printer.imprime(cx, cy, '" & s & "');" & chr(13)

s_script = s_script & _
	"printer.EndDoc();" & chr(13) & _
	"alert('Pedido " & pedido_selecionado & " foi impresso!!');" & chr(13) & _
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
	<td align="center"><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()"><img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
</tr>
</table>
</center>
</body>




<% else %>
<!-- ********************************************************** -->
<!-- **********  PÁGINA PARA EXIBIR O PEDIDO  ***************** -->
<!-- ********************************************************** -->
<body onload="ja_carregou=true;">
<center>

<form id="fPED" name="fPED" method="post">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="pedido_selecionado" id="pedido_selecionado" value='<%=pedido_selecionado%>'>

<!--  I D E N T I F I C A Ç Ã O   D O   P E D I D O -->  
<%=MontaHeaderIdentificacaoPedido(pedido_selecionado, r_pedido, 649)%>
<br>


<!--  EXIBE LINKS PARA A FAMÍLIA DE PEDIDOS?   -->
<%=exibe_familia_pedido(pedido_selecionado, v_pedido)%>


<!--  L O J A   -->
<table width="649" class="Q" cellSpacing="0">
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
	if r_pedido.indicador <> "" then
		strTextoIndicador = r_pedido.indicador
		if r_orcamentista_e_indicador.desempenho_nota <> "" then
			strTextoIndicador = strTextoIndicador & " (" & r_orcamentista_e_indicador.desempenho_nota & ")"
			end if
		end if
%>
	<td class="MD"><p class="Rf">LOJA</p><p class="C"><%=s%>&nbsp;</p></td>
	<td width="145" class="MD"><p class="Rf">INDICADOR</p><p class="C"><%=strTextoIndicador%>&nbsp;</p></td>
	<td width="145"><p class="Rf">VENDEDOR</p><p class="C"><%=r_pedido.vendedor%>&nbsp;</p></td>
	</tr>
	</table>

<br>

<!--  CLIENTE   -->
<table width="649" class="Q" cellSpacing="0">
	<tr>
<%	s = ""
	if Trim(cliente__nome) <> "" then
		s = Trim(cliente__nome)
		end if
	
	if cliente__tipo = ID_PF then s_aux="NOME DO CLIENTE" else s_aux="RAZÃO SOCIAL DO CLIENTE"
%>
	<td class="MD"><p class="Rf"><%=s_aux%></p><p class="C"><%=s%>&nbsp;</p></td>
		
		
<%	if cliente__tipo = ID_PF then s_aux="CPF" else s_aux="CNPJ"
	s = cnpj_cpf_formata(cliente__cnpj_cpf) 
%>
		<td width="145"><p class="Rf"><%=s_aux%></p><p class="C"><%=s%>&nbsp;</p></td>
	</tr>
	</table>

<!--  ENDEREÇO DO CLIENTE  -->
<table width="649" class="QS" cellSpacing="0">
	<tr>
<%	
	s = formata_endereco(cliente__endereco, cliente__endereco_numero, cliente__endereco_complemento, cliente__bairro, cliente__cidade, cliente__uf, cliente__cep)
%>		
		<td><p class="Rf">ENDEREÇO</p><p class="C"><%=s%>&nbsp;</p></td>
	</tr>
</table>

<!--  TELEFONE DO CLIENTE  -->
<table width="649" class="QS" cellSpacing="0">
	<tr>
<%	s = ""
	if Trim(cliente__tel_res) <> "" then
		s = telefone_formata(Trim(cliente__tel_res))
		s_aux=Trim(cliente__ddd_res)
		if s_aux<>"" then s = "(" & s_aux & ") " & s
		end if
	
	s2 = ""
	if Trim(cliente__tel_com) <> "" then
		s2 = telefone_formata(Trim(cliente__tel_com))
		s_aux = Trim(cliente__ddd_com)
		if s_aux<>"" then s2 = "(" & s_aux & ") " & s2
		s_aux = Trim(cliente__ramal_com)
		if s_aux<>"" then s2 = s2 & "  (R. " & s_aux & ")"
		end if

	s3 = ""
	if cliente__tipo = ID_PF then s3 = Trim(cliente__rg) else s3 = Trim(cliente__ie)
%>

<% if cliente__tipo = ID_PF then %>
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
<table width="649" class="QS" cellSpacing="0">
	<tr>
		<td><p class="Rf">E-MAIL</p><p class="C"><%=Trim(cliente__email)%>&nbsp;</p></td>
	</tr>
</table>

<!--  ENDEREÇO DE ENTREGA  -->
<%	
	s = pedido_formata_endereco_entrega(r_pedido, r_cliente)
%>		
<table width="649" class="QS" cellspacing="0" style="table-layout:fixed">
	<tr>
		<td align="left"><p class="Rf">ENDEREÇO DE ENTREGA</p><p class="C"><%=s%>&nbsp;</p></td>
	</tr>
    <%	if r_pedido.EndEtg_cod_justificativa <> "" then %>		
	<tr>
		<td align="left" style="word-wrap:break-word"><p class="C"><%=obtem_descricao_tabela_t_codigo_descricao(GRUPO_T_CODIGO_DESCRICAO__ENDETG_JUSTIFICATIVA,r_pedido.EndEtg_cod_justificativa)%>&nbsp;</p></td>
	</tr>
    <%end if %>
</table>

<!--  R E L A Ç Ã O   D E   P R O D U T O S  -->
<br>
<br>
<table class="Qx" cellSpacing="0">
	<tr bgColor="#FFFFFF">
	<td class="MB" valign="bottom"><p class="PLTe">Fabr</p></td>
	<td class="MB" valign="bottom"><p class="PLTe">Produto</p></td>
	<td class="MB" valign="bottom"><p class="PLTe">Descrição</p></td>
	<td class="MB" valign="bottom"><p class="PLTd">Qtd</p></td>
	<td class="MB" valign="bottom"><p class="PLTd">Falt</p></td>
	<% if blnTemRA Or ((r_pedido.permite_RA_status = 1) And (r_pedido.opcao_possui_RA = "S")) then %>
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
	 s_cor = "black"
	 if n <= Ubound(v_item) then
		with v_item(n)
			s_fabricante=.fabricante
			s_produto=.produto
			s_descricao=.descricao
			s_descricao_html=produto_formata_descricao_em_html(.descricao_html)
			s_qtde=.qtde
			s_preco_lista=formata_moeda(.preco_lista)
			if .desc_dado=0 then s_desc_dado="" else s_desc_dado=formata_perc_desc(.desc_dado)
			s_vl_unitario=formata_moeda(.preco_venda)
			if .preco_NF <> 0 then s_preco_NF=formata_moeda(.preco_NF) else s_preco_NF=""
			m_TotalItem=.qtde * .preco_venda
			m_TotalItemComRA=.qtde * .preco_NF
			s_vl_TotalItem=formata_moeda(m_TotalItem)
			m_TotalDestePedido=m_TotalDestePedido + m_TotalItem
			m_TotalDestePedidoComRA=m_TotalDestePedidoComRA + m_TotalItemComRA
			end with
		s_falta=""
		if Not IsPedidoEncerrado(r_pedido.st_entrega) then
			with v_disp(n)
				if .qtde_estoque_sem_presenca<>0 then s_falta=Cstr(.qtde_estoque_sem_presenca)
				s_cor = x_cor_item(.qtde, .qtde_estoque_vendido, .qtde_estoque_sem_presenca)
				end with
			end if
			
	 else
		s_fabricante=""
		s_produto=""
		s_descricao=""
		s_descricao_html=""
		s_qtde=""
		s_falta=""
		s_preco_lista=""
		s_desc_dado=""
		s_vl_unitario=""
		s_preco_NF=""
		s_vl_TotalItem=""
		end if
%>
	<% if (i > MIN_LINHAS_ITENS_IMPRESSAO_PEDIDO) And (s_produto = "") then %>
	<tr class="notPrint">
	<% else %>
	<tr>
	<% end if %>
	<td class="MDBE"><input name="c_fabricante" id="c_fabricante" class="PLLe" style="width:25px; color:<%=s_cor%>"
		value='<%=s_fabricante%>' readonly tabindex=-1></td>
	<td class="MDB"><input name="c_produto" id="c_produto" class="PLLe" style="width:54px; color:<%=s_cor%>"
		value='<%=s_produto%>' readonly tabindex=-1></td>
	<td class="MDB" style="width:269px;">
		<span class="PLLe" style="color:<%=s_cor%>"><%=s_descricao_html%></span>
		<input type="hidden" name="c_descricao" id="c_descricao" value='<%=s_descricao%>'>
	</td>
	<td class="MDB" align="right"><input name="c_qtde" id="c_qtde" class="PLLd" style="width:21px; color:<%=s_cor%>"
		value='<%=s_qtde%>' readonly tabindex=-1></td>
	<td class="MDB" align="right"><input name="c_qtde_falta" id="c_qtde_falta" class="PLLd" style="width:20px; color:<%=s_cor%>"
		value='<%=s_falta%>' readonly tabindex=-1></td>
	<% if blnTemRA Or ((r_pedido.permite_RA_status = 1) And (r_pedido.opcao_possui_RA = "S")) then %>
	<td class="MDB" align="right"><input name="c_vl_NF" id="c_vl_NF" class="PLLd" style="width:62px; color:<%=s_cor%>"
		value='<%=s_preco_NF%>' readonly tabindex=-1></td>
	<% end if %>
	<td class="MDB" align="right"><input name="c_preco_lista" id="c_preco_lista" class="PLLd" style="width:62px; color:<%=s_cor%>"
		value='<%=s_preco_lista%>' readonly tabindex=-1></td>
	<td class="MDB" align="right"><input name="c_desc" id="c_desc" class="PLLd" style="width:28px; color:<%=s_cor%>"
		value='<%=s_desc_dado%>' readonly tabindex=-1></td>
	<td class="MDB" align="right"><input name="c_vl_unitario" id="c_vl_unitario" class="PLLd" style="width:62px; color:<%=s_cor%>"
		value='<%=s_vl_unitario%>' readonly tabindex=-1></td>
	<td class="MDB" align="right"><input name="c_vl_total" id="c_vl_total" class="PLLd" style="width:70px; color:<%=s_cor%>" 
		value='<%=s_vl_TotalItem%>' readonly tabindex=-1></td>
	</tr>
<% next %>
	
	<tr>
	<td colspan="4">
		<table CellSpacing="0" CellPadding="0" width='100%' style="margin-top:4px;">
			<tr>
			<td width="20%">&nbsp;</td>
			<% if blnTemRA Or ((r_pedido.permite_RA_status = 1) And (r_pedido.opcao_possui_RA = "S")) then %>
			<td align="right">
				<table CellSpacing="0" CellPadding="0" style="margin-right:2px;">
					<tr>
					<td class="MTBE" NOWRAP><p class="PLTe">&nbsp;RA Líquido</p></td>
					<td class="MTBD" align="right"><input name="c_total_RA_Liquido" id="c_total_RA_Liquido" class="PLLd" style="width:70px;color:<%if r_pedido.vl_total_RA_liquido >=0 then Response.Write " green" else Response.Write " red"%>;" 
						value='<%=formata_moeda(r_pedido.vl_total_RA_liquido)%>' readonly tabindex=-1></td>
					</tr>
				</table>
			</td>
			<td align="right">
				<table CellSpacing="0" CellPadding="0" style="margin-right:2px;">
					<tr>
					<td class="MTBE" NOWRAP><p class="PLTe">&nbsp;RA Bruto</p></td>
					<td class="MTBD" align="right"><input name="c_total_RA" id="c_total_RA" class="PLLd" style="width:70px;color:<%if m_TotalFamiliaParcelaRA >=0 then Response.Write " green" else Response.Write " red"%>;" 
						value='<%=formata_moeda(m_TotalFamiliaParcelaRA)%>' READONLY tabindex=-1></td>
					</tr>
				</table>
			</td>
			<% end if %>
			<td align="right">
				<table CellSpacing="0" CellPadding="0">
				<tr>
				<td class="MTBE" NOWRAP><p class="PLTe">&nbsp;COM(%)</p></td>
				<td class="MTBD" align="right"><input name="c_perc_RT" id="c_perc_RT" class="PLLd" style="width:30px;color:blue;" 
					value='<%=formata_perc_RT(r_pedido.perc_RT)%>' READONLY tabindex=-1></td>
				</tr>
			</table>
			</td>
			</tr>
		</table>
	</td>
	<% if blnTemRA Or ((r_pedido.permite_RA_status = 1) And (r_pedido.opcao_possui_RA = "S")) then %>
	<td class="MD">&nbsp;</td>
	<td class="MDB" align="right">
		<input name="c_total_NF" id="c_total_NF" class="PLLd" style="width:70px;color:blue;" 
				value='<%=formata_moeda(m_TotalDestePedidoComRA)%>' READONLY tabindex=-1>
	</td>
	<td colspan="3" class="MD">&nbsp;</td>
	<% else %>
	<td colspan="4" class="MD">&nbsp;</td>
	<% end if %>
	<td class="MDB" align="right"><input name="c_total_geral" id="c_total_geral" class="PLLd" style="width:70px;color:blue;" 
		value='<%=formata_moeda(m_TotalDestePedido)%>' READONLY tabindex=-1></td>
	</tr>
</table>

<% if r_pedido.tipo_parcelamento = 0 then %>
<!--  TRATA VERSÃO ANTIGA DA FORMA DE PAGAMENTO   -->
<br>
<table class="Q" style="width:649px;" cellSpacing="0">
	<tr>
		<td class="MB" colspan="5"><p class="Rf">Observações </p>
			<textarea name="c_obs1" id="c_obs1" class="PLLe notPrint" rows="<%=Cstr(MAX_LINHAS_OBS1)%>" 
				style="width:642px;margin-left:2pt;" 
				readonly tabindex=-1><%=r_pedido.obs_1%></textarea>
			<span class="PLLe notVisible"><%
				s = substitui_caracteres(r_pedido.obs_1,chr(13),"<br>")
				if s = "" then s = "&nbsp;"
				Response.Write s %></span>
		</td>
	</tr>
	<tr>
		<td class="MB" colspan="5"><p class="Rf">Nº Nota Fiscal</p>
			<input name="c_obs2" id="c_obs2" class="PLLe" style="width:85px;margin-left:2pt;" 
				readonly tabindex=-1 value='<%=r_pedido.obs_2%>'>
		</td>
	</tr>
	<tr>
		<td class="MDB" NOWRAP width="10%"><p class="Rf">Parcelas</p>
			<input name="c_qtde_parcelas" id="c_qtde_parcelas" class="PLLc" style="width:60px;"
				readonly tabindex=-1 value='<%if (r_pedido.qtde_parcelas<>0) Or (r_pedido.forma_pagto<>"") then Response.write Cstr(r_pedido.qtde_parcelas)%>'>
		</td>
		<td class="MDB" NOWRAP valign="top"><p class="Rf">Entrega Imediata</p>
		<% 	if Cstr(r_pedido.st_etg_imediata) = Cstr(COD_ETG_IMEDIATA_NAO) then
				s = "NÃO"
			elseif Cstr(r_pedido.st_etg_imediata) = Cstr(COD_ETG_IMEDIATA_SIM) then
				s = "SIM"
			else
				s = ""
				end if
			
			if s <> "" then
				s_aux=formata_data_e_talvez_hora_hhmm(r_pedido.etg_imediata_data)
				if s_aux <> "" then s = s & " &nbsp; (" & iniciais_em_maiusculas(r_pedido.etg_imediata_usuario) & " - " & s_aux & ")"
				end if
			if s="" then s="&nbsp;"
		%>
		<p class="C" style="margin-top:3px;"><%=s%></p>
		</td>
		<td class="MDB" NOWRAP valign="top"><p class="Rf">Bem de Uso/Consumo</p>
		<% 	if Cstr(r_pedido.StBemUsoConsumo) = Cstr(COD_ST_BEM_USO_CONSUMO_NAO) then
				s = "NÃO"
			elseif Cstr(r_pedido.StBemUsoConsumo) = Cstr(COD_ST_BEM_USO_CONSUMO_SIM) then
				s = "SIM"
			else
				s = ""
				end if
			
			if s="" then s="&nbsp;"
		%>
		<p class="C" style="margin-top:3px;"><%=s%></p>
		</td>
		<td class="MDB" NOWRAP valign="top"><p class="Rf">Instalador Instala</p>
		<% 	if Cstr(r_pedido.InstaladorInstalaStatus) = Cstr(COD_INSTALADOR_INSTALA_NAO) then
				s = "NÃO"
			elseif Cstr(r_pedido.InstaladorInstalaStatus) = Cstr(COD_INSTALADOR_INSTALA_SIM) then
				s = "SIM"
			else
				s = ""
				end if
			
			if s="" then s="&nbsp;"
		%>
		<p class="C" style="margin-top:3px;"><%=s%></p>
		</td>
		<td class="MB" NOWRAP valign="top"><p class="Rf">Garantia Indicador</p>
		<% 	if Cstr(r_pedido.GarantiaIndicadorStatus) = Cstr(COD_GARANTIA_INDICADOR_STATUS__NAO) then
				s = "NÃO"
			elseif Cstr(r_pedido.GarantiaIndicadorStatus) = Cstr(COD_GARANTIA_INDICADOR_STATUS__SIM) then
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
				readonly tabindex=-1><%=r_pedido.forma_pagto%></textarea>
			<span class="PLLe notVisible"><%
				s = substitui_caracteres(r_pedido.forma_pagto,chr(13),"<br>")
				if s = "" then s = "&nbsp;"
				Response.Write s %></span>
		</td>
	</tr>
</table>
<% else %>
<!--  TRATA NOVA VERSÃO DA FORMA DE PAGAMENTO   -->
<br>
<table class="Q" style="width:649px;" cellSpacing="0">
	<tr>
		<td class="MB" colspan="6"><p class="Rf">Observações </p>
			<textarea name="c_obs1" id="c_obs1" class="PLLe notPrint" rows="<%=Cstr(MAX_LINHAS_OBS1)%>" 
				style="width:642px;margin-left:2pt;" 
				readonly tabindex=-1><%=r_pedido.obs_1%></textarea>
			<span class="PLLe notVisible"><%
				s = substitui_caracteres(r_pedido.obs_1,chr(13),"<br>")
				if s = "" then s = "&nbsp;"
				Response.Write s %></span>
		</td>
	</tr>
	<tr>
		<td class="MD" NOWRAP><p class="Rf">Nº Nota Fiscal</p>
			<input name="c_obs2" id="c_obs2" class="PLLe" style="width:67px;margin-left:2pt;" 
				readonly tabindex=-1 value='<%=r_pedido.obs_2%>'>
		</td>
		<td class="MD" NOWRAP><p class="Rf">NF Simples Remessa</p>
			<input name="c_obs3" id="c_obs3" class="PLLe" style="width:67px;margin-left:2pt;" 
				readonly tabindex=-1 value='<%=r_pedido.obs_3%>'>
		</td>
		<td class="MD" NOWRAP valign="top"><p class="Rf">Entrega Imediata</p>
		<% 	if Cstr(r_pedido.st_etg_imediata) = Cstr(COD_ETG_IMEDIATA_NAO) then
				s = "NÃO"
			elseif Cstr(r_pedido.st_etg_imediata) = Cstr(COD_ETG_IMEDIATA_SIM) then
				s = "SIM"
			else
				s = ""
				end if
			
			if s <> "" then
				s_aux=formata_data_e_talvez_hora_hhmm(r_pedido.etg_imediata_data)
				if s_aux <> "" then s = s & " &nbsp; (" & iniciais_em_maiusculas(r_pedido.etg_imediata_usuario) & " - " & s_aux & ")"
				end if
			if s="" then s="&nbsp;"
		%>
		<p class="C" style="margin-top:3px;"><%=s%></p>
		</td>
		<td class="MD" NOWRAP valign="top"><p class="Rf">Bem Uso/Consumo</p>
		<% 	if Cstr(r_pedido.StBemUsoConsumo) = Cstr(COD_ST_BEM_USO_CONSUMO_NAO) then
				s = "NÃO"
			elseif Cstr(r_pedido.StBemUsoConsumo) = Cstr(COD_ST_BEM_USO_CONSUMO_SIM) then
				s = "SIM"
			else
				s = ""
				end if
			
			if s="" then s="&nbsp;"
		%>
		<p class="C" style="margin-top:3px;"><%=s%></p>
		</td>
		<td class="MD" NOWRAP valign="top"><p class="Rf">Instalador Instala</p>
		<% 	if Cstr(r_pedido.InstaladorInstalaStatus) = Cstr(COD_INSTALADOR_INSTALA_NAO) then
				s = "NÃO"
			elseif Cstr(r_pedido.InstaladorInstalaStatus) = Cstr(COD_INSTALADOR_INSTALA_SIM) then
				s = "SIM"
			else
				s = ""
				end if
			
			if s="" then s="&nbsp;"
		%>
		<p class="C" style="margin-top:3px;"><%=s%></p>
		</td>
		<td NOWRAP valign="top"><p class="Rf">Garantia Indicador</p>
		<% 	if Cstr(r_pedido.GarantiaIndicadorStatus) = Cstr(COD_GARANTIA_INDICADOR_STATUS__NAO) then
				s = "NÃO"
			elseif Cstr(r_pedido.GarantiaIndicadorStatus) = Cstr(COD_GARANTIA_INDICADOR_STATUS__SIM) then
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
<table class="Q" style="width:649px;" cellSpacing="0">
  <tr>
	<td><p class="Rf">Forma de Pagamento</p></td>
  </tr>  
  <tr>
	<td>
	  <table width="100%" CellSpacing="0" CellPadding="0" border="0">
		<% if Cstr(r_pedido.tipo_parcelamento) = COD_FORMA_PAGTO_A_VISTA then %>
		<!--  À VISTA  -->
		<tr>
		  <td>
			<table CellSpacing="0" CellPadding="0" border="0">
			  <tr>
				<td><span class="C">À Vista&nbsp&nbsp;(<%=x_opcao_forma_pagamento(r_pedido.av_forma_pagto)%>)</span></td>
			  </tr>
			</table>
		  </td>
		</tr>
		<% elseif Cstr(r_pedido.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELA_UNICA then %>
		<!--  PARCELA ÚNICA  -->
		<tr>
		  <td>
			<table CellSpacing="0" CellPadding="0" border="0">
			  <tr>
				<td><span class="C">Parcela Única:&nbsp;&nbsp;<%=SIMBOLO_MONETARIO%>&nbsp;<%=formata_moeda(r_pedido.pu_valor)%>&nbsp;&nbsp;&nbsp;(<%=x_opcao_forma_pagamento(r_pedido.pu_forma_pagto)%>)&nbsp;&nbsp;vencendo após&nbsp;<%=formata_inteiro(r_pedido.pu_vencto_apos)%>&nbsp;dias</span></td>
			  </tr>
			</table>
		  </td>
		</tr>
		<% elseif Cstr(r_pedido.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_CARTAO then %>
		<!--  PARCELADO NO CARTÃO (INTERNET)  -->
		<tr>
		  <td>
			<table cellspacing="0" cellpadding="0" border="0">
			  <tr>
				<td><span class="C">Parcelado no Cartão (internet) em&nbsp;&nbsp;<%=Cstr(r_pedido.pc_qtde_parcelas)%>&nbsp;x&nbsp;&nbsp;<%=SIMBOLO_MONETARIO & " " & formata_moeda(r_pedido.pc_valor_parcela)%></span></td>
			  </tr>
			</table>
		  </td>
		</tr>
		<% elseif Cstr(r_pedido.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_CARTAO_MAQUINETA then %>
		<!--  PARCELADO NO CARTÃO (MAQUINETA)  -->
		<tr>
		  <td>
			<table cellspacing="0" cellpadding="0" border="0">
			  <tr>
				<td><span class="C">Parcelado no Cartão (maquineta) em&nbsp;&nbsp;<%=Cstr(r_pedido.pc_maquineta_qtde_parcelas)%>&nbsp;x&nbsp;&nbsp;<%=SIMBOLO_MONETARIO & " " & formata_moeda(r_pedido.pc_maquineta_valor_parcela)%></span></td>
			  </tr>
			</table>
		  </td>
		</tr>
		<% elseif Cstr(r_pedido.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_COM_ENTRADA then %>
		<!--  PARCELADO COM ENTRADA  -->
		<tr>
		  <td>
			<table CellSpacing="0" CellPadding="0" border="0">
			  <tr>
				<td><span class="C">Entrada:&nbsp;&nbsp;<%=SIMBOLO_MONETARIO%>&nbsp;<%=formata_moeda(r_pedido.pce_entrada_valor)%>&nbsp;&nbsp;&nbsp;(<%=x_opcao_forma_pagamento(r_pedido.pce_forma_pagto_entrada)%>)</span></td>
			  </tr>
			  <tr>
				<td><span class="C">Prestações:&nbsp;&nbsp;<%=formata_inteiro(r_pedido.pce_prestacao_qtde)%>&nbsp;x&nbsp;<%=SIMBOLO_MONETARIO%>&nbsp;<%=formata_moeda(r_pedido.pce_prestacao_valor)%>&nbsp;&nbsp;&nbsp;(<%=x_opcao_forma_pagamento(r_pedido.pce_forma_pagto_prestacao)%>)&nbsp;&nbsp;vencendo a cada&nbsp;<%=formata_inteiro(r_pedido.pce_prestacao_periodo)%>&nbsp;dias</span></td>
			  </tr>
			</table>
		  </td>
		</tr>
		<% elseif Cstr(r_pedido.tipo_parcelamento) = COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA then %>
		<!--  PARCELADO SEM ENTRADA  -->
		<tr>
		  <td>
			<table CellSpacing="0" CellPadding="0" border="0">
			  <tr>
				<td><span class="C">1ª Prestação:&nbsp;&nbsp;<%=SIMBOLO_MONETARIO%>&nbsp;<%=formata_moeda(r_pedido.pse_prim_prest_valor)%>&nbsp;&nbsp;&nbsp;(<%=x_opcao_forma_pagamento(r_pedido.pse_forma_pagto_prim_prest)%>)&nbsp;&nbsp;vencendo após&nbsp;<%=formata_inteiro(r_pedido.pse_prim_prest_apos)%>&nbsp;dias</span></td>
			  </tr>
			  <tr>
				<td><span class="C">Demais Prestações:&nbsp;&nbsp;<%=Cstr(r_pedido.pse_demais_prest_qtde)%>&nbsp;x&nbsp;<%=SIMBOLO_MONETARIO%>&nbsp;<%=formata_moeda(r_pedido.pse_demais_prest_valor)%>&nbsp;&nbsp;&nbsp;(<%=x_opcao_forma_pagamento(r_pedido.pse_forma_pagto_demais_prest)%>)&nbsp;&nbsp;vencendo a cada&nbsp;<%=Cstr(r_pedido.pse_demais_prest_periodo)%>&nbsp;dias</span></td>
			  </tr>
			</table>
		  </td>
		</tr>
		<% end if %>
	  </table>
	</td>
  </tr>
  <tr>
	<td class="MC"><p class="Rf">Informações Sobre Análise de Crédito</p>
	  <textarea name="c_forma_pagto" id="c_forma_pagto" class="PLLe notPrint" rows="<%=Cstr(MAX_LINHAS_FORMA_PAGTO)%>"
				style="width:642px;margin-left:2pt;"
				readonly tabindex=-1><%=r_pedido.forma_pagto%></textarea>
	  <span class="PLLe notVisible"><%
			s = substitui_caracteres(r_pedido.forma_pagto,chr(13),"<br>")
			if s = "" then s = "&nbsp;"
			Response.Write s %></span>
	</td>
  </tr>
</table>
<% end if %>


<!--  STATUS DE PAGAMENTO   -->
<br>
<table width="649" class="Q" cellSpacing="0">
<tr>
	<td width="16.67%" class="MD" valign="bottom"><p class="Rf">Status de Pagto</p></td>
	<td width="16.67%" class="MD" align="right" valign="bottom"><p class="Rf">VL Total&nbsp;&nbsp;(Família)&nbsp;</p></td>
	<td width="16.67%" class="MD" align="right" valign="bottom"><p class="Rf">VL Pago&nbsp;</p></td>
	<td width="16.67%" class="MD" align="right" valign="bottom"><p class="Rf">VL Devoluções&nbsp;</p></td>
	<td width="16.67%" class="MD" align="right" valign="bottom"><p class="Rf">VL Perdas&nbsp;</p></td>
	<td width="16.65%" align="right" valign="bottom"><p class="Rf">Saldo a Pagar&nbsp;</p></td>
</tr>
<tr>
	<% s_aux = x_status_pagto_cor(st_pagto) 
	   s = Ucase(x_status_pagto(st_pagto)) %>
	<td width="16.67%" class="MD"><p class="C" style="color:<%=s_aux%>;"><%=s%>&nbsp;</p></td>
	<% s = formata_moeda(vl_TotalFamiliaPrecoNF) %>
	<td width="16.67%" align="right" class="MD"><p class="Cd"><%=s%></p></td>
	<% s = formata_moeda(vl_TotalFamiliaPago) %>
	<td width="16.67%" align="right" class="MD"><p class="Cd" style="color:<%
		if vl_TotalFamiliaPago >= 0 then Response.Write "black" else Response.Write "red" 
		%>;"><%=s%></p></td>
	<% s = formata_moeda(vl_TotalFamiliaDevolucaoPrecoNF) %>
	<td width="16.67%" align="right" class="MD"><p class="Cd"><%=s%></p></td>
	<% s = formata_moeda(vl_total_perdas) %>
	<td width="16.67%" align="right" class="MD"><p class="Cd"><%=s%></p></td>
	<td width="16.65%" align="right"><p class="Cd" style="color:<% 
		if vl_saldo_a_pagar >= 0 then Response.Write "black" else Response.Write "red" 
		%>;"><%=s_vl_saldo_a_pagar%></p></td>
</tr>
</table>


<!--  ANÁLISE DE CRÉDITO   -->
<br>
<table width="649" class="Q" cellSpacing="0">
<tr>
	<%	s=x_analise_credito(r_pedido.analise_credito)
		if s <> "" then
			s_aux=formata_data_e_talvez_hora_hhmm(r_pedido.analise_credito_data)
			if Trim(r_pedido.analise_credito_usuario) <> "" then
				if s_aux <> "" then s_aux = s_aux & " - "
				s_aux = s_aux & iniciais_em_maiusculas(Trim(r_pedido.analise_credito_usuario))
				end if
			if s_aux <> "" then s = s & " &nbsp; (" & s_aux & ")"
			end if
		if s="" then s="&nbsp;"
	%>
	<td><p class="Rf">ANÁLISE DE CRÉDITO</p><p class="C" style="color:<%=x_analise_credito_cor(r_pedido.analise_credito)%>;"><%=s%></p></td>
</tr>
</table>


<% if s_devolucoes <> "" then %>
<!--  DEVOLUÇÕES   -->
<br>
<table width="649" class="Q" cellSpacing="0">
<tr>
	<td><p class="Rf" style="color:red;">DEVOLUÇÃO DE MERCADORIAS</p><p class="C"><%=s_devolucoes%></p></td>
</tr>
</table>
<% end if %>


<% if s_perdas <> "" then %>
<!--  PERDAS   -->
<br>
<table width="649" class="Q" cellSpacing="0">
<tr>
	<td><p class="Rf" style="color:red;">PERDAS</p><p class="C"><%=s_perdas%></p></td>
</tr>
</table>
<% end if %>


<% if IsEntregaAgendavel(r_pedido.st_entrega) then %>
<!--  DATA DE COLETA   -->
<br>
<table width="649" class="Q" cellSpacing="0">
<tr>
	<%	s=formata_data(r_pedido.a_entregar_data_marcada)
		if s="" then s="&nbsp;"
	%>
	<td><p class="Rf">DATA DE COLETA</p><p class="C"><%=s%></p></td>
</tr>
</table>
<% end if %>


<% if operacao_permitida(OP_CEN_PEDIDO_EXIBIR_DADOS_LOGISTICA, s_lista_operacoes_permitidas) then %>
<br>
<table width="649" class="Q" cellSpacing="0">
<tr>
	<td width="33%" class="MD" valign="bottom"><p class="Rf">Volumes</p></td>
	<td width="33%" class="MD" valign="bottom"><p class="Rf">Cubagem (m3)</p></td>
	<td width="34%" valign="bottom"><p class="Rf">Peso (kg)</p></td>
</tr>
<tr>
	<% s = formata_inteiro(total_volumes) %>
	<td width="33%" class="MD"><p class="C"><%=s%></p></td>
	<% s = formata_numero(total_cubagem, 2) %>
	<td width="33%" class="MD"><p class="C"><%=s%>&nbsp;</p></td>
	<% s = formata_numero(total_peso, 2) %>
	<td width="34%"><p class="C"><%=s%></p></td>
</tr>
</table>
<% end if %>


<% if r_pedido.transportadora_id <> "" then %>
<!--  TRANSPORTADORA   -->
<br>
<table width="649" class="Q" cellSpacing="0">
<tr>
	<%	s=formata_data_e_talvez_hora(r_pedido.transportadora_data)
		if s <> "" then s = s & " - "
		s = s & r_pedido.transportadora_id & " (" & x_transportadora(r_pedido.transportadora_id) & ")"
		if s="" then s="&nbsp;"
	%>
	<td class="MD"><p class="Rf">TRANSPORTADORA</p><p class="C"><%=s%></p></td>
	
	<%	if r_pedido.frete_status <> 0 then
			s = formata_moeda(r_pedido.frete_valor)
		else
			s = ""
			end if
		if s = "" then s = "&nbsp;"
	%>
	<td class="MD" align="right" style="width:65px;"><p class="Rf">FRETE (<%=SIMBOLO_MONETARIO%>)&nbsp;</p><p class="Cd"><%=s%></p></td>
	
	<%
		s=Trim(r_pedido.transportadora_num_coleta)
		if s="" then s="&nbsp;"
	%>
	<td class="MD" style="width:65px;"><p class="Rf">COLETA</p><p class="C"><%=s%></p></td>
	<%
		s=Trim(r_pedido.transportadora_contato)
		if s="" then s="&nbsp;"
	%>
	<td style="width:85px;"><p class="Rf">CONTATO</p><p class="C"><%=s%></p></td>
</tr>
</table>
<% end if %>


<!-- ************   SEPARADOR   ************ -->
<table class="notPrint" width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>

<table class="notPrint" width="649" cellPadding="0" CellSpacing="0">
<tr><td align="right"><span class="Rc">
	<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
	<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
	</span>
</td></tr>
</table>

<!-- ************   BOTÕES   ************ -->
<table class="notPrint" width="649" cellPadding="0" cellSpacing="0">
<tr>
	<td align="left"><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()" title="volta para página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a>
	</td>
	<td align="center"><div name="dIMPRESSORA" id="dIMPRESSORA">
		<a name="bIMPRESSORA" id="bIMPRESSORA" href="javascript:fPEDImpressora(fPED)" title="seleciona a impressora">
		<img src="../botao/impressora.gif" width="176" height="55" border="0"></a></div>
	</td>
	<td align="center"><div name="dMARGENS" id="dMARGENS">
		<a name="bMARGENS" id="bMARGENS" href="javascript:fPEDMargens(fPED)" title="configura as margens de impressão">
		<img src="../botao/margens.gif" width="176" height="55" border="0"></a></div>
	</td>
	<td align="right"><div name="dIMPRIME" id="dIMPRIME">
		<a name="bIMPRIME" id="bIMPRIME" href="javascript:fPEDImprime(fPED)" title="imprime o pedido em formulário contínuo">
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
