<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/Global.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ========================================================
'	  RelPedidoPreDevolucaoMercadoriaRecebe.asp
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
	if Not operacao_permitida(OP_CEN_PRE_DEVOLUCAO_RECEBIMENTO_MERCADORIA, s_lista_operacoes_permitidas) then  ' TODO
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if
	
	dim s_filtro, intQtdePreDevolucoes, intQtdePreDevolucoesItens
	dim s, rb_status, origem
    dim c_pedido, c_nota_fiscal, c_cpf_cnpj
	origem = ucase(Trim(request("origem")))
	intQtdePreDevolucoes = 0
    intQtdePreDevolucoesItens = 0

	if origem="A" then
	'	PARÂMETRO VEM PELA QUERYSTRING
		rb_status = Trim(Request("rb_status"))
        c_pedido = Trim(Request("c_pedido"))
        c_nota_fiscal = Trim(Request("c_nota_fiscal"))
        c_cpf_cnpj = retorna_so_digitos(Trim(Request("c_cpf_cnpj")))
	else
		rb_status = Trim(Request.Form("rb_status"))
        c_pedido = Trim(Request.Form("c_pedido"))
        c_nota_fiscal = Trim(Request.Form("c_nota_fiscal"))
        c_cpf_cnpj = retorna_so_digitos(Trim(Request.Form("c_cpf_cnpj")))
		end if

	dim alerta
	alerta = ""




' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________

function monta_link_pedido(byval id_pedido)
dim strLink
	monta_link_pedido = ""
	id_pedido = Trim("" & id_pedido)
	if id_pedido = "" then exit function
	strLink = "<a href='javascript:fPEDConsulta(" & _
				chr(34) & id_pedido & chr(34) & _
				")' title='clique para consultar o pedido " & id_pedido & "'>" & _
				id_pedido & "</a>"
	monta_link_pedido=strLink
end function

' ___________________________________
' OBTÉM DESCRIÇÃO STATUS DEVOLUÇÃO
'
function obtem_descricao_status_devolucao(byval st_codigo, byref st_devolucao_descricao, byref st_devolucao_cor)
dim s_descricao, s_cor
    st_codigo = Trim("" & st_codigo)
    if st_codigo = "" then exit function

    st_devolucao_descricao = ""
    st_devolucao_cor = ""

    s_descricao = ""
    s_cor = ""
    select case st_codigo
        case COD_ST_PEDIDO_DEVOLUCAO__CADASTRADA
                s_descricao = "Cadastrada"
                s_cor = "#0348E1"
            case COD_ST_PEDIDO_DEVOLUCAO__EM_ANDAMENTO
                s_descricao = "Em Andamento"
                s_cor = "#E26534"
            case COD_ST_PEDIDO_DEVOLUCAO__MERCADORIA_RECEBIDA
                s_descricao = "Mercadoria Recebida"
                s_cor = "#008080"
            case COD_ST_PEDIDO_DEVOLUCAO__FINALIZADA
                s_descricao = "Finalizada"
                s_cor = "#4FAB5B"
            case COD_ST_PEDIDO_DEVOLUCAO__REPROVADA
                s_descricao = "Reprovada"
                s_cor = "#B91832"
            case COD_ST_PEDIDO_DEVOLUCAO__CANCELADA
                s_descricao = "Cancelada"
                s_cor = "#C7465A"
            case else
                s_descricao = "Indefinido"
                s_cor = "#000000"
    end select
    st_devolucao_descricao = s_descricao
    st_devolucao_cor = s_cor
end function

' _____________________________________
' CONSULTA EXECUTA
'
sub consulta_executa
dim s, s2, s_aux, s_sql, x
dim r
dim cab_table, cab
dim id_devolucao
dim vlTotalItem, vlTotalDevolucao
dim st_devolucao_descricao, st_devolucao_cor

	s_sql = _
		"SELECT " & _
            " tPD.id AS id_devolucao," & _
            " tPD.pedido," & _
            " tPD.usuario_cadastro," & _
            " tPD.dt_hr_cadastro," & _
            " tPD.status," & _
            " tPD.cod_procedimento," & _
            " tPD.cod_devolucao_motivo," & _
            " tPD.vl_devolucao," & _
            " tPD.cod_credito_transacao," & _
            " tP.loja," & _
            " tP.data AS data_pedido," & _
            " tP.vendedor," & _
            " tP.transportadora_id," & _
            " tP.indicador," & _
            " tP.obs_2," & _
            " tP.obs_3," & _
            " tC.nome_iniciais_em_maiusculas AS nome_cliente" & _
        " FROM t_PEDIDO_DEVOLUCAO tPD" & _
        " INNER JOIN t_PEDIDO tP ON (tPD.pedido = tP.pedido)" & _
        " INNER JOIN t_CLIENTE tC ON (tP.id_cliente=tC.id)" & _
        " WHERE (" & _
		        "1 = 1" & _
		        ")"

	if rb_status = "CADASTRADA" then
		s_sql = s_sql & " AND (status = " & COD_ST_PEDIDO_DEVOLUCAO__CADASTRADA & ")"
	elseif rb_status = "EM_ANDAMENTO" then
		s_sql = s_sql & " AND (status = " & COD_ST_PEDIDO_DEVOLUCAO__EM_ANDAMENTO & ")"
	elseif rb_status = "MERCADORIA_RECEBIDA" then
		s_sql = s_sql & " AND (status = " & COD_ST_PEDIDO_DEVOLUCAO__MERCADORIA_RECEBIDA & ")"
		end if

    if c_pedido <> "" then
		s_sql = s_sql & " AND (tPD.pedido = '" & c_pedido & "')"
		end if

    if c_cpf_cnpj <> "" then
		s_sql = s_sql & " AND (tC.cnpj_cpf = '" & c_cpf_cnpj & "')"
		end if

    if c_nota_fiscal <> "" then
		s_sql = s_sql & " AND (tP.pedido IN (SELECT DISTINCT pedido FROM t_PEDIDO_ITEM_DEVOLVIDO WHERE (NFe_numero_NF = '" & c_nota_fiscal & "')))"
		end if

	cab_table = "<TABLE cellSpacing=0 cellPadding=0>" & chr(13)
	cab = "	<TR style='background:azure' NOWRAP>" & chr(13) & _
		  "		<TD class='MDTE tdDataHora' style='vertical-align:bottom'><P class='Rc'>DT Cadastro</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdDevolucaoID' style='vertical-align:bottom'><P class='R'>ID Devol</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdLoja' style='vertical-align:bottom'><P class='R'>Loja</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdPedido' style='vertical-align:bottom'><P class='R'>Pedido</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdNotaFiscal' style='vertical-align:bottom'><P class='R'>Nota Fiscal</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdVendedor' style='vertical-align:bottom'><P class='Rc'>Vendedor</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdIndicador' style='vertical-align:bottom'><P class='Rc'>Indicador</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdCliente' style='vertical-align:bottom'><P class='R'>Cliente</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdTransp' style='vertical-align:bottom'><P class='R'>Transp</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdMotivo' style='vertical-align:bottom'><P class='R'>Motivo</P></TD>" & chr(13) & _
          "     <TD class='MTD tdVlDevolucao' style='vertical-align:bottom'><P class='R'>VL Devolução</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdStatus' style='vertical-align:bottom'><P class='Rc'>Status</P></TD>" & chr(13) & _
		  "		<TD style='background:white;'>&nbsp;</TD>" & chr(13) & _
		  "	</TR>" & chr(13)
	
	x = cab_table & cab
	intQtdePreDevolucoes = 0
    intQtdePreDevolucoesItens = 0
	
	set r = cn.execute(s_sql)
	do while Not r.Eof
	
	 ' CONTAGEM
		intQtdePreDevolucoes = intQtdePreDevolucoes + 1
        id_devolucao = Trim("" & r("id_devolucao"))

        vlTotalDevolucao = 0

		x = x & "	<TR NOWRAP>" & chr(13)
		
	'> DATA DA PRÉ-DEVOLUÇÃO
		s = formata_data_hora_sem_seg(r("dt_hr_cadastro"))
		x = x & "		<TD class='MDTE tdDataHora'><P class='Cnc'>" & s & "</P></TD>" & chr(13)

    '> ID DEVOLUÇÃO
		x = x & "		<TD class='MTD tdDevolucaoID'><P class='Cn'>" & Trim("" & r("id_devolucao")) & "</P></TD>" & chr(13)

	'> LOJA
		x = x & "		<TD class='MTD tdLoja'><P class='Cn'>" & Trim("" & r("loja")) & "</P></TD>" & chr(13)

	'> PEDIDO
		s = monta_link_pedido(Trim("" & r("pedido")))
		x = x & "		<TD class='MTD tdPedido'><P class='Cn'>" & s & "</P></TD>" & chr(13)
    
    '> NOTA FISCAL
        s = Trim("" & r("obs_2"))
        if Trim("" & r("obs_3")) <> "" then s = s & ", " & Trim("" & r("obs_3"))
		x = x & "		<TD class='MTD tdNotaFiscal'><P class='Cn'>" & s & "</P></TD>" & chr(13)        

    '> VENDEDOR
		x = x & "		<TD class='MTD tdIndicador'><P class='Cnc'>" & Trim("" & r("vendedor")) & "</P></TD>" & chr(13)

	'> INDICADOR
		x = x & "		<TD class='MTD tdIndicador'><P class='Cnc'>" & Trim("" & r("indicador")) & "</P></TD>" & chr(13)

	'> CLIENTE
		s = Trim("" & r("nome_cliente"))
		if s = "" then s = "&nbsp;"
		x = x & "		<TD class='MTD tdCliente'><P class='Cn'>" & s & "</P></TD>" & chr(13)

	'> TRANSPORTADORA
		s = Trim("" & r("transportadora_id"))
		if s = "" then s = "&nbsp;"
		x = x & "		<TD class='MTD tdTransp'><P class='Cn'>" & s & "</P></TD>" & chr(13)

	'> MOTIVO DEVOLUÇÃO
		s = Trim("" & r("cod_devolucao_motivo"))
		if s <> "" then
            s = obtem_descricao_tabela_t_codigo_descricao(GRUPO_T_CODIGO_DESCRICAO__PEDIDO_DEVOLUCAO__MOTIVO,s)
        else 
            s = "&nbsp;"
            end if
		x = x & "		<TD class='MTD tdMotivo'><P class='Cn'>" & s & "</P></TD>" & chr(13)

    '> VALOR DA DEVOLUÇÃO
        s = formata_moeda(Trim("" & r("vl_devolucao")))
        x = x & "		<TD class='MTD tdVlDevolucao'><P class='Cn'>" & SIMBOLO_MONETARIO & " " & s & "</P></TD>" & chr(13)

	'> STATUS
        obtem_descricao_status_devolucao r("status"), st_devolucao_descricao, st_devolucao_cor

		x = x & "		<TD class='MTD tdStatus'><P class='Cnc' style='color:" & st_devolucao_cor & ";'>" & st_devolucao_descricao & "</P></TD>" & chr(13)

	'> BOTÃO P/ EXIBIR DEMAIS CAMPOS
		x = x & "		<TD valign='top' class='notPrint'>" & _
							"&nbsp;<a name='bExibeOcultaCampos' id='bExibeOcultaCampos' href='javascript:fExibeOcultaCampos(" & chr(34) & Cstr(intQtdePreDevolucoes) & chr(34) & ")' title='exibe ou oculta os campos adicionais'><img src='../botao/view_bottom.png' border='0'></a>" & _
						"</TD>" & chr(13)
		
		x = x & "	</TR>" & chr(13)

    '> ITENS DEVOLVIDOS
		s_sql = _
			"SELECT " & _
				"tPDI.fabricante," & _
				"tPDI.produto," & _
				"tPDI.qtde," & _
				"tPDI.qtde_estoque_venda," & _
				"tPDI.qtde_estoque_danificado," & _
				"tPDI.vl_unitario," & _
                "tPDI.id_pedido_devolucao," & _
				"tP.descricao," & _
                "tP.descricao_html" & _
		   " FROM t_PEDIDO_DEVOLUCAO_ITEM tPDI" & _
           " INNER JOIN t_PRODUTO tP ON ((tPDI.fabricante=tP.fabricante) AND (tPDI.produto=tP.produto))" & _
		   " WHERE" & _
				" (id_pedido_devolucao = " & Trim("" & r("id_devolucao")) & ")" & _
		   " ORDER BY" & _
				" tPDI.produto," & _
				" tPDI.fabricante"
		if rs.State <> 0 then rs.Close
		rs.open s_sql, cn
		x = x & "	<TR style='display:none;' id='TR_ITENS_" & Cstr(intQtdePreDevolucoes) & "'>" & chr(13) & _
				"		<TD class='ME MD'>&nbsp;</TD>" & chr(13) & _
				"		<TD colspan='11' class='MC MD'>" & chr(13) & _
				"			<table width='100%' cellspacing='0' cellpadding='0'>" & chr(13) & _
				"				<TR>" & chr(13) & _
				"					<td class='Rf tdWithPadding'>ITENS DEVOLVIDOS</td>" & chr(13) & _
				"				</TR>" & chr(13) & _
                "               <TR style='background:#FDF5EB' NOWRAP>" & chr(13) & _
                "                   <TD>" & chr(13) & _
                "                       <table width='100%' cellSpacing='0' cellPadding='0'>" & chr(13) & _
                "                           <TR>" & chr(13) & _
                "                               <TD class='MTD tdFabricante' style='vertical-align:bottom'><P class='Rc'>Fabricante</P></TD>" & chr(13) & _
                "                               <TD class='MTD tdProduto' style='vertical-align:bottom'><P class='Rc'>Produto</P></TD>" & chr(13) & _
                "                               <TD class='MTD tdProdutoDescricao' style='vertical-align:bottom; padding-left: 3px;'><P class='R'>Descrição</P></TD>" & chr(13) & _
                "                               <TD class='MTD tdQtde' style='vertical-align:bottom' align='right'><P class='R'>Qtde</P></TD>" & chr(13) & _
                "                               <TD class='MTD tdQtdeEstoque' style='vertical-align:bottom' align='right'><P class='R'>Estoque Venda</P></TD>" & chr(13) & _
                "                               <TD class='MTD tdQtdeEstoque' style='vertical-align:bottom' align='right'><P class='R'>Estoque Danificados</P></TD>" & chr(13) & _
                "                               <TD class='MC' style='vertical-align:bottom'><P class='R'>&nbsp;</P></TD>" & chr(13) & _
                "							</TR>" & chr(13) & _
				"						</table>" & chr(13) & _
				"					</TD>" & chr(13) & _
                "               </TR>" & chr(13)
		if rs.Eof then
			x = x & _
				"				<TR>" & chr(13) & _
				"					<td>&nbsp;</td>" & chr(13) & _
				"				</TR>" & chr(13)
			end if
		
		do while Not rs.Eof
            vlTotalItem = converte_numero(rs("vl_unitario"))*converte_numero(rs("qtde"))
            vlTotalDevolucao = vlTotalDevolucao+vlTotalItem
            intQtdePreDevolucoesItens=intQtdePreDevolucoesItens+1
            x = x & "           <input type='hidden' name='c_devolucao_id_" & Cstr(intQtdePreDevolucoesItens) & "' id='c_devolucao_id_" & Cstr(intQtdePreDevolucoesItens) & "' value='" & Trim("" & rs("id_pedido_devolucao")) & "'>" & chr(13)
            x = x & "           <input type='hidden' name='c_pedido_" & Cstr(intQtdePreDevolucoesItens) & "' id='c_pedido_" & Cstr(intQtdePreDevolucoesItens) & "' value='" & Trim("" & r("pedido")) & "'>" & chr(13)
			x = x & _
				"				<TR>" & chr(13) & _
				"					<TD>" & chr(13) & _
				"						<table width='100%' cellSpacing='0' cellPadding='0'>" & chr(13) & _
				"							<TR>" & chr(13) & _
				"								<TD class='Cn MD MC tdFabricante' align='center'>" & chr(13) & _
				"                                   <input type='text' name='c_fabricante_" & Cstr(intQtdePreDevolucoesItens) & "' id='c_fabricante_" & Cstr(intQtdePreDevolucoesItens) & "' class='PLLc' style='width:60px;' value='" & rs("fabricante") & "' readonly tabindex=-1>" & chr(13) & _
				"								</TD>" & chr(13) & _
				"								<TD class='Cn MD MC tdProduto' align='center'>" & chr(13) & _
				"                                   <input type='text' name='c_produto_" & Cstr(intQtdePreDevolucoesItens) & "' id='c_produto_" & Cstr(intQtdePreDevolucoesItens) & "' class='PLLc' style='width:65px;' value='" & rs("produto") & "' readonly tabindex=-1>" & chr(13) & _
				"								</TD>" & chr(13) & _
				"								<TD class='C MD MC tdProdutoDescricao' style='padding-left: 3px;' align='left' valign='top'>" & chr(13) & _
													rs("descricao_html") & _
                "                               </TD>" & chr(13) & _
                "								<TD class='C MD MC tdQtde' align='right'>" & chr(13) & _
                "                                   <input type='text' name='c_qtde_" & Cstr(intQtdePreDevolucoesItens) & "' id='c_qtde_" & Cstr(intQtdePreDevolucoesItens) & "' class='PLLd' style='width:35px' value='" & rs("qtde") & "' readonly tabindex=-1>" & chr(13) & _
                "                               </TD>" & chr(13) & _
                "								<TD class='Cn MD MC tdQtdeEstoque' align='right'>" & chr(13) & _
				"                                   <input type='text' maxlength='3' name='c_qtde_estoque_venda_" & Cstr(intQtdePreDevolucoesItens) & "' id='c_qtde_estoque_venda_" & Cstr(intQtdePreDevolucoesItens) & "' class='PLLd' style='width:35px;font-size:11pt;border:1px solid black;' value='" & rs("qtde_estoque_venda") & "' onkeypress='filtra_numerico();' onblur='fConsisteQtde(this, " & chr(34) & Cstr(intQtdePreDevolucoesItens) & chr(34) & ")'>" & chr(13) & _
                "                               </TD>" & chr(13) & _
                "								<TD class='Cn MD MC tdQtdeEstoque' align='right'>" & chr(13) & _
                "                                   <input type='text' maxlength='3' name='c_qtde_estoque_danificado_" & Cstr(intQtdePreDevolucoesItens) & "' id='c_qtde_estoque_danificado_" & Cstr(intQtdePreDevolucoesItens) & "' class='PLLd' style='width:35px;font-size:11pt;border:1px solid black;' value='" & rs("qtde_estoque_danificado") & "' onblur='fConsisteQtde(this, " & chr(34) & Cstr(intQtdePreDevolucoesItens) & chr(34) & ")'>" & chr(13) & _
				"								</TD>" & chr(13) & _
                "                               <TD class='MC'>&nbsp;</TD>" & chr(13) & _        
				"							</TR>" & chr(13) & _
				"						</table>" & chr(13) & _
				"					</TD>" & chr(13) & _
				"				</TR>" & chr(13)
			rs.MoveNext
			loop
			
		x = x & _
                "           </table>" & chr(13) & _
				"		</TD>" & chr(13) & _
				"	</TR>" & chr(13)
		
		if (intQtdePreDevolucoes mod 100) = 0 then
			Response.Write x
			x = ""
			end if
			
		r.MoveNext
		loop
	
	
'	TOTAL GERAL
	if intQtdePreDevolucoes > 0 then
		x = x & "	<TR>" & chr(13) & _
				"		<TD COLSPAN='12' class='MC' style='border-left:0px;border-right:0px;'>&nbsp;</TD>" & chr(13) & _
				"	</TR>" & chr(13) & _
				"	<TR NOWRAP style='background:honeydew'>" & chr(13) & _
				"		<TD COLSPAN='12' class='MT'><p class='C'>TOTAL: &nbsp; " & cstr(intQtdePreDevolucoes) & " pré-devoluções</p></TD>" & chr(13) & _
				"	</TR>" & chr(13)
		end if
	
  ' MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	if intQtdePreDevolucoes = 0 then
		x = cab_table & cab
		x = x & "	<TR NOWRAP>" & chr(13) & _
				"		<TD class='MT' colspan='12'><P class='ALERTA'>&nbsp;NENHUM REGISTRO ENCONTRADO&nbsp;</P></TD>" & chr(13) & _
				"	</TR>" & chr(13)
		end if

  ' FECHA TABELA
	x = x & "</TABLE>" & chr(13)
	
	x = x & "<input type=HIDDEN name='c_qtde_pre_devolucoes' id='c_qtde_pre_devolucoes' value='" & Cstr(intQtdePreDevolucoes) & "'>" & chr(13)
	x = x & "<input type=HIDDEN name='c_qtde_pre_devolucoes_itens' id='c_qtde_pre_devolucoes_itens' value='" & Cstr(intQtdePreDevolucoesItens) & "'>" & chr(13)

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
<script src="<%=URL_FILE__JQUERY%>" language="JavaScript" type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
window.status = 'Aguarde, executando a consulta...';

function fExibeOcultaCampos(indice_row) {
var row_ITENS;

	row_ITENS = document.getElementById("TR_ITENS_" + indice_row);

	if (row_ITENS.style.display.toString() == "none") {
	    row_ITENS.style.display = "";
	}
	else {
	    row_ITENS.style.display = "none";
	}
}

function fConsisteQtde(campo_atual, indice_row) {
var cq, cqev, cqed, iqev, iqed;
    cq = document.getElementById("c_qtde_" + indice_row);
    cqev = document.getElementById("c_qtde_estoque_venda_" + indice_row);
    cqed = document.getElementById("c_qtde_estoque_danificado_" + indice_row);

    iqev = converte_numero(cqev.value);
    iqed = converte_numero(cqed.value);
    
    if ((iqev + iqed) > parseInt(cq.value)) {
        alert("Quantidade atribuída ao(s) estoque(s) é superior a quantidade devolvida!!");
        campo_atual.value = 0;
        campo_atual.focus();
    }
}

function fPEDConsulta(id_pedido) {
    window.status = "Aguarde ...";
    fPED.pedido_selecionado.value = id_pedido;
    fPED.action = "pedido.asp"
    fPED.submit();
}
function fRELConfirma(f) {
var cq, cqv, cqd, n, cp, cqi, i, cdi, cp, ci, tt, p_a;
    cqi = document.getElementById("c_qtde_pre_devolucoes_itens");
    n = parseInt(cqi.value);
    for (i = 1; i <= n; i++) {
        cq = converte_numero(document.getElementById("c_qtde_" + i.toString()).value);
        cqv = converte_numero(document.getElementById("c_qtde_estoque_venda_" + i.toString()).value);
        cqd = converte_numero(document.getElementById("c_qtde_estoque_danificado_" + i.toString()).value);
        cdi = document.getElementById("c_devolucao_id_" + i.toString());
        cp = document.getElementById("c_pedido_" + i.toString());
        ci = document.getElementById("c_produto_" + i.toString());
        tt = cqv + cqd;
        if (tt > 0) {
            if (tt > cq) {
                alert("Quantidade atribuída ao estoque referente o produto " + ci.value + " do pedido " + cp.value + " é superior a quantidade devolvida!!");
                return;
            }
            else if (tt < cq) {
                alert("Quantidade atribuída ao estoque referente o produto " + ci.value + " do pedido " + cp.value + " é inferior a quantidade devolvida!!");
                return;
            }
        }
    }

    dCONFIRMA.style.visibility = "hidden";
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

<link href="<%=URL_FILE__E_CSS%>" Rel="stylesheet" Type="text/css">
<link href="<%=URL_FILE__EPRINTER_CSS%>" Rel="stylesheet" Type="text/css" media="print">
<link href="<%=URL_FILE__ESCREEN_CSS%>" Rel="stylesheet" Type="text/css" media="screen">

<style type="text/css">
html 
{
	overflow-y: scroll;
}
.tdWithPadding
{
	padding:1px;
}
.tdDataHora{
	vertical-align: top;
	width: 65px;
	}
.tdVendedor{
    vertical-align: top;
    width: 80px;
}
.tdLoja{
	vertical-align: top;
	text-align:center;
	font-weight: bold;
	width: 30px;
	}
.tdDevolucaoID{
	vertical-align: top;
	text-align:center;
	width: 40px;
	}
.tdPedido{
	vertical-align: top;
	font-weight: bold;
	width: 65px;
	}
.tdNotaFiscal{
	vertical-align: top;
	width: 65px;
	}
.tdIndicador{
	vertical-align: top;
	width: 80px;
	}
.tdCliente{
	vertical-align: top;
	width: 140px;
	}
.tdTransp{
	vertical-align: top;
	width: 80px;
	}
.tdContato{
	vertical-align: top;
	width: 100px;
	}
.tdTel{
	vertical-align: top;
	width: 90px;
	}
.tdMotivo{
	vertical-align: top;
	width: 260px;
	}
.tdVlDevolucao{
    vertical-align: top;
    width: 80px;
    text-align: right;
}
.tdStatus{
	vertical-align: top;
	width: 90px;
	}
.tdDataHoraMsg{
	vertical-align: top;
	width: 63px;
	}
.tdUsuarioMsg{
	vertical-align: top;
	width: 70px;
	}
.tdTextoMensagem{
	vertical-align: top;
	width: 785px;
	}
.tdFabricante{
    vertical-align: top;
    width: 63px;
}
.tdProduto{
    vertical-align: top;
    width: 70px;
}
.tdProdutoDescricao{
    vertical-align: top;
    width: 340px;
}
.tdQtde {
    vertical-align: top;
    width: 40px;
}
.tdQtdeEstoque{
    vertical-align: top;
    width: 70px;
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
<body onload="window.status='Concluído';focus();" link=#000000 alink=#000000 vlink=#000000>
<center>

<form id="fPED" name="fPED" method="post">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="pedido_selecionado" id="pedido_selecionado" value="">
</form>


<form id="fREL" name="fREL" method="post" action="RelPedidoPreDevolucaoMercadoriaRecebeGravaDados.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="rb_status" id="rb_status" value="<%=rb_status%>">
<input type="hidden" name="c_id_devolucao" id="c_id_devolucao" value="" />
<input type="hidden" name="pedido_selecionado" id="pedido_selecionado" value="">


<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="1024" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Registrar Mercadoria Recebida</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>

<!-- FILTROS -->
<% 
	s_filtro = "<table width='1024' cellPadding='0' CellSpacing='0' style='border-bottom:1px solid black' border='0'>" & chr(13)

	s = rb_status
	if s = "CADASTRADA" then
		s = "Cadastrada"
	elseif s = "EM_ANDAMENTO" then
		s = "Em Andamento"
	elseif s = "MERCADORIA_RECEBIDA" then
		s = "Mercadoria Recebida"
    elseif s = "FINALIZADA" then
		s = "Finalizada"
    elseif s = "REPROVADA" then
		s = "Reprovada"
    elseif s = "CANCELADA" then
		s = "Cancelada"
    elseif s = "" then
		s = "Todos"
	else
		s = "Parâmetro Desconhecido"
		end if

	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' NOWRAP>" & _
				"<p class='N'>Status da Pré-Devolução:&nbsp;</p></td><td valign='top'>" & _
				"<p class='N'>" & s & "</p></td>" & chr(13) & _
				"	</tr>" & chr(13)

    s = c_pedido
    if s <> "" then
        s_filtro = s_filtro & _
				    "	<tr>" & chr(13) & _
				    "		<td align='right' valign='top' NOWRAP>" & _
				    "<p class='N'>Pedido:&nbsp;</p></td><td valign='top'>" & _
				    "<p class='N'>" & s & "</p></td>" & chr(13) & _
				    "	</tr>" & chr(13)
        end if
    
    s = c_nota_fiscal
    if s <> "" then
        s_filtro = s_filtro & _
				    "	<tr>" & chr(13) & _
				    "		<td align='right' valign='top' NOWRAP>" & _
				    "<p class='N'>Nº Nota Fiscal:&nbsp;</p></td><td valign='top'>" & _
				    "<p class='N'>" & s & "</p></td>" & chr(13) & _
				    "	</tr>" & chr(13)
        end if

    s = c_cpf_cnpj
    if s <> "" then
        s = cnpj_cpf_formata(s)
        s_filtro = s_filtro & _
				    "	<tr>" & chr(13) & _
				    "		<td align='right' valign='top' NOWRAP>" & _
				    "<p class='N'>CPF/CNPJ Cliente:&nbsp;</p></td><td valign='top'>" & _
				    "<p class='N'>" & s & "</p></td>" & chr(13) & _
				    "	</tr>" & chr(13)
        end if

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
<table width="1024" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>
<br>


<table class="notPrint" width="1024" cellSpacing="0">
<tr>
	<td><a name="bVOLTAR" id="bVOLTAR"
		<% if origem="A" then %>
			href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>"
		<% else %>
			href="javascript:history.back()"
		<% end if %>
	title="volta para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
	<td align="right"><div name="dCONFIRMA" id="dCONFIRMA"><a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fRELConfirma(fREL)" title="executa a consulta">
		<img src="../botao/confirmar.gif" width="176" height="55" border="0"></a></div>
	</td>
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
