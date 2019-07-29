<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/Global.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ======================================================
'	  RelPedidosMktplaceRecebidosExec.asp
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
	if Not operacao_permitida(OP_CEN_REL_REGISTRO_PEDIDO_MARKETPLACE_RECEBIDO_CLIENTE, s_lista_operacoes_permitidas) then 
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	dim url_back, strUrlBotaoVoltar
	url_back = Trim(Request("url_back"))
	if url_back <> "" then
		strUrlBotaoVoltar = "RelPedidosMktplaceRecebidos.asp?url_back=X&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))
	else
		strUrlBotaoVoltar = "javascript:history.back()"
		end if

	dim alerta
	dim s, s_aux, s_filtro
	dim c_transportadora, c_loja, c_grupo_pedido_origem, c_pedido_origem
	dim s_nome_loja, s_nome_transportadora, s_nome_grupo_pedido_origem, s_nome_pedido_origem
	dim qtde_total_pedidos

	alerta = ""

	c_transportadora = Trim(Request.Form("c_transportadora"))
	c_loja = retorna_so_digitos(Trim(Request.Form("c_loja")))
    c_grupo_pedido_origem = Trim(Request.Form("c_grupo_pedido_origem"))
    c_pedido_origem = Trim(Request.Form("c_pedido_origem"))

	if alerta = "" then
		s_nome_transportadora = ""
		if c_transportadora <> "" then
			s = "SELECT nome FROM t_TRANSPORTADORA WHERE (id='" & c_transportadora & "')"
			if rs.State <> 0 then rs.Close
			rs.open s, cn
			if rs.Eof then 
				alerta=texto_add_br(alerta)
				alerta = "TRANSPORTADORA " & c_transportadora & " NÃO ESTÁ CADASTRADA."
			else
				s_nome_transportadora = iniciais_em_maiusculas(Trim("" & rs("nome")))
				end if
			end if
		end if

    if alerta = "" then
		s_nome_grupo_pedido_origem = ""
		if c_grupo_pedido_origem <> "" then
			s = "SELECT descricao FROM t_CODIGO_DESCRICAO WHERE (grupo='PedidoECommerce_Origem_Grupo' AND codigo='" & c_grupo_pedido_origem & "')"
			if rs.State <> 0 then rs.Close
			rs.open s, cn
			if rs.Eof then 
				alerta=texto_add_br(alerta)
				alerta = "ORIGEM DO PEDIDO (GRUPO) " & c_grupo_pedido_origem & " NÃO ESTÁ CADASTRADO."
			else
				s_nome_grupo_pedido_origem = Trim("" & rs("descricao"))
				end if
			end if
		end if

    if alerta = "" then
		s_nome_pedido_origem = ""
		if c_pedido_origem <> "" then
			s = "SELECT descricao FROM t_CODIGO_DESCRICAO WHERE (grupo='PedidoECommerce_Origem' AND codigo='" & c_pedido_origem & "')"
			if rs.State <> 0 then rs.Close
			rs.open s, cn
			if rs.Eof then 
				alerta=texto_add_br(alerta)
				alerta = "ORIGEM DO PEDIDO (GRUPO) " & c_pedido_origem & " NÃO ESTÁ CADASTRADO."
			else
				s_nome_pedido_origem = Trim("" & rs("descricao"))
				end if
			end if
		end if
		
	if alerta = "" then
		s_nome_loja = ""
		if c_loja <> "" then
			s = "SELECT * FROM t_LOJA WHERE (CONVERT(smallint,loja) = " & c_loja & ")"
			if rs.State <> 0 then rs.Close
			rs.open s, cn
			if rs.Eof then 
				alerta=texto_add_br(alerta)
				alerta = "LOJA " & c_loja & " NÃO ESTÁ CADASTRADA."
			else
				s_nome_loja = iniciais_em_maiusculas(Trim("" & rs("nome")))
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
dim x
dim cab, cab_table
dim s_sql, s_where
dim n_reg
dim strTransportadora, strTransportadoraAnterior, strTransportadoraAux, strPlural, strObs2, strObs3
dim strCidade, strUf, strCidadeUf
dim intQtdeTotalPedidos, intQtdeTransportadoras
dim intQtdeSubTotalPedidos, s_grupo_origem

'	CRITÉRIOS DE RESTRIÇÃO
	s_where = "(p.st_entrega = '" & ST_ENTREGA_ENTREGUE & "')" & _
			  " AND (p.MarketplacePedidoRecebidoRegistradoStatus = " & COD_ST_PEDIDO_RECEBIDO_NAO & ")" & _
			  " AND (p.MarketplacePedidoRecebidoRegistrarStatus = " & COD_ST_PEDIDO_RECEBIDO_SIM & ")" & _
              " AND (p.marketplace_codigo_origem IS NOT NULL) AND (LEN(Coalesce(p.marketplace_codigo_origem,'')) > 0)"
	
'	FILTRO: TRANSPORTADORA
	if c_transportadora <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (p.transportadora_id = '" & c_transportadora & "')"
		end if

'   FILTRO: ORIGEM DO PEDIDO (GRUPO)
    s = ""
    if c_grupo_pedido_origem <> "" then
        s_grupo_origem = "SELECT codigo FROM t_CODIGO_DESCRICAO WHERE (codigo_pai = '" & c_grupo_pedido_origem & "') AND grupo='PedidoECommerce_Origem'"
        if rs.State <> 0 then rs.Close
	    rs.open s_grupo_origem, cn
		if rs.Eof then
            alerta = "ORIGEM DO PEDIDO (GRUPO) " & c_grupo_pedido_origem & " NÃO EXISTE."
        else
            do while Not rs.Eof
                if s <> "" then s = s & ", "
                s = s & "'" & rs("codigo") & "'"      
                rs.MoveNext
            loop
             s = " p.marketplace_codigo_origem IN (" & s & ")"
        end if
        if s <> "" then 
			if s_where <> "" then s_where = s_where & " AND"
			s_where = s_where & " (" & s & ")"
		end if
    end if

'	FILTRO: ORIGEM DO PEDIDO
	if c_pedido_origem <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (p.marketplace_codigo_origem = '" & c_pedido_origem & "')"
		end if

'	FILTRO: LOJA
	if c_loja <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (p.numero_loja = " & c_loja & ")"
		end if

'	MONTA SQL DE CONSULTA
	s_sql = "SELECT" & _
				" p.transportadora_id," & _
				" p.pedido," & _
                " p.pedido_bs_x_ac," & _
                " p.pedido_bs_x_marketplace," & _
                " p.marketplace_codigo_origem," & _
				" p.loja," & _
                " p.MarketplacePedidoRecebidoRegistrarDataRecebido," & _
				" c.cidade," & _
				" c.uf," & _
				" c.nome_iniciais_em_maiusculas," & _
                " Sum(tPI.qtde*tPI.preco_venda) AS vl_pedido" & _
			" FROM t_PEDIDO p" &_
            " INNER JOIN t_PEDIDO_ITEM tPI ON (p.pedido=tPI.pedido)" & _
            " INNER JOIN t_CLIENTE c ON (p.id_cliente=c.id)" & _
			" WHERE " & _
				s_where & _
            " GROUP BY p.transportadora_id" & _
	           " ,p.pedido" & _
	           " ,p.pedido_bs_x_ac" & _
	           " ,p.pedido_bs_x_marketplace" & _
	           " ,p.marketplace_codigo_origem" & _
	           " ,p.loja" & _
               " ,p.MarketplacePedidoRecebidoRegistrarDataRecebido" & _
	           " ,c.cidade" & _
	           " ,c.uf" & _
	           " ,c.nome_iniciais_em_maiusculas" & _
			" ORDER BY" & _
				" p.transportadora_id," & _
				" p.pedido"
	
	
  ' CABEÇALHO
	cab_table = "<table cellspacing=0 cellpadding=0 class='MB'>" & chr(13)
	
	cab = _
		"	<tr style='background:azure' nowrap>" & chr(13) & _
		"		<td class='MDTE tdPedido' valign='bottom'><span class='R'>Nº Pedido</span></td>" &  chr(13) & _
		"		<td class='MTD tdPedido' valign='bottom'><span class='Rc'>Nº Magento</span></td>" &  chr(13) & _
		"		<td class='MTD tdPedido' valign='bottom'><span class='R' style='text-align:left;'>Nº Marketplace</span></td>" &  chr(13) & _
		"		<td class='MTD tdMarketplace' valign='bottom'><span class='R' style='text-align:left;'>Marketplace</span></td>" &  chr(13) & _
		"		<td class='MTD tdCliente' valign='bottom'><span class='R' style='text-align:left;'>Cliente</span></td>" &  chr(13) & _
		"		<td class='MTD tdPedido' align='right' valign='bottom'><span class='Rc'>Valor</span></td>" &  chr(13) & _
		"		<td class='MTD tdDataEntrega' align='center' valign='bottom'><span style='font-weight:bold; text-align:left;' class='R'>Recebido</span></td>" & chr(13) & _
        "		<td class='MTD tdRecebido' align='center' valign='bottom' nowrap><span class='Rc'>&nbsp;</span></td>" &  chr(13) & _
		"	</tr>" & chr(13)
	
'	LAÇO P/ LEITURA DO RECORDSET
	x = cab_table
	
	n_reg = 0
	intQtdeTotalPedidos = 0
	intQtdeTransportadoras = 0
	
	strTransportadoraAnterior = "XXXXXXXXXXXXXXXXXXXX"
	
	set r = cn.execute(s_sql)
	do while Not r.Eof

		strTransportadora = Trim("" & r("transportadora_id"))
		if strTransportadora <> strTransportadoraAnterior then
			intQtdeTransportadoras = intQtdeTransportadoras + 1
		'	SUB-TOTAIS POR TRANSPORTADORA
		'	EXIBE SUB-TOTAL DA TRANSPORTADORA ANTERIOR?
			if intQtdeTotalPedidos > 0 then
				if intQtdeSubTotalPedidos > 1 then strPlural = "s" else strPlural = ""
				x = x & _
					"	<tr style='background:ivory;'>" & chr(13) & _
					"		<td class='MDTE' colspan='8'>" & _
								"<span class='C' style='text-align:left;'>" & formata_inteiro(intQtdeSubTotalPedidos) & " pedido" & strPlural & "</span>" & _
					"		</td>" & chr(13) & _
					"	</tr>" & chr(13)
				end if
			
			strTransportadoraAux = strTransportadora
			if strTransportadoraAux = "" then strTransportadoraAux = "SEM TRANSPORTADORA"
			
			if intQtdeTotalPedidos > 0 then
			x = x & _
					"	<tr>" & chr(13) & _
					"		<td colspan='8' class='MC'>&nbsp;</td>" & chr(13) & _
					"	</tr>" & chr(13)
				end if
				
			x = x & _
				"	<tr style='background:azure'>" & chr(13) & _
				"		<td colspan='8' class='MC ME MD'><span class='C'>" & strTransportadoraAux & "</span></td>" & chr(13) & _
				"	</tr>" & chr(13)
			
		'	TÍTULO DAS COLUNAS
			x = x & _
				cab
			
			intQtdeSubTotalPedidos = 0
			strTransportadoraAnterior = strTransportadora
			end if
	
		n_reg = n_reg + 1
		intQtdeTotalPedidos = intQtdeTotalPedidos + 1
		intQtdeSubTotalPedidos = intQtdeSubTotalPedidos + 1

		x = x & "	<tr onmouseover='realca_cor_mouse_over(this);' onmouseout='realca_cor_mouse_out(this);'>" & chr(13)

    '>  Nº PEDIDO
	x = x & "		<td class='MDTE tdPedido'>" & _
						"<span class='Cc'>" & _
							"<a href='javascript:fPEDConsulta(" & chr(34) & r("pedido") & chr(34) & "," & chr(34) & usuario & chr(34) & ")' title='clique para consultar o pedido'>" & _
							Trim("" & r("pedido")) & _
							"</a>" & _
						"</span>" & _
					"</td>" & chr(13)

    '>  Nº MAGENTO
	x = x & "		<td class='MTD tdPedido'>" & _
						"<span class='Cc'>" & _
							"<a href='javascript:fPEDConsulta(" & chr(34) & r("pedido") & chr(34) & "," & chr(34) & usuario & chr(34) & ")' title='clique para consultar o pedido'>" & _
							Trim("" & r("pedido_bs_x_ac")) & _
							"</a>" & _
						"</span>" & _
					"</td>" & chr(13)

    '>  Nº MARKETPLACE
	x = x & "		<td class='MTD tdPedido'>" & _
						"<span class='Cc'>" & _
							"<a href='javascript:fPEDConsulta(" & chr(34) & r("pedido") & chr(34) & "," & chr(34) & usuario & chr(34) & ")' title='clique para consultar o pedido'>" & _
							Trim("" & r("pedido_bs_x_marketplace")) & _
							"</a>" & _
						"</span>" & _
					"</td>" & chr(13)

    '>  MARKETPLACE
	x = x & "		<td class='MTD tdMarketplace'>" & _
						"<span class='Cc'>" & _
							"<a href='javascript:fPEDConsulta(" & chr(34) & r("pedido") & chr(34) & "," & chr(34) & usuario & chr(34) & ")' title='clique para consultar o pedido'>" & _
							obtem_descricao_tabela_t_codigo_descricao("PedidoECommerce_Origem", Trim("" & r("marketplace_codigo_origem"))) & _
							"</a>" & _
						"</span>" & _
					"</td>" & chr(13)

    '>  NOME DO CLIENTE
		x = x & "		<td class='MTD tdCliente'>" & _
							"<span class='C' style='text-align:left;'>" & _
								Trim("" & r("nome_iniciais_em_maiusculas")) & _
							"</span>" & _
						"</td>" & chr(13)
		
	'>  VALOR
		x = x & "		<td class='MTD tdDataEntrega' align='right'>" & _
							"<span class='C' style='text-align:left;'>" & _
								formata_moeda(r("vl_pedido")) & _
							"</span>" & _
						"</td>" & chr(13)

	'>  DATA RECEBIMENTO
		x = x & "		<td class='MTD tdDataEntrega' align='center'>" & _
							"<input type='text' class='Cc cDtColeta' style='border:0;width:70px;' name='c_dt_recebimento' id='c_dt_recebimento' " & _
							"value = '" & formata_data(r("MarketplacePedidoRecebidoRegistrarDataRecebido")) & "' readonly" & _
							">" & _
						"</td>" & chr(13)

    '>  CHECKBOX
		x = x & "		<td class='MTD tdCkb tdRecebido' align='center'>" & _
							"<input type='checkbox' name='ckb_recebido' id='ckb_recebido' class='Cc CKB_REC' " & _
								" value='" & Trim("" & r("pedido")) & "'" & _
								">" & _
						"</td>" & chr(13)

		x = x & "	</tr>" & chr(13)
		
		if (n_reg mod 100) = 0 then
			Response.Write x
			x = ""
			end if
		
		r.MoveNext
		loop

	qtde_total_pedidos = intQtdeTotalPedidos

  ' MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	if intQtdeTotalPedidos = 0 then
		x = cab_table & _
			cab & _
			"	<tr nowrap>" & chr(13) & _
			"		<td class='MC MD ME ALERTA' colspan='8' align='center'><span class='ALERTA'>&nbsp;NENHUM PEDIDO ENCONTRADO&nbsp;</span></td>" & chr(13) & _
			"	</tr>" & chr(13)
	else
	'	SUB-TOTAL DA ÚLTIMA TRANSPORTADORA
		if intQtdeSubTotalPedidos > 1 then strPlural = "s" else strPlural = ""
		x = x & _
			"	<tr style='background:ivory;'>" & chr(13) & _
			"		<td class='MDTE' colspan='8' align='left'>" & _
						"<span class='C' style='text-align:left;'>" & formata_inteiro(intQtdeSubTotalPedidos) & " pedido" & strPlural & "</span>" & _
			"		</td>" & chr(13) & _
			"	</tr>" & chr(13)
		
	'	TOTAL GERAL
		if intQtdeTransportadoras > 1 then
			if intQtdeTotalPedidos > 1 then strPlural = "s" else strPlural = ""
			x = x & _
				"	<tr>" & chr(13) & _
				"		<td colspan='8' class='MC' align='left'>&nbsp;</td>" & chr(13) & _
				"	</tr>" & chr(13) & _
				"	<tr>" & chr(13) & _
				"		<td colspan='8' align='left'><span class='C' style='text-align:left;'>TOTAL GERAL</span></td>" & chr(13) & _
				"	</tr>" & chr(13) & _
				"	<tr style='background:ivory;'>" & chr(13) & _
				"		<td class='MDTE' colspan='8' align='left'>" & _
							"<span class='C' style='text-align:left;'>" & formata_inteiro(intQtdeTotalPedidos) & " pedido" & strPlural & "</span>" & _
				"		</td>" & chr(13) & _
				"	</tr>" & chr(13)
			end if
		end if

  ' FECHA TABELA
	x = x & "</table>" & chr(13)
	
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


<%=DOCTYPE_LEGADO%>

<html>


<head>
	<title>CENTRAL</title>
	</head>



<% if False then 'APENAS P/ HABILITAR O INTELLISENSE DURANTE O DESENVOLVIMENTO!! %>
<script src="../Global/jquery.js" language="JavaScript" type="text/javascript"></script>
<% end if %>

<script src="<%=URL_FILE__JQUERY%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_I18N%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>

<script type="text/javascript">
    $(document).ready(function () {

        $('#bMarcarTodos').click(function () {
            $('input[name=ckb_recebido]')
                .prop('checked', true)
                .parents("td.tdCkb").addClass("CKB_HIGHLIGHT");
        });

        $('#bDesmarcarTodos').click(function () {
            $('input[name=ckb_recebido]')
                .prop('checked', false)
                .parents("td.tdCkb").removeClass("CKB_HIGHLIGHT");
        });

		$("#divPedidoConsulta").hide();

		sizeDivPedidoConsulta();

		$('#divInternoPedidoConsulta').addClass('divFixo');

		$(document).keyup(function(e) {
			if (e.keyCode == 27) fechaDivPedidoConsulta();
		});

		$("#divPedidoConsulta").click(function() {
			fechaDivPedidoConsulta();
		});

		$("#imgFechaDivPedidoConsulta").click(function() {
			fechaDivPedidoConsulta();
		});

		// EXIBE O REALCE NOS CHECKBOXES QUE SÃO EXIBIDOS INICIALMENTE ASSINALADOS
		$(".CKB_REC:enabled").each(function() {
			if ($(this).is(":checked")) {
				$(this).parents("td.tdCkb").addClass("CKB_HIGHLIGHT");
			}
			else {
				$(this).parents("td.tdCkb").removeClass("CKB_HIGHLIGHT");
			}
		})

		// EVENTO P/ REALÇAR OU NÃO CONFORME SE MARCA/DESMARCA O CHECKBOX
		$(".CKB_REC:enabled").click(function() {
			if ($(this).is(":checked")) {
				$(this).parents("td.tdCkb").addClass("CKB_HIGHLIGHT");
			}
			else {
				$(this).parents("td.tdCkb").removeClass("CKB_HIGHLIGHT");
			}
		})
	});
</script>

<script language="JavaScript" type="text/javascript">
var windowScrollTopAnterior;
window.status='Aguarde, executando a consulta ...';

//Every resize of window
$(window).resize(function() {
	sizeDivPedidoConsulta();
});

function sizeDivPedidoConsulta() {
	var newHeight = $(document).height() + "px";
	$("#divPedidoConsulta").css("height", newHeight);
}

function fechaDivPedidoConsulta() {
	$(window).scrollTop(windowScrollTopAnterior);
	$("#divPedidoConsulta").fadeOut();
	$("#iframePedidoConsulta").attr("src", "");
}

function realca_cor_mouse_over(c) {
	c.style.backgroundColor = 'palegreen';
}

function realca_cor_mouse_out(c) {
	c.style.backgroundColor = '';
}

function fPEDConsulta(id_pedido, usuario) {
	windowScrollTopAnterior = $(window).scrollTop();
	sizeDivPedidoConsulta();
	$("#iframePedidoConsulta").attr("src", "PedidoConsultaView.asp?pedido_selecionado=" + id_pedido + "&pedido_selecionado_inicial=" + id_pedido + "&usuario=" + usuario);
	$("#divPedidoConsulta").fadeIn();
}

function fRELConcluir(s_id){
	window.status = "Aguarde ...";
	fREL.pedido_selecionado.value=s_id;
	fREL.submit(); 
}

function fRELGravaDados(f) {
	var i, intQtdeTratados;
	var s, p, d, dtR;

	intQtdeTratados = 0;

	for (i = 1; i < f.ckb_recebido.length; i++) {
	    if (f.ckb_recebido[i].checked) {
	        intQtdeTratados++;
	    }
	}

	if (intQtdeTratados == 0) {
	    alert('Nenhum pedido foi selecionado!!');
	    return;
	}

	window.status = "Aguarde ...";
	f.action = "RelPedidosMktplaceRecebidosGravaDados.asp";
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

<link href="<%=URL_FILE__E_CSS%>" rel="stylesheet" type="text/css">
<link href="<%=URL_FILE__EPRINTER_CSS%>" rel="stylesheet" type="text/css" media="print">
<link href="<%=URL_FILE__JQUERY_UI_CSS%>" rel="stylesheet" type="text/css">

<style type="text/css">
.tdDataEntrega{
	width: 70px;
	}
.tdPedido{
	width: 70px;
	}
.tdCliente{
	width: 240px;
	}
.tdMarketplace{
	width: 150px;
	}
.tdRecebido{
	width: 40px;
	}
.cDtColeta
{
	background-color:transparent;
}
.CKB_HIGHLIGHT
{
	background-color:#90EE90;
}
#divPedidoConsulta
{
	position:absolute;
	top:0;
	left:0;
	width:100%;
	z-index:1000;
	background-color:#808080;
	opacity: 1;
}
#divInternoPedidoConsulta
{
	position:absolute;
	top:6%;
	left:5%;
	width:90%;
	height:90%;
	z-index:1000;
	background-color:#808080;
	opacity: 1;
}
#divInternoPedidoConsulta.divFixo
{
	position:fixed;
	top:6%;
}
#imgFechaDivPedidoConsulta
{
	position:fixed;
	top:6%;
	left: 50%;
	margin-left: -16px; /* -1 * image width / 2 */
	margin-top: -32px;
	z-index:1001;
}
#iframePedidoConsulta
{
	position:absolute;
	top:0;
	left:0;
	width:100%;
	height:100%;
	border: solid 4px black;
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
<table cellspacing="0">
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

<form id="fREL" name="fREL" method="post" action="Pedido.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="pedido_selecionado" id="pedido_selecionado" value=''>
<input type="hidden" name="c_transportadora" id="c_transportadora" value="<%=c_transportadora%>">
<input type="hidden" name="c_loja" id="c_loja" value="<%=c_loja%>">
<input type="hidden" name="c_dt_hoje" id="c_dt_hoje" value="<%=formata_data(Date)%>" />
<input type="hidden" name="c_usuario_sessao" id="c_usuario_sessao" value="<%=usuario%>" />

<!--  ASSEGURA CRIAÇÃO DE UM ARRAY DE CAMPOS, MESMO QUANDO HOUVER SOMENTE 1 LINHA!! -->
<input type="hidden" name="ckb_recebido" id="ckb_recebido" value="">
<input type="hidden" name="c_dt_recebimento" id="c_dt_recebimento" value="">


<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="709" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Registro de Pedidos Marketplace Recebidos Pelo Cliente</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>

<!-- FILTROS -->
<% 
	s_filtro = "<table width='709' cellpadding='0' cellspacing='0' style='border-bottom:1px solid black' border='0'>"

	s = c_transportadora
	if s = "" then 
		s = "todas"
	else
		if (s_nome_transportadora <> "") And (s_nome_transportadora <> c_transportadora) then s = s & "  (" & s_nome_transportadora & ")"
		end if
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>Transportadora:&nbsp;</span></td><td align='left' valign='top'>" & _
			   "<span class='N'>" & s & "</span></td></tr>"

    s = c_grupo_pedido_origem
	if s = "" then 
		s = "todos"
	else
		s = s_nome_grupo_pedido_origem
		end if
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>Origem do Pedido (Grupo):&nbsp;</span></td><td align='left' valign='top'>" & _
			   "<span class='N'>" & s & "</span></td></tr>"

    s = c_pedido_origem
	if s = "" then 
		s = "todos"
	else
		s = s_nome_pedido_origem
		end if
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>Origem do Pedido:&nbsp;</span></td><td align='left' valign='top'>" & _
			   "<span class='N'>" & s & "</span></td></tr>"

	s = c_loja
	if s = "" then 
		s = "todas"
	else
		s = s & "  (" & s_nome_loja & ")"
		end if
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>Loja:&nbsp;</span></td><td align='left' valign='top'>" & _
			   "<span class='N'>" & s & "</span></td></tr>"
	
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>Emissão:&nbsp;</span></td><td align='left' valign='top' width='99%'>" & _
			   "<span class='N'>" & formata_data_hora(Now) & "</span></td></tr>"

	s_filtro = s_filtro & "</table>"
	Response.Write s_filtro
%>

<!--  RELATÓRIO  -->
<br>
<% consulta_executa %>

<!-- ************   SEPARADOR   ************ -->
<table width="709" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>
<br>

<table class="notPrint" width="709" cellpadding="0" cellspacing="0" style="margin-top:5px;">
<tr>
	<td align="left">		
		<button type="button" name="bMarcarTodos" id="bMarcarTodos" class="Button BTN_LNK" style="margin-left:6px;margin-bottom:2px;width:150px">Marcar todos</button>
		&nbsp;
		<button type="button" name="bDesmarcarTodos" id="bDesmarcarTodos" class="Button BTN_LNK" style="margin-left:6px;margin-right:6px;margin-bottom:2px;width:150px">Desmarcar todos</button>
	</td>
</tr>
</table>

<!-- ************   SEPARADOR   ************ -->
<table width="709" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>
<br>

<table class="notPrint" width="709" cellspacing="0" border="0">
<tr>
	<% if qtde_total_pedidos > 0 then %>
	<td align="left">
		<a name="bVOLTAR" id="bVOLTAR" href="<%=strUrlBotaoVoltar%>" title="volta para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a>
	</td>
	<td>&nbsp;</td>
	<td align="right">
		<a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fRELGravaDados(fREL)" title="grava os dados"><img src="../botao/confirmar.gif" width="176" height="55" border="0"></a>
	</td>
	<% else %>
	<td align="center"><a name="bVOLTAR" id="bVOLTAR" href="<%=strUrlBotaoVoltar%>" title="volta para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a>
	<% end if %>
	</td>
</tr>
</table>

</form>

</center>

<div id="divPedidoConsulta"><center><div id="divInternoPedidoConsulta"><img id="imgFechaDivPedidoConsulta" src="../imagem/close_button_32.png" title="clique para fechar o painel de consulta" /><iframe id="iframePedidoConsulta"></iframe></div></center></div>

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
