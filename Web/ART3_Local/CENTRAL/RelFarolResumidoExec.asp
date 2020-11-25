<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/estoque.asp" -->
<!-- #include file = "../global/Global.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =================================================================
'	  RelFarolResumidoExec.asp
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

	dim usuario
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	if Not operacao_permitida(OP_CEN_REL_FAROL_RESUMIDO, s_lista_operacoes_permitidas) then
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	dim s, s_aux, strScript, msg_erro
	dim c_dt_periodo_inicio, c_dt_periodo_termino, c_perc_est_cresc, perc_est_cresc, c_fabricante, c_grupo, c_subgrupo, c_potencia_BTU, c_ciclo, c_posicao_mercado
	
	
'	OBTÉM DADOS DO FORMULÁRIO
	c_dt_periodo_inicio = Trim(Request.Form("c_dt_periodo_inicio"))
	c_dt_periodo_termino = Trim(Request.Form("c_dt_periodo_termino"))
	c_fabricante = Trim(Request.Form("c_fabricante"))
	c_grupo = Ucase(Trim(Request.Form("c_grupo")))
	c_subgrupo = Ucase(Trim(Request.Form("c_subgrupo")))
	c_potencia_BTU = Trim(Request.Form("c_potencia_BTU"))
	c_ciclo = Trim(Request.Form("c_ciclo"))
	c_posicao_mercado = Trim(Request.Form("c_posicao_mercado"))
	c_perc_est_cresc = Trim(Request.Form("c_perc_est_cresc"))
	perc_est_cresc = converte_numero(c_perc_est_cresc)
	
'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	dim alerta
	alerta=""

	if alerta = "" then
		if c_dt_periodo_inicio = "" then
			alerta = texto_add_br(alerta)
			alerta = alerta & "É necessário informar a data inicial do período de vendas!!"
		elseif Not IsDate(c_dt_periodo_inicio) then
			alerta = texto_add_br(alerta)
			alerta = alerta & "Data inicial inválida do período de vendas!!"
			end if

		if c_dt_periodo_termino = "" then
			alerta = texto_add_br(alerta)
			alerta = alerta & "É necessário informar a data final do período de vendas!!"
		elseif Not IsDate(c_dt_periodo_termino) then
			alerta = texto_add_br(alerta)
			alerta = alerta & "Data final inválida do período de vendas!!"
			end if
		end if
	
	dim s_nome_fabricante
	dim s_where_temp, v_fabricantes, cont, v_grupos, v_subgrupos
	s_nome_fabricante = ""
	s_where_temp = ""
	v_fabricantes = ""
	if alerta = "" then
		if c_fabricante <> "" then
		    s = "SELECT nome, razao_social FROM t_FABRICANTE WHERE "
		    v_fabricantes = split(c_fabricante, ", ")
		    for cont = LBound(v_fabricantes) to UBound(v_fabricantes)
                if s_where_temp <> "" then s_where_temp = s_where_temp & " OR"
		        s_where_temp = s_where_temp & " (fabricante = '" & v_fabricantes(cont) & "')"
            next
            s = s & s_where_temp
			set rs = cn.Execute(s)
			if rs.Eof then
				alerta = "FABRICANTE '" & c_fabricante & "' NÃO ESTÁ CADASTRADO"
			else
				s_nome_fabricante = Trim("" & rs("razao_social"))
				if s_nome_fabricante = "" then s_nome_fabricante = Trim("" & rs("nome"))
				end if
			end if
		end if
	
	if alerta = "" then
		call set_default_valor_texto_bd(usuario, "RelFarolResumidoFiltro|c_dt_periodo_inicio", c_dt_periodo_inicio)
		call set_default_valor_texto_bd(usuario, "RelFarolResumidoFiltro|c_dt_periodo_termino", c_dt_periodo_termino)
		call set_default_valor_texto_bd(usuario, "RelFarolResumidoFiltro|c_perc_est_cresc", c_perc_est_cresc)
		call set_default_valor_texto_bd(usuario, "RelFarolResumidoFiltro|c_fabricante", c_fabricante)
		call set_default_valor_texto_bd(usuario, "RelFarolResumidoFiltro|c_grupo", c_grupo)
		call set_default_valor_texto_bd(usuario, "RelFarolResumidoFiltro|c_subgrupo", c_subgrupo)
		call set_default_valor_texto_bd(usuario, "RelFarolResumidoFiltro|c_potencia_BTU", c_potencia_BTU)
		call set_default_valor_texto_bd(usuario, "RelFarolResumidoFiltro|c_ciclo", c_ciclo)
		call set_default_valor_texto_bd(usuario, "RelFarolResumidoFiltro|c_posicao_mercado", c_posicao_mercado)
		end if
	
	dim s_log, s_log_aux
	if alerta = "" then
		s_log_aux = " Período de vendas = "
		s_aux = c_dt_periodo_inicio
		if s_aux = "" then s_aux = "N.I."
		s_log_aux = s_log_aux & s_aux & " a "
		s_aux = c_dt_periodo_termino
		if s_aux = "" then s_aux = "N.I."
		s_log_aux = s_log_aux & s_aux
		s_aux = c_fabricante
		if s_aux = "" then s_aux = "N.I."
		s_log_aux = s_log_aux & "; Fabricante = " & s_aux
		s_aux = c_grupo
		if s_aux = "" then s_aux = "N.I."
		s_log_aux = s_log_aux & "; Grupo de produtos = " & s_aux
		s_aux = c_subgrupo
		if s_aux = "" then s_aux = "N.I."
		s_log_aux = s_log_aux & "; Subgrupo de produtos = " & s_aux
		s_aux = c_potencia_BTU
		if s_aux = "" then s_aux = "N.I."
		s_log_aux = s_log_aux & "; BTU/h = " & s_aux
		s_aux = c_ciclo
		if s_aux = "" then s_aux = "N.I."
		s_log_aux = s_log_aux & "; Ciclo = " & s_aux
		s_aux = c_posicao_mercado
		if s_aux = "" then s_aux = "N.I."
		s_log_aux = s_log_aux & "; Posição Mercado = " & s_aux
		s_aux = c_perc_est_cresc
		if s_aux = "" then s_aux = "N.I." else s_aux = s_aux & "%"
		s_log_aux = s_log_aux & "; Percentual estimado de crescimento = " & s_aux
		s_log = "Sucesso na exibição da página p/ geração da planilha do Farol Resumido:" & s_log_aux
		grava_log usuario, "", "", "", OP_LOG_FAROL_RESUMIDO, s_log
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


<%=DOCTYPE_LEGADO%>

<html>


<head>
	<title>CENTRAL</title>
	</head>


<% if False then 'APENAS P/ HABILITAR O INTELLISENSE DURANTE O DESENVOLVIMENTO!! %>
<script src="../Global/jquery.js" language="JavaScript" type="text/javascript"></script>
<% end if %>

<script src="<%=URL_FILE__JQUERY%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__CONSTXL_JS%>" language="JavaScript" type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
	$(document).ready(function() {
		$("#divBotaoGeraPlanilha").click(function() {
			if ($("#c_qtde_total_registros").val() == "0") {
				alert("A consulta não retornou nenhum resultado!!");
				return;
			}
			$(this).hide();
			$("#textoMensagem").text("Aguarde, a planilha está sendo gerada...");
			$("#divMsgAguarde").removeClass("modalOk");
			$("#divMsgAguarde").removeClass("modalErro");
			$("#divMsgAguarde").addClass("modalAguarde");
			$("#divMsgAguarde").show("fast");
			$("#btnFechar").hide();
			agendaExecucaoGeraPlanilha();
		});
		$("#btnFechar").click(function() {
			$(this).hide();
			$("#divBotaoGeraPlanilha").show("fast");
			$("#divMsgAguarde").hide();
		});
	});
</script>

<script language="JavaScript" type="text/javascript">
var nExcelMajorVersion = 0;
	
function agendaExecucaoGeraPlanilha() {
	setTimeout('GeraPlanilha()', 1000);
}

function CarregaCelulaComHtml(range, sDescricaoHtml) {
var c;
	if (sDescricaoHtml.toString().indexOf("<") == -1) return;
	c = document.getElementById("c_clipboard");
	c.innerHTML = sDescricaoHtml;
	c.createTextRange().execCommand("Copy");
	range.PasteSpecial();
}
</script>


<%
'	GERA O SCRIPT QUE IRÁ ATIVAR O EXCEL E GERAR A PLANILHA NO LADO DO CLIENTE
'	~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~

'	PREPARA CONSULTA AO BANCO DE DADOS
'	~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	dim r, r_aux
	dim n, n_reg, n_reg_total
	dim s_sql, s_sql_lista_base, s_sql_qtde_vendida, s_sql_qtde_devolvida, s_sql_qtde_estoque_venda
	
'	CRIA O RECORDSET AUXILIAR
	if Not cria_recordset_otimista(r_aux, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	
'	MONTA O SQL QUE SELECIONA A RELAÇÃO DE PRODUTOS
'	A LÓGICA CONSISTE EM SELECIONAR:
'		1) PRODUTOS QUE TENHAM SALDO NO ESTOQUE DE VENDA
'		2) PRODUTOS QUE CONSTEM COMO 'VENDÁVEIS'
'	OBS: O USO DE 'UNION' SIMPLES ELIMINA AS LINHAS DUPLICADAS DOS RESULTADOS
'		 O USO DE 'UNION ALL' RETORNARIA TODAS AS LINHAS, INCLUSIVE AS DUPLICADAS
	s_sql_lista_base = _
		"SELECT DISTINCT" & _
			" fabricante," & _
			" produto" & _
		" FROM t_ESTOQUE_ITEM" & _
		" WHERE" & _
			" ((qtde - qtde_utilizada) > 0) "
	
	s_where_temp = ""
	if c_fabricante <> "" then
	for cont = Lbound(v_fabricantes) to Ubound(v_fabricantes)
	    if s_where_temp <> "" then s_where_temp = s_where_temp & " OR"
		s_where_temp = s_where_temp & _
			   " (fabricante = '" & v_fabricantes(cont) & "')"
	next
	s_sql_lista_base = s_sql_lista_base & "AND"
	s_sql_lista_base = s_sql_lista_base & "(" & s_where_temp & ")"
    end if
    
	
	s_sql_lista_base = s_sql_lista_base & _
		" UNION " & _
		"SELECT DISTINCT" & _
			" t_PRODUTO.fabricante," & _
			" t_PRODUTO.produto" & _
		" FROM t_PRODUTO" & _
			" INNER JOIN" & _
				"(" & _
					"SELECT DISTINCT" & _
						" fabricante," & _
						" produto" & _
					" FROM t_PRODUTO_LOJA" & _
					" WHERE" & _
						" (vendavel = 'S') "
	s_where_temp = ""
	if c_fabricante <> "" then
	for cont = Lbound(v_fabricantes) to Ubound(v_fabricantes)
	if s_where_temp <> "" then s_where_temp = s_where_temp & " OR"
		s_where_temp = s_where_temp & _
						" (fabricante = '" & v_fabricantes(cont) & "')"
	next
	s_sql_lista_base = s_sql_lista_base & "AND"
	s_sql_lista_base = s_sql_lista_base & "(" & s_where_temp & ")"
		end if
	
	s_sql_lista_base = s_sql_lista_base & _
				") tPL_AUX ON (t_PRODUTO.fabricante=tPL_AUX.fabricante) AND (t_PRODUTO.produto=tPL_AUX.produto)" & _
		" WHERE" & _
			" (excluido_status = 0) "
	
	s_where_temp = ""
	if c_fabricante <> "" then
	for cont = Lbound(v_fabricantes) to Ubound(v_fabricantes)
	    if s_where_temp <> "" then s_where_temp = s_where_temp & " OR"
		    s_where_temp = s_where_temp & _
			    " (t_PRODUTO.fabricante = '" & v_fabricantes(cont) & "')"
	next
	s_sql_lista_base = s_sql_lista_base & "AND"
	s_sql_lista_base = s_sql_lista_base & "(" & s_where_temp & ")"
	end if
	
	s_sql_lista_base = s_sql_lista_base & _
		" UNION " & _
		" SELECT DISTINCT" & _
			" fabricante," & _
			" produto" & _
		" FROM t_PRODUTO" & _
		" WHERE" & _
			" (farol_qtde_comprada > 0) "
	s_where_temp = ""
	if c_fabricante <> "" then
	for cont = Lbound(v_fabricantes) to Ubound(v_fabricantes)
	if s_where_temp <> "" then s_where_temp = s_where_temp & " OR"
		s_where_temp = s_where_temp & _
			" (fabricante = '" & v_fabricantes(cont) & "')"
			next
			s_sql_lista_base = s_sql_lista_base & "AND"
			s_sql_lista_base = s_sql_lista_base & "(" & s_where_temp & ")"
		end if
	
'	SELECT DOS PRODUTOS VENDIDOS NO PERÍODO
	s_sql_qtde_vendida = _
		"SELECT" & _
			" SUM(qtde)" & _
		" FROM t_PEDIDO" & _
			" INNER JOIN t_PEDIDO_ITEM ON (t_PEDIDO.pedido=t_PEDIDO_ITEM.pedido)" & _
		" WHERE" & _
			" (t_PEDIDO_ITEM.fabricante=t_PROD_LISTA_BASE.fabricante)" & _
			" AND (t_PEDIDO_ITEM.produto=t_PROD_LISTA_BASE.produto)" & _
			" AND (st_entrega <> '" & ST_ENTREGA_CANCELADO & "')"
	
	if c_dt_periodo_inicio <> "" then
		s_sql_qtde_vendida = s_sql_qtde_vendida & _
			" AND (t_PEDIDO.data >= " & bd_formata_data(StrToDate(c_dt_periodo_inicio)) & ")"
		end if
	
	if c_dt_periodo_termino <> "" then
		s_sql_qtde_vendida = s_sql_qtde_vendida & _
			" AND (t_PEDIDO.data < " & bd_formata_data(StrToDate(c_dt_periodo_termino)+1) & ")"
		end if
	
'	SELECT DOS PRODUTOS DEVOLVIDOS NO PERÍODO
	s_sql_qtde_devolvida = _
		"SELECT" &_ 
			" SUM(qtde)" & _
		" FROM t_PEDIDO_ITEM_DEVOLVIDO" & _
		" WHERE" & _
			" (t_PEDIDO_ITEM_DEVOLVIDO.fabricante = t_PROD_LISTA_BASE.fabricante)" & _
			" AND (t_PEDIDO_ITEM_DEVOLVIDO.produto = t_PROD_LISTA_BASE.produto)"
	
	if c_dt_periodo_inicio <> "" then
		s_sql_qtde_devolvida = s_sql_qtde_devolvida & _
			" AND (t_PEDIDO_ITEM_DEVOLVIDO.devolucao_data >= " & bd_formata_data(StrToDate(c_dt_periodo_inicio)) & ")"
		end if
	
	if c_dt_periodo_termino <> "" then
		s_sql_qtde_devolvida = s_sql_qtde_devolvida & _
			" AND (t_PEDIDO_ITEM_DEVOLVIDO.devolucao_data < " & bd_formata_data(StrToDate(c_dt_periodo_termino)+1) & ")"
		end if
	
'	SELECT DA QUANTIDADE DISPONÍVEL NO ESTOQUE DE VENDA
	s_sql_qtde_estoque_venda = _
		"SELECT" & _
			" SUM(qtde-qtde_utilizada)" & _
		" FROM t_ESTOQUE_ITEM" &_
		" WHERE" & _
			" (t_ESTOQUE_ITEM.fabricante = t_PROD_LISTA_BASE.fabricante)" & _
			" AND (t_ESTOQUE_ITEM.produto = t_PROD_LISTA_BASE.produto)" & _
			" AND ((qtde - qtde_utilizada) > 0)"
	
'	SELECT COMPLETO
	s_sql = _
		"SELECT" & _
			" fabricante," & _
			" produto," & _
			" descricao," & _
			" descricao_html," & _
			" grupo," & _
			" subgrupo," & _
			" potencia_BTU," & _
			" ciclo," & _
			" posicao_mercado," & _
			" descontinuado," & _
			" Coalesce(farol_qtde_comprada, 0) AS farol_qtde_comprada," & _
			" Coalesce(qtde_vendida, 0) AS qtde_vendida," & _
			" Coalesce(qtde_devolvida, 0) AS qtde_devolvida," & _
			" Coalesce(qtde_estoque_venda, 0) AS qtde_estoque_venda" & _
		" FROM (" & _
			"SELECT" & _
				" t_PROD_LISTA_BASE.fabricante," & _
				" t_PROD_LISTA_BASE.produto," & _
				" t_PRODUTO.descricao," & _
				" t_PRODUTO.descricao_html," & _
				" Coalesce(t_PRODUTO.grupo, '') AS grupo," & _
				" Coalesce(t_PRODUTO.subgrupo, '') AS subgrupo," & _
				" Coalesce(t_PRODUTO.potencia_BTU, '') AS potencia_BTU," & _
				" Coalesce(t_PRODUTO.ciclo, '') AS ciclo," & _
				" Coalesce(t_PRODUTO.posicao_mercado, '') AS posicao_mercado," & _
				" Coalesce(t_PRODUTO.descontinuado, '') AS descontinuado," & _
				" t_PRODUTO.farol_qtde_comprada," & _
				" (" & s_sql_qtde_vendida & ") AS qtde_vendida," & _
				" (" & s_sql_qtde_devolvida & ") AS qtde_devolvida," & _
				" (" & s_sql_qtde_estoque_venda & ") AS qtde_estoque_venda" & _
			" FROM (" & s_sql_lista_base & ") t_PROD_LISTA_BASE" & _
				" LEFT JOIN t_PRODUTO ON (t_PROD_LISTA_BASE.fabricante = t_PRODUTO.fabricante) AND (t_PROD_LISTA_BASE.produto = t_PRODUTO.produto)" & _
			" WHERE" & _
				" (descricao <> '.')" & _
				" AND (descricao <> '*')" & _
			") tREL" & _
		" WHERE" & _
			" (UPPER(descontinuado) <> 'S') "

	s_where_temp = ""
	if c_grupo <> "" then
		v_grupos = split(c_grupo, ", ")
		for cont = Lbound(v_grupos) to Ubound(v_grupos)
			if s_where_temp <> "" then s_where_temp = s_where_temp & " OR"
			s_where_temp = s_where_temp & _
				" (grupo = '" & v_grupos(cont) & "')"
		next
		s_sql = s_sql & "AND "
		s_sql = s_sql & "(" & s_where_temp & ")"
		end if
	
	s_where_temp = ""
	if c_subgrupo <> "" then
		v_subgrupos = split(c_subgrupo, ", ")
		for cont = Lbound(v_subgrupos) to Ubound(v_subgrupos)
			if s_where_temp <> "" then s_where_temp = s_where_temp & " OR"
			s_where_temp = s_where_temp & _
				" (subgrupo = '" & v_subgrupos(cont) & "')"
		next
		s_sql = s_sql & "AND "
		s_sql = s_sql & "(" & s_where_temp & ")"
		end if

	if c_potencia_BTU <> "" then
		s_sql = s_sql & _
			" AND (potencia_BTU = " & c_potencia_BTU & ")"
		end if
	
	if c_ciclo <> "" then
		s_sql = s_sql & _
			" AND (ciclo = '" & c_ciclo & "')"
		end if
	
	if c_posicao_mercado <> "" then
		s_sql = s_sql & _
			" AND (posicao_mercado = '" & c_posicao_mercado & "')"
		end if
	
'	SE AS QUANTIDADES ESPECIFICADAS FOREM TODAS IGUAIS A ZERO, NÃO EXIBE NO RELATÓRIO
'	IMPORTANTE: LEMBRANDO QUE O RESULTADO DO 'SELECT' EM QUE ALGUMAS DAS QUANTIDADES É CALCULADA PODE RETORNAR 'NULL' E QUE A FUNÇÃO
'				COALESCE() NÃO PODE RECEBER UMA CONSULTA SQL COMO PARÂMETRO.
	s_sql = "SELECT " & _
				"*" & _
			" FROM (" & s_sql & ") tRelFinal" & _
			" WHERE" & _
				" (" & _
					"NOT (" & _
						"((qtde_vendida - qtde_devolvida) <= 0)" & _
						" AND (qtde_estoque_venda = 0)" & _
						" AND (farol_qtde_comprada = 0)" & _
					")" & _
				")"
	
	s_sql = s_sql & _
		" ORDER BY" & _
			" fabricante," & _
			" produto"
	
	
'	MONTA SCRIPT P/ EXECUTAR NO BROWSER
'	~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'	DECLARAÇÕES
	strScript = _
	"<script language='JavaScript' type='text/javascript'>" & chr(13) & _
	"var xlFN_LISTAGEM = 'Arial';" & chr(13) & _
	"var xlFS_LISTAGEM = 10;" & chr(13) & _
	"var xlFS_CABECALHO = 12;" & chr(13) & _
	"var xlMargemEsq=1;" & chr(13) & _
	"var xlOffSetArray=2;" & chr(13) & _
	"var xlNumLinha, xlNumLinhaAux, xlDadosMinIndex, xlDadosMaxIndex;" & chr(13) & _
	"var xlFabricante, xlProduto, xlDescricao, xlGrupo, xlSubgrupo, xlPotenciaBTU, xlCiclo, xlPosicaoMercado, xlQtdeVendida, xlQtdeEstimadaAVender, xlQtdeEstoqueVenda, xlQtdeComprada, xlQtdeSaldo;" & chr(13) & _
	"var xlCellPercEstCresc;" & chr(13) & _
	"var oXL, oWB, oWS, oRange, oBorders, oFont, oStyle;" & chr(13) & _
	"var i_dados_inicio = 0;" & chr(13) & _
	"var i_dados_fim = 0;" & chr(13) & _
	"function montaLinha(fabricante, produto, descricao, descricao_html, grupo, subgrupo, potencia_BTU, ciclo, posicao_mercado, qtde_vendida, qtde_estoque_venda, qtde_comprada) {" & chr(13) &_
	"var i, s;" & chr(13) & _
	"	xlNumLinha++;" & chr(13) & _
	"	oRange=oWS.Range(excel_converte_numeracao_digito_para_letra(xlFabricante) + xlNumLinha.toString());" & chr(13) & _
	"	oRange.Value=fabricante;" & chr(13) & _
	"	oRange=oWS.Range(excel_converte_numeracao_digito_para_letra(xlProduto) + xlNumLinha.toString());" & chr(13) & _
	"	oRange.Value=produto;" & chr(13) & _
	"	oRange=oWS.Range(excel_converte_numeracao_digito_para_letra(xlDescricao) + xlNumLinha.toString());" & chr(13) & _
	"	oRange.Value=descricao;" & chr(13) & _
	"	oRange=oWS.Range(excel_converte_numeracao_digito_para_letra(xlGrupo) + xlNumLinha.toString());" & chr(13) & _
	"	oRange.Value=grupo;" & chr(13) & _
	"	oRange=oWS.Range(excel_converte_numeracao_digito_para_letra(xlSubgrupo) + xlNumLinha.toString());" & chr(13) & _
	"	oRange.Value=subgrupo;" & chr(13) & _
	"	if (potencia_BTU != 0) {" & chr(13) & _
	"		oRange=oWS.Range(excel_converte_numeracao_digito_para_letra(xlPotenciaBTU) + xlNumLinha.toString());" & chr(13) & _
	"		oRange.Value=potencia_BTU;" & chr(13) & _
	"	}" & chr(13) & _
	"	oRange=oWS.Range(excel_converte_numeracao_digito_para_letra(xlCiclo) + xlNumLinha.toString());" & chr(13) & _
	"	oRange.Value=ciclo;" & chr(13) & _
	"	oRange=oWS.Range(excel_converte_numeracao_digito_para_letra(xlPosicaoMercado) + xlNumLinha.toString());" & chr(13) & _
	"	oRange.Value=posicao_mercado;" & chr(13) & _
	"	oRange=oWS.Range(excel_converte_numeracao_digito_para_letra(xlQtdeVendida) + xlNumLinha.toString());" & chr(13) & _
	"	oRange.Value=qtde_vendida;" & chr(13) & _
	"	oRange=oWS.Range(excel_converte_numeracao_digito_para_letra(xlQtdeEstoqueVenda) + xlNumLinha.toString());" & chr(13) & _
	"	oRange.Value=qtde_estoque_venda;" & chr(13) & _
	"	oRange=oWS.Range(excel_converte_numeracao_digito_para_letra(xlQtdeComprada) + xlNumLinha.toString());" & chr(13) & _
	"	oRange.Value=qtde_comprada;" & chr(13) & _
	"	// oRange=oWS.Range(excel_converte_numeracao_digito_para_letra(xlDescricao) + xlNumLinha.toString());" & chr(13) & _
	"	// if (nExcelMajorVersion >= 11) CarregaCelulaComHtml(oRange, descricao_html);" & chr(13) & _
	"	oRange=oWS.Range(excel_converte_numeracao_digito_para_letra(xlQtdeEstimadaAVender) + xlNumLinha.toString());" & chr(13) & _
	"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlQtdeVendida) + xlNumLinha.toString() + '+(' + xlCellPercEstCresc + '/100)*' + excel_converte_numeracao_digito_para_letra(xlQtdeVendida) + xlNumLinha.toString();" & chr(13) & _
	"	oRange.FormulaLocal='=' + xlNumLinhaAux;" & chr(13) & _
	"	oRange=oWS.Range(excel_converte_numeracao_digito_para_letra(xlQtdeSaldo) + xlNumLinha.toString());" & chr(13) & _
	"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlQtdeComprada) + xlNumLinha.toString() + '+' + excel_converte_numeracao_digito_para_letra(xlQtdeEstoqueVenda) + xlNumLinha.toString() + '-' + excel_converte_numeracao_digito_para_letra(xlQtdeEstimadaAVender) + xlNumLinha.toString();" & chr(13) & _
	"	oRange.FormulaLocal='=' + xlNumLinhaAux;" & chr(13) & _
	"	// LINHA DE SEPARAÇÃO" & chr(13) & _
	"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlDadosMinIndex) + xlNumLinha.toString() + ':' + excel_converte_numeracao_digito_para_letra(xlDadosMaxIndex) + xlNumLinha.toString();" & chr(13) & _
	"	oRange=oWS.Range(xlNumLinhaAux);" & chr(13) & _
	"	oBorders=oRange.Borders(xlEdgeBottom);" & chr(13) & _
	"	oBorders.LineStyle=xlDot;" & chr(13) & _
	"	oBorders.Weight=xlHairline;" & chr(13) & _
	"	oBorders.ColorIndex=xlAutomatic;" & chr(13) & _
	"}" & chr(13) & _
	"" & chr(13) & _
	"function GeraPlanilha() {" & chr(13) & _
	"	var i, n, s;" & chr(13)
	
'	TENTA CRIAR UMA INSTÂNCIA DO EXCEL
	strScript = strScript & _
	"	try {" & chr(13) & _
	"		oXL = new ActiveXObject('Excel.Application');" & chr(13) & _ 
	"		}" & chr(13) & _
	"	catch (e) {" & chr(13) & _
	"		alert('Falha ao acionar o Excel!!\n   1) Verifique se o Excel está corretamente instalado\n   2) Verifique se está ativada a opção: Inicializar e executar scripts de controles ActiveX não marcados como seguros');" & chr(13) & _
	"		history.back();" & chr(13) & _
	"		return;" & chr(13) & _
	"		}" & chr(13) 

'	CONFIGURA O EXCEL
	strScript = strScript & _
	"" & chr(13) & _
	"try {" & chr(13) & _
	"	nExcelMajorVersion = parseInt(oXL.Version.split('.',1));" & chr(13) & _
	"	oXL.Visible=true;" & chr(13) & _
	"	oXL.DisplayAlerts=false;" & chr(13) & _
	"	oXL.SheetsInNewWorkbook=1;" & chr(13) & _
	"	oWB=oXL.Workbooks.Add;" & chr(13) & _
	"	oWB.Windows(1).WindowState=xlMaximized;" & chr(13) & _
	"	oWS=oWB.Worksheets(1);" & chr(13) & _
	"	oWS.PageSetup.PaperSize = xlPaperA4;" & chr(13) & _
	"	oWS.PageSetup.Orientation = xlPortrait;" & chr(13) & _
	"	oWS.PageSetup.LeftMargin = 2;" & chr(13) & _
	"	oWS.PageSetup.RightMargin = 2;" & chr(13) & _
	"	oWS.PageSetup.TopMargin = 15;" & chr(13) & _
	"	oWS.PageSetup.BottomMargin = 15;" & chr(13) & _
	"	oWS.PageSetup.HeaderMargin = 5;" & chr(13) & _
	"	oWS.PageSetup.FooterMargin = 5;" & chr(13) & _
	"	oWS.PageSetup.CenterHorizontally = true;" & chr(13) & _
	"	oXL.Windows(1).DisplayGridlines=false;" & chr(13) & _
	"	oXL.Windows(1).DisplayHeadings=true;" & chr(13) & _
	"	oStyles=oWB.Styles('Normal');" & chr(13) & _
	"	oStyles.IncludeNumber=true;" & chr(13) & _
	"	oStyles.IncludeFont=true;" & chr(13) & _
	"	oStyles.IncludeAlignment=true;" & chr(13) & _
	"	oStyles.IncludeBorder=true;" & chr(13) & _
	"	oStyles.IncludePatterns=true;" & chr(13) & _
	"	oStyles.IncludeProtection=true;" & chr(13) & _
	"	oFont=oStyles.Font" & chr(13) & _
	"	oFont.Name=xlFN_LISTAGEM;" & chr(13) & _
	"	oFont.Size=xlFS_LISTAGEM;" & chr(13) & _
	"	oFont.Bold=false;" & chr(13) & _
	"	oFont.Italic=false;" & chr(13) & _
	"	oFont.Underline=xlUnderlineStyleNone;" & chr(13) & _
	"	oFont.Strikethrough=false;" & chr(13) & _
	"	oFont.ColorIndex=xlAutomatic;" & chr(13) & _
	"	oStyles.HorizontalAlignment=xlLeft;" & chr(13) & _
	"	oStyles.VerticalAlignment=xlTop;" & chr(13) & _
	"	oStyles.WrapText=false;" & chr(13) & _
	"	oStyles.Orientation=0;" & chr(13) & _
	"	oStyles.IndentLevel=0;" & chr(13) & _
	"	oStyles.ShrinkToFit=false;" & chr(13) & _
	"	oStyles.Borders(xlLeft).LineStyle=xlNone;" & chr(13) & _
	"	oStyles.Borders(xlRight).LineStyle=xlNone;" & chr(13) & _
	"	oStyles.Borders(xlTop).LineStyle=xlNone;" & chr(13) & _
	"	oStyles.Borders(xlBottom).LineStyle=xlNone;" & chr(13) & _
	"	oStyles.Borders(xlDiagonalDown).LineStyle=xlNone;" & chr(13) & _
	"	oStyles.Borders(xlDiagonalUp).LineStyle=xlNone;" & chr(13) & _
	"	oWS.Cells.Style='Normal';" & chr(13) & _
	"	oWS.DisplayPageBreaks=false;" & chr(13) & _
	"	oWS.Name='Farol-Resumido';" & chr(13) & _
	"	oXL.DisplayAlerts=true;" & chr(13)

'	POSIÇÃO DAS COLUNAS
	strScript = strScript & _
	"	xlFabricante=xlMargemEsq+1;" & chr(13) & _
	"	xlProduto=xlFabricante+1;" & chr(13) & _
	"	xlDescricao=xlProduto+1;" & chr(13) & _
	"	xlGrupo=xlDescricao+1;" & chr(13) & _
	"	xlSubgrupo=xlGrupo+1;" & chr(13) & _
	"	xlPotenciaBTU=xlSubgrupo+1;" & chr(13) & _
	"	xlCiclo=xlPotenciaBTU+1;" & chr(13) & _
	"	xlPosicaoMercado=xlCiclo+1;" & chr(13) & _
	"	xlQtdeVendida=xlPosicaoMercado+1;" & chr(13) & _
	"	xlQtdeEstimadaAVender=xlQtdeVendida+1;" & chr(13) & _
	"	xlQtdeEstoqueVenda=xlQtdeEstimadaAVender+1;" & chr(13) & _
	"	xlQtdeComprada=xlQtdeEstoqueVenda+1;" & chr(13) & _
	"	xlQtdeSaldo=xlQtdeComprada+1;" & chr(13)

'	ARRAY USADO P/ TRANSFERIR DADOS P/ O EXCEL
	strScript = strScript & _
	"	xlDadosMinIndex=(xlMargemEsq+1);" & chr(13) & _
	"	xlDadosMaxIndex=xlQtdeSaldo;" & chr(13)

'	CONFIGURA COLUNAS
	strScript = strScript & _
	"	oWS.Columns(xlMargemEsq).ColumnWidth=1;" & chr(13) & _
	"	oWS.Columns(xlFabricante).NumberFormat='@';" & chr(13) & _
	"	oWS.Columns(xlFabricante).WrapText=false;" & chr(13) & _
	"	oWS.Columns(xlFabricante).ColumnWidth=5;" & chr(13) & _
	"	oWS.Columns(xlProduto).NumberFormat='@';" & chr(13) & _
	"	oWS.Columns(xlProduto).WrapText=false;" & chr(13) & _
	"	oWS.Columns(xlProduto).ColumnWidth=8;" & chr(13) & _
	"	oWS.Columns(xlDescricao).NumberFormat='@';" & chr(13) & _
	"	oWS.Columns(xlDescricao).WrapText=false;" & chr(13) & _
	"	oWS.Columns(xlDescricao).ColumnWidth=50;" & chr(13) & _
	"	oWS.Columns(xlGrupo).NumberFormat='@';" & chr(13) & _
	"	oWS.Columns(xlGrupo).WrapText=false;" & chr(13) & _
	"	oWS.Columns(xlGrupo).HorizontalAlignment=xlCenter;" & chr(13) & _
	"	oWS.Columns(xlGrupo).ColumnWidth=6;" & chr(13) & _
	"	oWS.Columns(xlSubgrupo).NumberFormat='@';" & chr(13) & _
	"	oWS.Columns(xlSubgrupo).WrapText=false;" & chr(13) & _
	"	oWS.Columns(xlSubgrupo).HorizontalAlignment=xlCenter;" & chr(13) & _
	"	oWS.Columns(xlSubgrupo).ColumnWidth=10;" & chr(13) & _
	"	oWS.Columns(xlPotenciaBTU).NumberFormat='General';" & chr(13) & _
	"	oWS.Columns(xlPotenciaBTU).HorizontalAlignment=xlRight;" & chr(13) & _
	"	oWS.Columns(xlPotenciaBTU).WrapText=false;" & chr(13) & _
	"	oWS.Columns(xlPotenciaBTU).ColumnWidth=10;" & chr(13) & _
	"	oWS.Columns(xlCiclo).NumberFormat='@';" & chr(13) & _
	"	oWS.Columns(xlCiclo).WrapText=false;" & chr(13) & _
	"	oWS.Columns(xlCiclo).HorizontalAlignment=xlCenter;" & chr(13) & _
	"	oWS.Columns(xlCiclo).ColumnWidth=6;" & chr(13) & _
	"	oWS.Columns(xlPosicaoMercado).NumberFormat='@';" & chr(13) & _
	"	oWS.Columns(xlPosicaoMercado).WrapText=false;" & chr(13) & _
	"	oWS.Columns(xlPosicaoMercado).HorizontalAlignment=xlCenter;" & chr(13) & _
	"	oWS.Columns(xlPosicaoMercado).ColumnWidth=12;" & chr(13) & _
	"	oWS.Columns(xlQtdeVendida).NumberFormat='General';" & chr(13) & _
	"	oWS.Columns(xlQtdeVendida).WrapText=false;" & chr(13) & _
	"	oWS.Columns(xlQtdeVendida).HorizontalAlignment=xlRight;" & chr(13) & _
	"	oWS.Columns(xlQtdeVendida).Font.Bold=true;" & chr(13) & _
	"	oWS.Columns(xlQtdeVendida).ColumnWidth=10;" & chr(13) & _
	"	oWS.Columns(xlQtdeEstimadaAVender).NumberFormat='General';" & chr(13) & _
	"	oWS.Columns(xlQtdeEstimadaAVender).WrapText=false;" & chr(13) & _
	"	oWS.Columns(xlQtdeEstimadaAVender).HorizontalAlignment=xlRight;" & chr(13) & _
	"	oWS.Columns(xlQtdeEstimadaAVender).Font.Bold=true;" & chr(13) & _
	"	oWS.Columns(xlQtdeEstimadaAVender).ColumnWidth=10;" & chr(13) & _
	"	oWS.Columns(xlQtdeEstoqueVenda).NumberFormat='General';" & chr(13) & _
	"	oWS.Columns(xlQtdeEstoqueVenda).WrapText=false;" & chr(13) & _
	"	oWS.Columns(xlQtdeEstoqueVenda).HorizontalAlignment=xlRight;" & chr(13) & _
	"	oWS.Columns(xlQtdeEstoqueVenda).Font.Bold=true;" & chr(13) & _
	"	oWS.Columns(xlQtdeEstoqueVenda).ColumnWidth=10;" & chr(13) & _
	"	oWS.Columns(xlQtdeComprada).NumberFormat='General';" & chr(13) & _
	"	oWS.Columns(xlQtdeComprada).WrapText=false;" & chr(13) & _
	"	oWS.Columns(xlQtdeComprada).HorizontalAlignment=xlRight;" & chr(13) & _
	"	oWS.Columns(xlQtdeComprada).Font.Bold=true;" & chr(13) & _
	"	oWS.Columns(xlQtdeComprada).ColumnWidth=10;" & chr(13) & _
	"	oWS.Columns(xlQtdeSaldo).NumberFormat='General';" & chr(13) & _
	"	oWS.Columns(xlQtdeSaldo).WrapText=false;" & chr(13) & _
	"	oWS.Columns(xlQtdeSaldo).HorizontalAlignment=xlRight;" & chr(13) & _
	"	oWS.Columns(xlQtdeSaldo).Font.Bold=true;" & chr(13) & _
	"	oWS.Columns(xlQtdeSaldo).ColumnWidth=10;" & chr(13)
	
'	CAMPOS DO CABEÇALHO
	strScript = strScript & _
	"	xlNumLinha=1;" & chr(13) & _
	"	oWS.Rows(xlNumLinha).RowHeight=1;" & chr(13) & _
	"	xlNumLinha++;" & chr(13) & _
	"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlFabricante);" & chr(13) & _
	"	oRange=oWS.Range(xlNumLinhaAux + xlNumLinha);" & chr(13) & _
	"	oRange.WrapText=false;" & chr(13) & _
	"	oFont=oRange.Font;" & chr(13) & _
	"	oFont.Name=xlFN_LISTAGEM;" & chr(13) & _
	"	oFont.Size=xlFS_CABECALHO;" & chr(13) & _
	"	oFont.Bold=true;" & chr(13) & _
	"	oRange.Value='Farol Resumido';" & chr(13)
	
	s = ""
	s_aux = c_dt_periodo_inicio
	if s_aux = "" then s_aux = "N.I."
	s = s & s_aux & " a "
	s_aux = c_dt_periodo_termino
	if s_aux = "" then s_aux = "N.I."
	s = s & s_aux
	strScript = strScript & _
	"	xlNumLinha++;" & chr(13) & _
	"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlFabricante);" & chr(13) & _
	"	oRange=oWS.Range(xlNumLinhaAux + xlNumLinha);" & chr(13) & _
	"	oRange.WrapText=false;" & chr(13) & _
	"	oFont=oRange.Font;" & chr(13) & _
	"	oFont.Name=xlFN_LISTAGEM;" & chr(13) & _
	"	oFont.Size=xlFS_CABECALHO-2;" & chr(13) & _
	"	oFont.Bold=true;" & chr(13) & _
	"	oRange.Value='Período de vendas: " & s & "';" & chr(13)
	
	s = c_fabricante
	if s = "" then s = "N.I."
	strScript = strScript & _
	"	xlNumLinha++;" & chr(13) & _
	"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlFabricante);" & chr(13) & _
	"	oRange=oWS.Range(xlNumLinhaAux + xlNumLinha);" & chr(13) & _
	"	oRange.WrapText=false;" & chr(13) & _
	"	oFont=oRange.Font;" & chr(13) & _
	"	oFont.Name=xlFN_LISTAGEM;" & chr(13) & _
	"	oFont.Size=xlFS_CABECALHO-2;" & chr(13) & _
	"	oFont.Bold=true;" & chr(13) & _
	"	oRange.Value='Fabricante: " & s & "';" & chr(13)
	
	s = c_grupo
	if s = "" then s = "N.I."
	strScript = strScript & _
	"	xlNumLinha++;" & chr(13) & _
	"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlFabricante);" & chr(13) & _
	"	oRange=oWS.Range(xlNumLinhaAux + xlNumLinha);" & chr(13) & _
	"	oRange.WrapText=false;" & chr(13) & _
	"	oFont=oRange.Font;" & chr(13) & _
	"	oFont.Name=xlFN_LISTAGEM;" & chr(13) & _
	"	oFont.Size=xlFS_CABECALHO-2;" & chr(13) & _
	"	oFont.Bold=true;" & chr(13) & _
	"	oRange.Value='Grupo de produtos: " & s & "';" & chr(13)

	s = c_subgrupo
	if s = "" then s = "N.I."
	strScript = strScript & _
	"	xlNumLinha++;" & chr(13) & _
	"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlFabricante);" & chr(13) & _
	"	oRange=oWS.Range(xlNumLinhaAux + xlNumLinha);" & chr(13) & _
	"	oRange.WrapText=false;" & chr(13) & _
	"	oFont=oRange.Font;" & chr(13) & _
	"	oFont.Name=xlFN_LISTAGEM;" & chr(13) & _
	"	oFont.Size=xlFS_CABECALHO-2;" & chr(13) & _
	"	oFont.Bold=true;" & chr(13) & _
	"	oRange.Value='Subgrupo de produtos: " & s & "';" & chr(13)
	
	s = c_potencia_BTU
	if s = "" then s = "N.I." else s = formata_inteiro(s)
	strScript = strScript & _
	"	xlNumLinha++;" & chr(13) & _
	"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlFabricante);" & chr(13) & _
	"	oRange=oWS.Range(xlNumLinhaAux + xlNumLinha);" & chr(13) & _
	"	oRange.WrapText=false;" & chr(13) & _
	"	oFont=oRange.Font;" & chr(13) & _
	"	oFont.Name=xlFN_LISTAGEM;" & chr(13) & _
	"	oFont.Size=xlFS_CABECALHO-2;" & chr(13) & _
	"	oFont.Bold=true;" & chr(13) & _
	"	oRange.Value='BTU/h: " & s & "';" & chr(13)
	
	s = c_ciclo
	if s = "" then s = "N.I."
	strScript = strScript & _
	"	xlNumLinha++;" & chr(13) & _
	"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlFabricante);" & chr(13) & _
	"	oRange=oWS.Range(xlNumLinhaAux + xlNumLinha);" & chr(13) & _
	"	oRange.WrapText=false;" & chr(13) & _
	"	oFont=oRange.Font;" & chr(13) & _
	"	oFont.Name=xlFN_LISTAGEM;" & chr(13) & _
	"	oFont.Size=xlFS_CABECALHO-2;" & chr(13) & _
	"	oFont.Bold=true;" & chr(13) & _
	"	oRange.Value='Ciclo: " & s & "';" & chr(13)
	
	s = c_posicao_mercado
	if s = "" then s = "N.I."
	strScript = strScript & _
	"	xlNumLinha++;" & chr(13) & _
	"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlFabricante);" & chr(13) & _
	"	oRange=oWS.Range(xlNumLinhaAux + xlNumLinha);" & chr(13) & _
	"	oRange.WrapText=false;" & chr(13) & _
	"	oFont=oRange.Font;" & chr(13) & _
	"	oFont.Name=xlFN_LISTAGEM;" & chr(13) & _
	"	oFont.Size=xlFS_CABECALHO-2;" & chr(13) & _
	"	oFont.Bold=true;" & chr(13) & _
	"	oRange.Value='Posição Mercado: " & s & "';" & chr(13)
	
	strScript = strScript & _
	"	xlNumLinha++;" & chr(13) & _
	"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlFabricante);" & chr(13) & _
	"	oRange=oWS.Range(xlNumLinhaAux + xlNumLinha);" & chr(13) & _
	"	oRange.WrapText=false;" & chr(13) & _
	"	oFont=oRange.Font;" & chr(13) & _
	"	oFont.Name=xlFN_LISTAGEM;" & chr(13) & _
	"	oFont.Size=xlFS_CABECALHO-2;" & chr(13) & _
	"	oFont.Bold=true;" & chr(13) & _
	"	oRange.Value='Emissão: " & formata_data_hora_sem_seg(Now) & "';" & chr(13) & _
	"	xlNumLinha++;" & chr(13) & _
	"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlQtdeSaldo);" & chr(13) & _
	"	xlCellPercEstCresc=xlNumLinhaAux + xlNumLinha;" & chr(13) & _
	"	oRange=oWS.Range(xlNumLinhaAux + xlNumLinha);" & chr(13) & _
	"	oRange.WrapText=false;" & chr(13) & _
	"	oFont=oRange.Font;" & chr(13) & _
	"	oFont.Name=xlFN_LISTAGEM;" & chr(13) & _
	"	oFont.Size=xlFS_CABECALHO-2;" & chr(13) & _
	"	oFont.Bold=true;" & chr(13) & _
	"	oFont.Color=-16751104;" & chr(13) & _
	"	oRange.Value=" & bd_formata_numero(perc_est_cresc) & ";" & chr(13) & _
	"	oRange.NumberFormat='0.0;[Red]-0.0';" & chr(13) & _
	"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlQtdeComprada);" & chr(13) & _
	"	oRange=oWS.Range(xlNumLinhaAux + xlNumLinha);" & chr(13) & _
	"	oRange.WrapText=false;" & chr(13) & _
	"	oRange.HorizontalAlignment=xlRight;" & chr(13) & _
	"	oFont=oRange.Font;" & chr(13) & _
	"	oFont.Name=xlFN_LISTAGEM;" & chr(13) & _
	"	oFont.Size=xlFS_CABECALHO-2;" & chr(13) & _
	"	oFont.Bold=true;" & chr(13) & _
	"	oRange.Value='Percentual estimado de crescimento (%):';" & chr(13)
	
'	BORDAS P/ OS TÍTULOS DAS COLUNAS
	strScript = strScript & _
	"	xlNumLinha++;" & chr(13) & _
	"	oWS.Rows(xlNumLinha).RowHeight=4;" & chr(13) & _
	"	xlNumLinha++;" & chr(13) & _
	"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlDadosMinIndex) + xlNumLinha.toString() + ':' + excel_converte_numeracao_digito_para_letra(xlDadosMaxIndex) + xlNumLinha.toString();" & chr(13) & _
	"	oRange=oWS.Range(xlNumLinhaAux);" & chr(13) & _
	"	oBorders=oRange.Borders(xlEdgeTop);" & chr(13) & _
	"	oBorders.LineStyle=xlContinuous;" & chr(13) & _
	"	oBorders.Weight=xlMedium;" & chr(13) & _
	"	oBorders.ColorIndex=xlAutomatic;" & chr(13) & _
	"	oBorders=oRange.Borders(xlEdgeBottom);" & chr(13) & _
	"	oBorders.LineStyle=xlContinuous;" & chr(13) & _
	"	oBorders.Weight=xlMedium;" & chr(13) & _
	"	oBorders.ColorIndex=xlAutomatic;" & chr(13)

'	TÍTULOS DA COLUNAS
	strScript = strScript & _
	"	oRange.WrapText=true;" & chr(13) & _
	"	oRange.VerticalAlignment=xlVAlignBottom;" & chr(13) & _
	"	oFont=oRange.Font;" & chr(13) & _
	"	oFont.Name=xlFN_LISTAGEM;" & chr(13) & _
	"	oFont.Size=xlFS_LISTAGEM;" & chr(13) & _
	"	oFont.Bold=true;" & chr(13) & _
	"	oFont.Italic=false;" & chr(13) & _
	"	oRange=oWS.Range(excel_converte_numeracao_digito_para_letra(xlFabricante) + xlNumLinha.toString());" & chr(13) & _
	"	oRange.Value='Fabr';" & chr(13) & _
	"	oRange=oWS.Range(excel_converte_numeracao_digito_para_letra(xlProduto) + xlNumLinha.toString());" & chr(13) & _
	"	oRange.Value='Produto';" & chr(13) & _
	"	oRange=oWS.Range(excel_converte_numeracao_digito_para_letra(xlDescricao) + xlNumLinha.toString());" & chr(13) & _
	"	oRange.Value='Descrição';" & chr(13) & _
	"	oRange=oWS.Range(excel_converte_numeracao_digito_para_letra(xlGrupo) + xlNumLinha.toString());" & chr(13) & _
	"	oRange.Value='Grupo';" & chr(13) & _
	"	oRange=oWS.Range(excel_converte_numeracao_digito_para_letra(xlSubgrupo) + xlNumLinha.toString());" & chr(13) & _
	"	oRange.Value='Subgrupo';" & chr(13) & _
	"	oRange=oWS.Range(excel_converte_numeracao_digito_para_letra(xlPotenciaBTU) + xlNumLinha.toString());" & chr(13) & _
	"	oRange.Value='BTU/h';" & chr(13) & _
	"	oRange=oWS.Range(excel_converte_numeracao_digito_para_letra(xlCiclo) + xlNumLinha.toString());" & chr(13) & _
	"	oRange.Value='Ciclo';" & chr(13) & _
	"	oRange=oWS.Range(excel_converte_numeracao_digito_para_letra(xlPosicaoMercado) + xlNumLinha.toString());" & chr(13) & _
	"	oRange.Value='Posição Mercado';" & chr(13) & _
	"	oRange=oWS.Range(excel_converte_numeracao_digito_para_letra(xlQtdeVendida) + xlNumLinha.toString());" & chr(13) & _
	"	oRange.Value='Venda (Histórico)';" & chr(13) & _
	"	oRange=oWS.Range(excel_converte_numeracao_digito_para_letra(xlQtdeEstimadaAVender) + xlNumLinha.toString());" & chr(13) & _
	"	oRange.Value='Venda (Previsão)';" & chr(13) & _
	"	oRange=oWS.Range(excel_converte_numeracao_digito_para_letra(xlQtdeEstoqueVenda) + xlNumLinha.toString());" & chr(13) & _
	"	oRange.Value='Estoque';" & chr(13) & _
	"	oRange=oWS.Range(excel_converte_numeracao_digito_para_letra(xlQtdeComprada) + xlNumLinha.toString());" & chr(13) & _
	"	oRange.Value='Comprado';" & chr(13) & _
	"	oRange=oWS.Range(excel_converte_numeracao_digito_para_letra(xlQtdeSaldo) + xlNumLinha.toString());" & chr(13) & _
	"	oRange.Value='Saldo';" & chr(13)

'	BORDAS VERTICAIS
	strScript = strScript & _
	"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlFabricante) + xlNumLinha.toString();" & chr(13) & _
	"	oRange=oWS.Range(xlNumLinhaAux);" & chr(13) & _
	"	oBorders=oRange.Borders(xlEdgeLeft);" & chr(13) & _
	"	oBorders.LineStyle=xlContinuous;" & chr(13) & _
	"	oBorders.Weight=xlThin;" & chr(13) & _
	"	oBorders.ColorIndex=xlAutomatic;" & chr(13) & _
	"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlProduto) + xlNumLinha.toString();" & chr(13) & _
	"	oRange=oWS.Range(xlNumLinhaAux);" & chr(13) & _
	"	oBorders=oRange.Borders(xlEdgeLeft);" & chr(13) & _
	"	oBorders.LineStyle=xlContinuous;" & chr(13) & _
	"	oBorders.Weight=xlThin;" & chr(13) & _
	"	oBorders.ColorIndex=xlAutomatic;" & chr(13) & _
	"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlDescricao) + xlNumLinha.toString();" & chr(13) & _
	"	oRange=oWS.Range(xlNumLinhaAux);" & chr(13) & _
	"	oBorders=oRange.Borders(xlEdgeLeft);" & chr(13) & _
	"	oBorders.LineStyle=xlContinuous;" & chr(13) & _
	"	oBorders.Weight=xlThin;" & chr(13) & _
	"	oBorders.ColorIndex=xlAutomatic;" & chr(13) & _
	"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlGrupo) + xlNumLinha.toString();" & chr(13) & _
	"	oRange=oWS.Range(xlNumLinhaAux);" & chr(13) & _
	"	oBorders=oRange.Borders(xlEdgeLeft);" & chr(13) & _
	"	oBorders.LineStyle=xlContinuous;" & chr(13) & _
	"	oBorders.Weight=xlThin;" & chr(13) & _
	"	oBorders.ColorIndex=xlAutomatic;" & chr(13) & _
	"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlSubgrupo) + xlNumLinha.toString();" & chr(13) & _
	"	oRange=oWS.Range(xlNumLinhaAux);" & chr(13) & _
	"	oBorders=oRange.Borders(xlEdgeLeft);" & chr(13) & _
	"	oBorders.LineStyle=xlContinuous;" & chr(13) & _
	"	oBorders.Weight=xlThin;" & chr(13) & _
	"	oBorders.ColorIndex=xlAutomatic;" & chr(13) & _
	"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlPotenciaBTU) + xlNumLinha.toString();" & chr(13) & _
	"	oRange=oWS.Range(xlNumLinhaAux);" & chr(13) & _
	"	oBorders=oRange.Borders(xlEdgeLeft);" & chr(13) & _
	"	oBorders.LineStyle=xlContinuous;" & chr(13) & _
	"	oBorders.Weight=xlThin;" & chr(13) & _
	"	oBorders.ColorIndex=xlAutomatic;" & chr(13) & _
	"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlCiclo) + xlNumLinha.toString();" & chr(13) & _
	"	oRange=oWS.Range(xlNumLinhaAux);" & chr(13) & _
	"	oBorders=oRange.Borders(xlEdgeLeft);" & chr(13) & _
	"	oBorders.LineStyle=xlContinuous;" & chr(13) & _
	"	oBorders.Weight=xlThin;" & chr(13) & _
	"	oBorders.ColorIndex=xlAutomatic;" & chr(13) & _
	"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlPosicaoMercado) + xlNumLinha.toString();" & chr(13) & _
	"	oRange=oWS.Range(xlNumLinhaAux);" & chr(13) & _
	"	oBorders=oRange.Borders(xlEdgeLeft);" & chr(13) & _
	"	oBorders.LineStyle=xlContinuous;" & chr(13) & _
	"	oBorders.Weight=xlThin;" & chr(13) & _
	"	oBorders.ColorIndex=xlAutomatic;" & chr(13) & _
	"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlQtdeVendida) + xlNumLinha.toString();" & chr(13) & _
	"	oRange=oWS.Range(xlNumLinhaAux);" & chr(13) & _
	"	oBorders=oRange.Borders(xlEdgeLeft);" & chr(13) & _
	"	oBorders.LineStyle=xlContinuous;" & chr(13) & _
	"	oBorders.Weight=xlThin;" & chr(13) & _
	"	oBorders.ColorIndex=xlAutomatic;" & chr(13) & _
	"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlQtdeEstimadaAVender) + xlNumLinha.toString();" & chr(13) & _
	"	oRange=oWS.Range(xlNumLinhaAux);" & chr(13) & _
	"	oBorders=oRange.Borders(xlEdgeLeft);" & chr(13) & _
	"	oBorders.LineStyle=xlContinuous;" & chr(13) & _
	"	oBorders.Weight=xlThin;" & chr(13) & _
	"	oBorders.ColorIndex=xlAutomatic;" & chr(13) & _
	"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlQtdeEstoqueVenda) + xlNumLinha.toString();" & chr(13) & _
	"	oRange=oWS.Range(xlNumLinhaAux);" & chr(13) & _
	"	oBorders=oRange.Borders(xlEdgeLeft);" & chr(13) & _
	"	oBorders.LineStyle=xlContinuous;" & chr(13) & _
	"	oBorders.Weight=xlThin;" & chr(13) & _
	"	oBorders.ColorIndex=xlAutomatic;" & chr(13) & _
	"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlQtdeComprada) + xlNumLinha.toString();" & chr(13) & _
	"	oRange=oWS.Range(xlNumLinhaAux);" & chr(13) & _
	"	oBorders=oRange.Borders(xlEdgeLeft);" & chr(13) & _
	"	oBorders.LineStyle=xlContinuous;" & chr(13) & _
	"	oBorders.Weight=xlThin;" & chr(13) & _
	"	oBorders.ColorIndex=xlAutomatic;" & chr(13) & _
	"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlQtdeSaldo) + xlNumLinha.toString();" & chr(13) & _
	"	oRange=oWS.Range(xlNumLinhaAux);" & chr(13) & _
	"	oBorders=oRange.Borders(xlEdgeLeft);" & chr(13) & _
	"	oBorders.LineStyle=xlContinuous;" & chr(13) & _
	"	oBorders.Weight=xlThin;" & chr(13) & _
	"	oBorders.ColorIndex=xlAutomatic;" & chr(13) & _
	"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlQtdeSaldo) + xlNumLinha.toString();" & chr(13) & _
	"	oRange=oWS.Range(xlNumLinhaAux);" & chr(13) & _
	"	oBorders=oRange.Borders(xlEdgeRight);" & chr(13) & _
	"	oBorders.LineStyle=xlContinuous;" & chr(13) & _
	"	oBorders.Weight=xlThin;" & chr(13) & _
	"	oBorders.ColorIndex=xlAutomatic;" & chr(13)
	
'	CONGELA BLOCO SUPERIOR
	strScript = strScript & _
	"	i_dados_inicio=xlNumLinha+1;" & chr(13) & _
	"	oRange=oWS.Range('A' + i_dados_inicio.toString());" & chr(13) & _
	"	oRange.Select;" & chr(13) & _
	"	oXL.ActiveWindow.FreezePanes=true;" & chr(13)
	
'	EXIBE NÚMEROS FORMATADOS COMO INTEIRO E NEGATIVOS EM VERMELHO
	strScript = strScript & _
		"	i_dados_inicio=xlNumLinha+1;" & chr(13) & _
		"	n=65536;" & chr(13) & _
		"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlPotenciaBTU) + i_dados_inicio.toString() + ':' + excel_converte_numeracao_digito_para_letra(xlPotenciaBTU) + n.toString();" & chr(13) & _
		"	oRange=oWS.Range(xlNumLinhaAux);" & chr(13) & _
		"	oRange.NumberFormat='#,##0;[Red]-#,##0';" & chr(13) & _
		"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlQtdeVendida) + i_dados_inicio.toString() + ':' + excel_converte_numeracao_digito_para_letra(xlQtdeSaldo) + n.toString();" & chr(13) & _
		"	oRange=oWS.Range(xlNumLinhaAux);" & chr(13) & _
		"	oRange.NumberFormat='#,##0;[Red]-#,##0';" & chr(13)
	
'	EXIBE NÚMEROS DIFERENTES DE ZERO EM NEGRITO E ZEROS C/ FONTE NORMAL
	strScript = strScript & _
		"	i_dados_inicio=xlNumLinha+1;" & chr(13) & _
		"	n=65536;" & chr(13) & _
		"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlQtdeVendida) + i_dados_inicio.toString() + ':' + excel_converte_numeracao_digito_para_letra(xlQtdeSaldo) + n.toString();" & chr(13) & _
		"	oRange=oWS.Range(xlNumLinhaAux);" & chr(13) & _
		"	oRange.FormatConditions.Add(xlCellValue, xlEqual, ""=0"");" & chr(13) & _
		"//  Método SetFirstPriority existe somente a partir da versão 2007 (versão 12)" & chr(13) & _
		"	if (nExcelMajorVersion >= 12) oRange.FormatConditions(oRange.FormatConditions.Count).SetFirstPriority();" & chr(13) & _
		"	oRange.FormatConditions(1).Font.Bold=false;" & chr(13) & _
		"//  Propriedade StopIfTrue existe somente a partir da versão 2007 (versão 12)" & chr(13) & _
		"	if (nExcelMajorVersion >= 12) oRange.FormatConditions(1).StopIfTrue = false;" & chr(13) & _
		"	oRange.FormatConditions.Add(xlCellValue, xlNotEqual, ""=0"");" & chr(13) & _
		"//  Método SetFirstPriority existe somente a partir da versão 2007 (versão 12)" & chr(13) & _
		"	if (nExcelMajorVersion >= 12) oRange.FormatConditions(oRange.FormatConditions.Count).SetFirstPriority();" & chr(13) & _
		"	oRange.FormatConditions(1).Font.Bold=true;" & chr(13) & _
		"//  Propriedade StopIfTrue existe somente a partir da versão 2007 (versão 12)" & chr(13) & _
		"	if (nExcelMajorVersion >= 12) oRange.FormatConditions(1).StopIfTrue = false;" & chr(13)
	
	Response.Write strScript
	strScript = ""
	
	
'	LAÇO P/ LEITURA DOS DADOS DO BD
'	~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
	n_reg = 0
	n_reg_total = 0
	
	set r = cn.execute(s_sql)
	do while Not r.Eof
		
	'	CONTAGEM
		n_reg = n_reg + 1
		n_reg_total = n_reg_total + 1
		
		strScript = strScript & _
			"	montaLinha('" & Trim("" & r("fabricante")) & "','" & Trim("" & r("produto")) & "','" & substitui_caracteres(Trim("" & r("descricao")), "'", "\'") & "','" & substitui_caracteres(Trim("" & r("descricao_html")), "'", "\'") & "','" & Trim("" & r("grupo")) & "','" & Trim("" & r("subgrupo")) & "', "  & Trim("" & r("potencia_BTU")) & ",'" & Trim("" & r("ciclo")) & "','" & Trim("" & r("posicao_mercado")) & "', " & Cstr(r("qtde_vendida") - r("qtde_devolvida")) & ", " & Cstr(r("qtde_estoque_venda")) & ", " & Cstr(r("farol_qtde_comprada")) & ");" & chr(13)
		
		if n_reg = 1 then
			strScript = strScript & _
				"	i_dados_inicio = xlNumLinha;" & chr(13)
			end if
		
		Response.Write strScript
		strScript = ""
		
		r.MoveNext
		loop
	
	strScript = strScript & _
		"	i_dados_fim = xlNumLinha;" & chr(13)
	
'	BORDAS VERTICAIS
	strScript = strScript & _
	"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlFabricante) + i_dados_inicio.toString() + ':' + excel_converte_numeracao_digito_para_letra(xlFabricante) + i_dados_fim.toString();" & chr(13) & _
	"	oRange=oWS.Range(xlNumLinhaAux);" & chr(13) & _
	"	oBorders=oRange.Borders(xlEdgeLeft);" & chr(13) & _
	"	oBorders.LineStyle=xlContinuous;" & chr(13) & _
	"	oBorders.Weight=xlThin;" & chr(13) & _
	"	oBorders.ColorIndex=xlAutomatic;" & chr(13) & _
	"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlProduto) + i_dados_inicio.toString() + ':' + excel_converte_numeracao_digito_para_letra(xlProduto) + i_dados_fim.toString();" & chr(13) & _
	"	oRange=oWS.Range(xlNumLinhaAux);" & chr(13) & _
	"	oBorders=oRange.Borders(xlEdgeLeft);" & chr(13) & _
	"	oBorders.LineStyle=xlContinuous;" & chr(13) & _
	"	oBorders.Weight=xlThin;" & chr(13) & _
	"	oBorders.ColorIndex=xlAutomatic;" & chr(13) & _
	"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlDescricao) + i_dados_inicio.toString() + ':' + excel_converte_numeracao_digito_para_letra(xlDescricao) + i_dados_fim.toString();" & chr(13) & _
	"	oRange=oWS.Range(xlNumLinhaAux);" & chr(13) & _
	"	oBorders=oRange.Borders(xlEdgeLeft);" & chr(13) & _
	"	oBorders.LineStyle=xlContinuous;" & chr(13) & _
	"	oBorders.Weight=xlThin;" & chr(13) & _
	"	oBorders.ColorIndex=xlAutomatic;" & chr(13) & _
	"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlGrupo) + i_dados_inicio.toString() + ':' + excel_converte_numeracao_digito_para_letra(xlGrupo) + i_dados_fim.toString();" & chr(13) & _
	"	oRange=oWS.Range(xlNumLinhaAux);" & chr(13) & _
	"	oBorders=oRange.Borders(xlEdgeLeft);" & chr(13) & _
	"	oBorders.LineStyle=xlContinuous;" & chr(13) & _
	"	oBorders.Weight=xlThin;" & chr(13) & _
	"	oBorders.ColorIndex=xlAutomatic;" & chr(13) & _
	"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlSubgrupo) + i_dados_inicio.toString() + ':' + excel_converte_numeracao_digito_para_letra(xlSubgrupo) + i_dados_fim.toString();" & chr(13) & _
	"	oRange=oWS.Range(xlNumLinhaAux);" & chr(13) & _
	"	oBorders=oRange.Borders(xlEdgeLeft);" & chr(13) & _
	"	oBorders.LineStyle=xlContinuous;" & chr(13) & _
	"	oBorders.Weight=xlThin;" & chr(13) & _
	"	oBorders.ColorIndex=xlAutomatic;" & chr(13) & _
	"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlPotenciaBTU) + i_dados_inicio.toString() + ':' + excel_converte_numeracao_digito_para_letra(xlPotenciaBTU) + i_dados_fim.toString();" & chr(13) & _
	"	oRange=oWS.Range(xlNumLinhaAux);" & chr(13) & _
	"	oBorders=oRange.Borders(xlEdgeLeft);" & chr(13) & _
	"	oBorders.LineStyle=xlContinuous;" & chr(13) & _
	"	oBorders.Weight=xlThin;" & chr(13) & _
	"	oBorders.ColorIndex=xlAutomatic;" & chr(13) & _
	"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlCiclo) + i_dados_inicio.toString() + ':' + excel_converte_numeracao_digito_para_letra(xlCiclo) + i_dados_fim.toString();" & chr(13) & _
	"	oRange=oWS.Range(xlNumLinhaAux);" & chr(13) & _
	"	oBorders=oRange.Borders(xlEdgeLeft);" & chr(13) & _
	"	oBorders.LineStyle=xlContinuous;" & chr(13) & _
	"	oBorders.Weight=xlThin;" & chr(13) & _
	"	oBorders.ColorIndex=xlAutomatic;" & chr(13) & _
	"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlPosicaoMercado) + i_dados_inicio.toString() + ':' + excel_converte_numeracao_digito_para_letra(xlPosicaoMercado) + i_dados_fim.toString();" & chr(13) & _
	"	oRange=oWS.Range(xlNumLinhaAux);" & chr(13) & _
	"	oBorders=oRange.Borders(xlEdgeLeft);" & chr(13) & _
	"	oBorders.LineStyle=xlContinuous;" & chr(13) & _
	"	oBorders.Weight=xlThin;" & chr(13) & _
	"	oBorders.ColorIndex=xlAutomatic;" & chr(13) & _
	"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlQtdeVendida) + i_dados_inicio.toString() + ':' + excel_converte_numeracao_digito_para_letra(xlQtdeVendida) + i_dados_fim.toString();" & chr(13) & _
	"	oRange=oWS.Range(xlNumLinhaAux);" & chr(13) & _
	"	oBorders=oRange.Borders(xlEdgeLeft);" & chr(13) & _
	"	oBorders.LineStyle=xlContinuous;" & chr(13) & _
	"	oBorders.Weight=xlThin;" & chr(13) & _
	"	oBorders.ColorIndex=xlAutomatic;" & chr(13) & _
	"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlQtdeEstimadaAVender) + i_dados_inicio.toString() + ':' + excel_converte_numeracao_digito_para_letra(xlQtdeEstimadaAVender) + i_dados_fim.toString();" & chr(13) & _
	"	oRange=oWS.Range(xlNumLinhaAux);" & chr(13) & _
	"	oBorders=oRange.Borders(xlEdgeLeft);" & chr(13) & _
	"	oBorders.LineStyle=xlContinuous;" & chr(13) & _
	"	oBorders.Weight=xlThin;" & chr(13) & _
	"	oBorders.ColorIndex=xlAutomatic;" & chr(13) & _
	"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlQtdeEstoqueVenda) + i_dados_inicio.toString() + ':' + excel_converte_numeracao_digito_para_letra(xlQtdeEstoqueVenda) + i_dados_fim.toString();" & chr(13) & _
	"	oRange=oWS.Range(xlNumLinhaAux);" & chr(13) & _
	"	oBorders=oRange.Borders(xlEdgeLeft);" & chr(13) & _
	"	oBorders.LineStyle=xlContinuous;" & chr(13) & _
	"	oBorders.Weight=xlThin;" & chr(13) & _
	"	oBorders.ColorIndex=xlAutomatic;" & chr(13) & _
	"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlQtdeComprada) + i_dados_inicio.toString() + ':' + excel_converte_numeracao_digito_para_letra(xlQtdeComprada) + i_dados_fim.toString();" & chr(13) & _
	"	oRange=oWS.Range(xlNumLinhaAux);" & chr(13) & _
	"	oBorders=oRange.Borders(xlEdgeLeft);" & chr(13) & _
	"	oBorders.LineStyle=xlContinuous;" & chr(13) & _
	"	oBorders.Weight=xlThin;" & chr(13) & _
	"	oBorders.ColorIndex=xlAutomatic;" & chr(13) & _
	"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlQtdeSaldo) + i_dados_inicio.toString() + ':' + excel_converte_numeracao_digito_para_letra(xlQtdeSaldo) + i_dados_fim.toString();" & chr(13) & _
	"	oRange=oWS.Range(xlNumLinhaAux);" & chr(13) & _
	"	oBorders=oRange.Borders(xlEdgeLeft);" & chr(13) & _
	"	oBorders.LineStyle=xlContinuous;" & chr(13) & _
	"	oBorders.Weight=xlThin;" & chr(13) & _
	"	oBorders.ColorIndex=xlAutomatic;" & chr(13) & _
	"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlQtdeSaldo) + i_dados_inicio.toString() + ':' + excel_converte_numeracao_digito_para_letra(xlQtdeSaldo) + i_dados_fim.toString();" & chr(13) & _
	"	oRange=oWS.Range(xlNumLinhaAux);" & chr(13) & _
	"	oBorders=oRange.Borders(xlEdgeRight);" & chr(13) & _
	"	oBorders.LineStyle=xlContinuous;" & chr(13) & _
	"	oBorders.Weight=xlThin;" & chr(13) & _
	"	oBorders.ColorIndex=xlAutomatic;" & chr(13)
	
'	LINHA SEPARADORA
	strScript = strScript & _
		"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlDadosMinIndex) + xlNumLinha.toString() + ':' + excel_converte_numeracao_digito_para_letra(xlDadosMaxIndex) + xlNumLinha.toString();" & chr(13) & _
		"	oRange=oWS.Range(xlNumLinhaAux);" & chr(13) & _
		"	oBorders=oRange.Borders(xlEdgeBottom);" & chr(13) & _
		"	oBorders.LineStyle=xlContinuous;" & chr(13) & _
		"	oBorders.Weight=xlMedium;" & chr(13) & _
		"	oBorders.ColorIndex=xlAutomatic;" & chr(13)
	
'	LINHA DE TOTAIS
	strScript = strScript & _
		"	xlNumLinha++;" & chr(13) & _
		"	oWS.Rows(xlNumLinha).RowHeight=6;" & chr(13) & _
		"	xlNumLinha++;" & chr(13) & _
		"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlPosicaoMercado) + xlNumLinha.toString();" & chr(13) & _
		"	oRange=oWS.Range(xlNumLinhaAux);" & chr(13) & _
		"	oRange.WrapText=false;" & chr(13) & _
		"	oRange.HorizontalAlignment=xlRight;" & chr(13) & _
		"	oFont=oRange.Font;" & chr(13) & _
		"	oFont.Name=xlFN_LISTAGEM;" & chr(13) & _
		"	oFont.Size=xlFS_LISTAGEM;" & chr(13) & _
		"	oFont.Bold=true;" & chr(13) & _
		"	oFont.Italic=false;" & chr(13) & _
		"	oRange.Value='TOTAL';" & _
		"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlQtdeVendida) + xlNumLinha.toString();" & chr(13) & _
		"	oRange=oWS.Range(xlNumLinhaAux);" & chr(13) & _
		"	oRange.WrapText=false;" & chr(13) & _
		"	oFont=oRange.Font;" & chr(13) & _
		"	oFont.Name=xlFN_LISTAGEM;" & chr(13) & _
		"	oFont.Size=xlFS_LISTAGEM;" & chr(13) & _
		"	oFont.Bold=true;" & chr(13) & _
		"	oFont.Italic=false;" & chr(13) & _
		"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlQtdeVendida) + i_dados_inicio.toString() + ':' + excel_converte_numeracao_digito_para_letra(xlQtdeVendida) + i_dados_fim.toString();" & chr(13) & _
		"	oRange.FormulaLocal='=SOMA(' + xlNumLinhaAux + ')';" & chr(13) & _
		"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlQtdeEstimadaAVender) + xlNumLinha.toString();" & chr(13) & _
		"	oRange=oWS.Range(xlNumLinhaAux);" & chr(13) & _
		"	oRange.WrapText=false;" & chr(13) & _
		"	oFont=oRange.Font;" & chr(13) & _
		"	oFont.Name=xlFN_LISTAGEM;" & chr(13) & _
		"	oFont.Size=xlFS_LISTAGEM;" & chr(13) & _
		"	oFont.Bold=true;" & chr(13) & _
		"	oFont.Italic=false;" & chr(13) & _
		"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlQtdeEstimadaAVender) + i_dados_inicio.toString() + ':' + excel_converte_numeracao_digito_para_letra(xlQtdeEstimadaAVender) + i_dados_fim.toString();" & chr(13) & _
		"	oRange.FormulaLocal='=SOMA(' + xlNumLinhaAux + ')';" & chr(13) & _
		"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlQtdeEstoqueVenda) + xlNumLinha.toString();" & chr(13) & _
		"	oRange=oWS.Range(xlNumLinhaAux);" & chr(13) & _
		"	oRange.WrapText=false;" & chr(13) & _
		"	oFont=oRange.Font;" & chr(13) & _
		"	oFont.Name=xlFN_LISTAGEM;" & chr(13) & _
		"	oFont.Size=xlFS_LISTAGEM;" & chr(13) & _
		"	oFont.Bold=true;" & chr(13) & _
		"	oFont.Italic=false;" & chr(13) & _
		"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlQtdeEstoqueVenda) + i_dados_inicio.toString() + ':' + excel_converte_numeracao_digito_para_letra(xlQtdeEstoqueVenda) + i_dados_fim.toString();" & chr(13) & _
		"	oRange.FormulaLocal='=SOMA(' + xlNumLinhaAux + ')';" & chr(13) & _
		"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlQtdeComprada) + xlNumLinha.toString();" & chr(13) & _
		"	oRange=oWS.Range(xlNumLinhaAux);" & chr(13) & _
		"	oRange.WrapText=false;" & chr(13) & _
		"	oFont=oRange.Font;" & chr(13) & _
		"	oFont.Name=xlFN_LISTAGEM;" & chr(13) & _
		"	oFont.Size=xlFS_LISTAGEM;" & chr(13) & _
		"	oFont.Bold=true;" & chr(13) & _
		"	oFont.Italic=false;" & chr(13) & _
		"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlQtdeComprada) + i_dados_inicio.toString() + ':' + excel_converte_numeracao_digito_para_letra(xlQtdeComprada) + i_dados_fim.toString();" & chr(13) & _
		"	oRange.FormulaLocal='=SOMA(' + xlNumLinhaAux + ')';" & chr(13) & _
		"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlQtdeSaldo) + xlNumLinha.toString();" & chr(13) & _
		"	oRange=oWS.Range(xlNumLinhaAux);" & chr(13) & _
		"	oRange.WrapText=false;" & chr(13) & _
		"	oFont=oRange.Font;" & chr(13) & _
		"	oFont.Name=xlFN_LISTAGEM;" & chr(13) & _
		"	oFont.Size=xlFS_LISTAGEM;" & chr(13) & _
		"	oFont.Bold=true;" & chr(13) & _
		"	oFont.Italic=false;" & chr(13) & _
		"	xlNumLinhaAux=excel_converte_numeracao_digito_para_letra(xlQtdeSaldo) + i_dados_inicio.toString() + ':' + excel_converte_numeracao_digito_para_letra(xlQtdeSaldo) + i_dados_fim.toString();" & chr(13) & _
		"	oRange.FormulaLocal='=SOMA(' + xlNumLinhaAux + ')';" & chr(13)
	
'	AJUSTES FINAIS NA PLANILHA
	strScript = strScript & _
		"	oXL.Windows(oWB.Name).Activate;" & chr(13) & _
		"	oXL.Sheets(1).Select;" & chr(13) & _
		"	oXL.Sheets(1).Range('A'+i_dados_inicio.toString()).Select;" & chr(13) & _
		"	oXL.Sheets(1).Range('A1').Select;" & chr(13) & _
		"	$(""#textoMensagem"").text(""A planilha foi gerada com sucesso!!"");" & chr(13) & _
		"	$(""#divMsgAguarde"").removeClass(""modalAguarde"");" & chr(13) & _
		"	$(""#divMsgAguarde"").addClass(""modalOk"");" & chr(13) & _
		"	$(""#btnFechar"").show(""fast"");" & chr(13)
	
	strScript = strScript & _
		"	}" & chr(13) & _
		"catch (e) {" & chr(13) & _
		"		$(""#textoMensagem"").text(""Ocorreu um erro durante a geração da planilha: "" + e.message);" & chr(13) & _
		"		$(""#divMsgAguarde"").removeClass(""modalAguarde"");" & chr(13) & _
		"		$(""#divMsgAguarde"").addClass(""modalErro"");" & chr(13) & _
		"		$(""#btnFechar"").show(""fast"");" & chr(13) & _
		"		alert('Ocorreu um erro durante a geração da planilha: ' + e.message);" & chr(13) & _
		"		return;" & chr(13) & _
		"	}" & chr(13) 
	
'	FECHAMENTO DO SCRIPT
	strScript = strScript & _
		"}" & chr(13) & _
		"</script>" & chr(13)

	Response.Write strScript

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

<link href="<%=URL_FILE__E_CSS%>" rel="stylesheet" type="text/css">

<style type="text/css">
.modalBase {
	display: none;
	position:   relative;
	z-index:    1000;
	top:        0;
	left:       0;
	height:     200px;
	width:      649px;
	background-color: rgb( 255, 255, 255 );
	opacity: 1.0;
	filter:Alpha(opacity=100);
	background-position: 50% 70%;
	background-repeat: no-repeat;
	border:solid 2px;
}
.modalAguarde {
	background-image: url('../imagem/ajax-loader-1.gif');
}
.modalOk {
	background-image: url('../imagem/Ok_redondo_peq.jpg');
}
.modalErro {
	background-image: url('../imagem/erro_x_peq.png');
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

<form id="f" name="f" method="post" action="FarolResumidoFiltro.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="c_dt_periodo_inicio" id="c_dt_periodo_inicio" value="<%=c_dt_periodo_inicio%>">
<input type="hidden" name="c_dt_periodo_termino" id="c_dt_periodo_termino" value="<%=c_dt_periodo_termino%>">
<input type="hidden" name="c_perc_est_cresc" id="c_perc_est_cresc" value="<%=c_perc_est_cresc%>">
<input type="hidden" name="c_qtde_total_registros" id="c_qtde_total_registros" value="<%=n_reg_total%>" />


<!-- O BUTTON É UM ELEMENTO QUE POSSUI A PROPRIEDADE INNERHTML E TAMBÉM O MÉTODO CREATETEXTRANGE(), NECESSÁRIOS P/ COPIAR E COLAR O HTML COMO TEXTO FORMATADO NO EXCEL -->
<button style="display:none;" name="c_clipboard" id="c_clipboard"></button>


<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black;">

<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Farol Resumido<span class="C">&nbsp;</span></span></td>
</tr>
</table>
<br>
<br>

<div>

<!--  TEXTO EXPLICATIVO -->
<table width="540" cellpadding="4" cellspacing="4">
<tr>
	<td align="left" colspan="2"><span class="Expl">DICA<br />Para o correto funcionamento desta página, é necessário:</span></td>
</tr>
<tr>
	<td align="left" style="width:20px;">&nbsp;</td>
	<td align="left"><span class="Expl">1) NÃO copiar nada para a área de transferência (clipboard) durante o processamento, ou seja, não usar o "Copiar e Colar"!</span></td>
</tr>
<tr>
	<td align="left" style="width:20px;">&nbsp;</td>
	<td align="left"><span class="Expl">2) Este site esteja adicionado na zona de sites confiáveis.</span></td>
</tr>
<tr>
	<td align="left" style="width:20px;">&nbsp;</td>
	<td align="left"><span class="Expl">3) Nas configurações de segurança, a opção "Permitir acesso Programático à área de transferência" esteja selecionada com "Habilitar".</span></td>
</tr>
</table>
<br>

<% if n_reg_total > 0 then %>
<span id="divBotaoGeraPlanilha" class="Botao C" style='width:240px;font-size:10pt;padding-top:5px;padding-bottom:5px;'>&nbsp;&nbsp;Clique aqui para gerar a planilha&nbsp;&nbsp;</span>
<% else %>
<span id="spnMsgResultadoVazio" class="ALERTA" style='width:240px;font-size:10pt;padding-top:5px;padding-bottom:5px;'>&nbsp;&nbsp;Nenhum registro encontrado&nbsp;&nbsp;</span>
<% end if %>
<div id="divMsgAguarde" class="C modalBase modalAguarde">
	<br />
	<span id="textoMensagem" class="C" style="margin-top:32px;font-size:12pt;">Aguarde, a planilha está sendo gerada...</span>
	<div style="top:8px;left:620px;position:absolute;"><span id="btnFechar" class="Botao C" style='display:none;width:25px;font-size:8pt;padding-top:2px;padding-bottom:2px;'>&nbsp;X&nbsp;</span></div>
</div>

</div>

<!-- ************   SEPARADOR   ************ -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black;">
<tr><td align="left" class="Rc">&nbsp;</td></tr>
</table>
<br>


<table class="notPrint" width="649" cellspacing="0">
<tr>
	<td align="center"><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()" title="volta para página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a>
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