<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/Global.asp"    -->
<%
'     =================================
'	  O R C A M E N T O N O V O . A S P
'     =================================
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

	dim i, j, usuario, loja, cliente_selecionado, msg_erro
	dim idxSelecionado, blnAchou

	usuario = Trim(Session("usuario_atual"))
	loja = Trim(Session("loja_atual"))
	If (usuario = "") then Response.Redirect("Aviso.asp?id=" & ERR_SESSAO) 
	If (loja = "") then Response.Redirect("Aviso.asp?id=" & ERR_SESSAO) 
	
	
	cliente_selecionado = Trim(request("cliente_selecionado"))
	if (cliente_selecionado = "") then Response.Redirect("Aviso.asp?id=" & ERR_CLIENTE_NAO_ESPECIFICADO)

	dim cn, r, strSql
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	dim r_cliente
	set r_cliente = New cl_CLIENTE
	if Not x_cliente_bd(cliente_selecionado, r_cliente) then Response.Redirect("aviso.asp?id=" & ERR_CLIENTE_FALHA_RECUPERAR_DADOS)

	dim rb_end_entrega, EndEtg_endereco, EndEtg_endereco_numero, EndEtg_endereco_complemento, EndEtg_bairro, EndEtg_cidade, EndEtg_uf, EndEtg_cep,EndEtg_obs
	rb_end_entrega = Trim(Request.Form("rb_end_entrega"))
	EndEtg_endereco = Trim(Request.Form("EndEtg_endereco"))
	EndEtg_endereco_numero = Trim(Request.Form("EndEtg_endereco_numero"))
	EndEtg_endereco_complemento = Trim(Request.Form("EndEtg_endereco_complemento"))
	EndEtg_bairro = Trim(Request.Form("EndEtg_bairro"))
	EndEtg_cidade = Trim(Request.Form("EndEtg_cidade"))
	EndEtg_uf = Trim(Request.Form("EndEtg_uf"))
	EndEtg_cep = Trim(Request.Form("EndEtg_cep"))
    EndEtg_obs = Trim(Request.Form("EndEtg_obs"))
	dim vendedor
	vendedor = ""
	
	dim alerta
	alerta=""
	
	if Trim("" & r_cliente.cep) <> "" then
		if Len(retorna_so_digitos(Trim("" & r_cliente.cep))) < 8 then
			alerta=texto_add_br(alerta)
			alerta=alerta & "O CEP do cadastro do cliente est� incompleto (CEP: " & Trim("" & r_cliente.cep) & ")"
			end if
		end if

	if rb_end_entrega = "S" then
		if EndEtg_cep <> "" then
			if Len(retorna_so_digitos(EndEtg_cep)) < 8 then
				alerta=texto_add_br(alerta)
				alerta=alerta & "O CEP do endere�o de entrega est� incompleto (CEP: " & EndEtg_cep & ")"
				end if
			end if
		end if

	strSql = "SELECT vendedor FROM t_ORCAMENTISTA_E_INDICADOR WHERE (apelido = '" & usuario & "')"
	set r = cn.execute(strSql)
	if r.Eof then
		alerta = "FALHA AO LOCALIZAR O REGISTRO NO BANCO DE DADOS"
	else
		vendedor = Trim("" & r("vendedor"))
		end if

	if alerta = "" then
		if vendedor = "" then alerta = "N�O H� NENHUM VENDEDOR DEFINIDO PARA ATEND�-LO"
		end if
	
'	CONSIST�NCIAS P/ EMISS�O DE NFe
	dim s_tabela_municipios_IBGE
	s_tabela_municipios_IBGE = ""
	if alerta = "" then
		if rb_end_entrega = "S" then
		'	MUNIC�PIO DE ACORDO C/ TABELA DO IBGE?
			dim s_lista_sugerida_municipios
			dim v_lista_sugerida_municipios
			dim iCounterLista, iNumeracaoLista
			if Not consiste_municipio_IBGE_ok(EndEtg_cidade, EndEtg_uf, s_lista_sugerida_municipios, msg_erro) then
				if alerta <> "" then alerta = alerta & "<br><br>" & String(80,"=") & "<br><br>"
				if msg_erro <> "" then
					alerta = alerta & msg_erro
				else
					alerta = alerta & "Munic�pio '" & EndEtg_cidade & "' n�o consta na rela��o de munic�pios do IBGE para a UF de '" & EndEtg_uf & "'!!"
					if s_lista_sugerida_municipios <> "" then
						alerta = alerta & "<br>" & _
										  "Localize o munic�pio na lista abaixo e verifique se a grafia est� correta!!"
						v_lista_sugerida_municipios = Split(s_lista_sugerida_municipios, chr(13))
						iNumeracaoLista=0
						for iCounterLista=LBound(v_lista_sugerida_municipios) to UBound(v_lista_sugerida_municipios)
							if Trim("" & v_lista_sugerida_municipios(iCounterLista)) <> "" then
								iNumeracaoLista=iNumeracaoLista+1
								s_tabela_municipios_IBGE = s_tabela_municipios_IBGE & _
													"	<tr>" & chr(13) & _
													"		<td align='right'>" & chr(13) & _
													"			<span class='N'>&nbsp;" & Cstr(iNumeracaoLista) & "." & "</span>" & chr(13) & _
													"		</td>" & chr(13) & _
													"		<td align='left'>" & chr(13) & _
													"			<span class='N'>" & Trim("" & v_lista_sugerida_municipios(iCounterLista)) & "</span>" & chr(13) & _
													"		</td>" & chr(13) & _
													"	</tr>" & chr(13)
								end if
							next

						if s_tabela_municipios_IBGE <> "" then
							s_tabela_municipios_IBGE = _
									"<table cellspacing='0' cellpadding='1'>" & chr(13) & _
									"	<tr>" & chr(13) & _
									"		<td align='center'>" & chr(13) & _
									"			<p class='N'>" & "Rela��o de munic�pios de '" & EndEtg_uf & "' que se iniciam com a letra '" & Ucase(left(EndEtg_cidade,1)) & "'" & "</p>" & chr(13) & _
									"		</td>" & chr(13) & _
									"	</tr>" & chr(13) & _
									"	<tr>" & chr(13) & _
									"		<td align='center'>" & chr(13) &_
									"			<table cellspacing='0' border='1'>" & chr(13) & _
													s_tabela_municipios_IBGE & _
									"			</table>" & chr(13) & _
									"		</td>" & chr(13) & _
									"	</tr>" & chr(13) & _
									"</table>" & chr(13)
							end if
						end if
					end if
				end if 'if Not consiste_municipio_IBGE_ok()
			end if 'if rb_end_entrega = "S"
		end if 'if alerta = ""

	dim rs, rs2
	if Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	if Not cria_recordset_otimista(rs2, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

	dim s_fabricante, s_produto, s_qtde, n_qtde, s_preco_lista, s_descricao
	dim n, intIdxProduto
	dim vProduto
	redim vProduto(0)
	set vProduto(0) = New cl_ITEM_PEDIDO
	vProduto(0).qtde = 0

	if alerta = "" then
		if isLojaHabilitadaProdCompostoECommerce(loja) then
			n = Request.Form("c_produto").Count
			for i = 1 to n
				s_fabricante = Trim(Request.Form("c_fabricante")(i))
				s_fabricante = normaliza_codigo(s_fabricante, TAM_MIN_FABRICANTE)
				s_produto = Trim(Request.Form("c_produto")(i))
				s_produto = normaliza_codigo(s_produto, TAM_MIN_PRODUTO)
				s_qtde = Trim(Request.Form("c_qtde")(i))
				if IsNumeric(s_qtde) then n_qtde = CLng(s_qtde) else n_qtde = 0
			'	INFORMOU APENAS O C�DIGO DO PRODUTO NA TELA ANTERIOR
			'	TENTA RECUPERAR O C�DIGO DO FABRICANTE (VERIFICANDO SE H� AMBIGUIDADE)
				if (s_fabricante = "") And (s_produto <> "") then
				'	VERIFICA SE � PRODUTO COMPOSTO
					strSql = "SELECT " & _
								"*" & _
							" FROM t_EC_PRODUTO_COMPOSTO t_EC_PC" & _
							" WHERE" & _
								" (produto_composto = '" & s_produto & "')"
					if rs.State <> 0 then rs.Close
					rs.Open strSql, cn
				'	� PRODUTO COMPOSTO
					if Not rs.Eof then
						s_fabricante = Trim("" & rs("fabricante_composto"))
						rs.MoveNext
						if Not rs.Eof then
							alerta=texto_add_br(alerta)
							alerta=alerta & "H� mais de um produto composto com o c�digo '" & s_produto & "'!!<br />Informe o c�digo do fabricante para resolver a ambiguidade!!"
							end if
						if alerta = "" then
							strSql = "SELECT " & _
										"*" & _
									" FROM t_EC_PRODUTO_COMPOSTO_ITEM t_EC_PCI" & _
									" WHERE" & _
										" (fabricante_composto = '" & s_fabricante & "')" & _
										" AND (produto_composto = '" & s_produto & "')" & _
										" AND (excluido_status = 0)" & _
									" ORDER BY" & _
										" sequencia"
							if rs.State <> 0 then rs.Close
							rs.Open strSql, cn
							do while Not rs.Eof
								strSql = "SELECT " & _
											"*" & _
										" FROM t_PRODUTO tP" & _
											" INNER JOIN t_PRODUTO_LOJA tPL ON (tP.fabricante = tPL.fabricante) AND (tP.produto = tPL.produto)" & _
										" WHERE" & _
											" (tP.fabricante = '" & Trim("" & rs("fabricante_item")) & "')" & _
											" AND (tP.produto = '" & Trim("" & rs("produto_item")) & "')" & _
											" AND (loja = '" & loja & "')"
								if rs2.State <> 0 then rs2.Close
								rs2.Open strSql, cn
								if rs2.Eof then
									alerta=texto_add_br(alerta)
									alerta=alerta & "O produto (" & Trim("" & rs("fabricante_item")) & ")" & Trim("" & rs("produto_item")) & " n�o est� dispon�vel para a loja " & loja & "!!"
								else
									blnAchou = False
									idxSelecionado = -1
									for j=LBound(vProduto) to UBound(vProduto)
										if (Trim("" & vProduto(j).fabricante) = Trim("" & rs("fabricante_item"))) And (Trim("" & vProduto(j).produto) = Trim("" & rs("produto_item"))) then
											blnAchou = True
											idxSelecionado = j
											exit for
											end if
										next

									if Not blnAchou then
										if Trim(vProduto(ubound(vProduto)).produto) <> "" then
											redim preserve vProduto(ubound(vProduto)+1)
											set vProduto(ubound(vProduto)) = New cl_ITEM_PEDIDO
											vProduto(ubound(vProduto)).qtde = 0
											end if
										idxSelecionado = ubound(vProduto)
										end if

									with vProduto(idxSelecionado)
										.fabricante = Trim("" & rs("fabricante_item"))
										.produto = Trim("" & rs("produto_item"))
										.qtde = .qtde + (n_qtde * rs("qtde"))
										.preco_lista = rs2("preco_lista")
										.descricao = Trim("" & rs2("descricao"))
										.descricao_html = Trim("" & rs2("descricao_html"))
										end with
									end if
								rs.MoveNext
								loop
							end if
				'	� PRODUTO NORMAL
					else
						strSql = "SELECT " & _
									"*" & _
								" FROM t_PRODUTO tP" & _
									" INNER JOIN t_PRODUTO_LOJA tPL ON (tP.fabricante = tPL.fabricante) AND (tP.produto = tPL.produto)" & _
								" WHERE" & _
									" (tP.produto = '" & s_produto & "')" & _
									" AND (loja = '" & loja & "')"
						if rs.State <> 0 then rs.Close
						rs.Open strSql, cn
						if rs.Eof then
							alerta=texto_add_br(alerta)
							alerta=alerta & "Produto '" & s_produto & "' n�o foi encontrado para a loja " & loja & "!!"
						else
							blnAchou = False
							idxSelecionado = -1
							for j=LBound(vProduto) to UBound(vProduto)
								if (Trim("" & vProduto(j).fabricante) = Trim("" & rs("fabricante"))) And (Trim("" & vProduto(j).produto) = Trim("" & rs("produto"))) then
									blnAchou = True
									idxSelecionado = j
									exit for
									end if
								next

							if Not blnAchou then
								if Trim(vProduto(ubound(vProduto)).produto) <> "" then
									redim preserve vProduto(ubound(vProduto)+1)
									set vProduto(ubound(vProduto)) = New cl_ITEM_PEDIDO
									vProduto(ubound(vProduto)).qtde = 0
									end if
								idxSelecionado = ubound(vProduto)
								end if

							with vProduto(idxSelecionado)
								.fabricante = Trim("" & rs("fabricante"))
								.produto = Trim("" & rs("produto"))
								.qtde = .qtde + n_qtde
								.preco_lista = rs("preco_lista")
								.descricao = Trim("" & rs("descricao"))
								.descricao_html = Trim("" & rs("descricao_html"))
								end with
							rs.MoveNext
							if Not rs.Eof then
								alerta=texto_add_br(alerta)
								alerta=alerta & "H� mais de um produto com o c�digo '" & s_produto & "'!!<br />Informe o c�digo do fabricante para resolver a ambiguidade!!"
								end if
							end if
						end if ' Produto composto ou normal?
			'	INFORMOU O C�DIGO DO FABRICANTE E DO PRODUTO NA TELA ANTERIOR
				elseif (s_fabricante <> "") And (s_produto <> "") then
				'	VERIFICA SE � PRODUTO COMPOSTO
					strSql = "SELECT " & _
								"*" & _
							" FROM t_EC_PRODUTO_COMPOSTO t_EC_PC" & _
							" WHERE" & _
								" (fabricante_composto = '" & s_fabricante & "')" & _
								" AND (produto_composto = '" & s_produto & "')"
					if rs.State <> 0 then rs.Close
					rs.Open strSql, cn
				'	� PRODUTO COMPOSTO
					if Not rs.Eof then
						strSql = "SELECT " & _
									"*" & _
								" FROM t_EC_PRODUTO_COMPOSTO_ITEM t_EC_PCI" & _
								" WHERE" & _
									" (fabricante_composto = '" & s_fabricante & "')" & _
									" AND (produto_composto = '" & s_produto & "')" & _
									" AND (excluido_status = 0)" & _
								" ORDER BY" & _
									" sequencia"
						if rs.State <> 0 then rs.Close
						rs.Open strSql, cn
						do while Not rs.Eof
							strSql = "SELECT " & _
										"*" & _
									" FROM t_PRODUTO tP" & _
										" INNER JOIN t_PRODUTO_LOJA tPL ON (tP.fabricante = tPL.fabricante) AND (tP.produto = tPL.produto)" & _
									" WHERE" & _
										" (tP.fabricante = '" & Trim("" & rs("fabricante_item")) & "')" & _
										" AND (tP.produto = '" & Trim("" & rs("produto_item")) & "')" & _
										" AND (loja = '" & loja & "')"
							if rs2.State <> 0 then rs2.Close
							rs2.Open strSql, cn
							if rs2.Eof then
								alerta=texto_add_br(alerta)
								alerta=alerta & "O produto (" & Trim("" & rs("fabricante_item")) & ")" & Trim("" & rs("produto_item")) & " n�o est� dispon�vel para a loja " & loja & "!!"
							else
								blnAchou = False
								idxSelecionado = -1
								for j=LBound(vProduto) to UBound(vProduto)
									if (Trim("" & vProduto(j).fabricante) = Trim("" & rs("fabricante_item"))) And (Trim("" & vProduto(j).produto) = Trim("" & rs("produto_item"))) then
										blnAchou = True
										idxSelecionado = j
										exit for
										end if
									next

								if Not blnAchou then
									if Trim(vProduto(ubound(vProduto)).produto) <> "" then
										redim preserve vProduto(ubound(vProduto)+1)
										set vProduto(ubound(vProduto)) = New cl_ITEM_PEDIDO
										vProduto(ubound(vProduto)).qtde = 0
										end if
									idxSelecionado = ubound(vProduto)
									end if

								with vProduto(idxSelecionado)
									.fabricante = Trim("" & rs("fabricante_item"))
									.produto = Trim("" & rs("produto_item"))
									.qtde = .qtde + (n_qtde * rs("qtde"))
									.preco_lista = rs2("preco_lista")
									.descricao = Trim("" & rs2("descricao"))
									.descricao_html = Trim("" & rs2("descricao_html"))
									end with
								end if
							rs.MoveNext
							loop
				'	� PRODUTO NORMAL
					else
						strSql = "SELECT " & _
									"*" & _
								" FROM t_PRODUTO tP" & _
									" INNER JOIN t_PRODUTO_LOJA tPL ON (tP.fabricante = tPL.fabricante) AND (tP.produto = tPL.produto)" & _
								" WHERE" & _
									" (tP.fabricante = '" & s_fabricante & "')" & _
									" AND (tP.produto = '" & s_produto & "')" & _
									" AND (loja = '" & loja & "')"
						if rs.State <> 0 then rs.Close
						rs.Open strSql, cn
						if rs.Eof then
							alerta=texto_add_br(alerta)
							alerta=alerta & "Produto (" & s_fabricante & ")" & s_produto & " n�o foi encontrado para a loja " & loja & "!!"
						else
							blnAchou = False
							idxSelecionado = -1
							for j=LBound(vProduto) to UBound(vProduto)
								if (Trim("" & vProduto(j).fabricante) = Trim("" & rs("fabricante"))) And (Trim("" & vProduto(j).produto) = Trim("" & rs("produto"))) then
									blnAchou = True
									idxSelecionado = j
									exit for
									end if
								next

							if Not blnAchou then
								if Trim(vProduto(ubound(vProduto)).produto) <> "" then
									redim preserve vProduto(ubound(vProduto)+1)
									set vProduto(ubound(vProduto)) = New cl_ITEM_PEDIDO
									vProduto(ubound(vProduto)).qtde = 0
									end if
								idxSelecionado = ubound(vProduto)
								end if

							with vProduto(idxSelecionado)
								.fabricante = Trim("" & rs("fabricante"))
								.produto = Trim("" & rs("produto"))
								.qtde = .qtde + n_qtde
								.preco_lista = rs("preco_lista")
								.descricao = Trim("" & rs("descricao"))
								.descricao_html = Trim("" & rs("descricao_html"))
								end with
							end if
						end if ' Produto composto ou normal?
					end if
				next
			
			if alerta = "" then
				n = 0
				for i=LBound(vProduto) to UBound(vProduto)
					if Trim(vProduto(i).produto) <> "" then n = n + 1
					next
				if n > MAX_ITENS then
					alerta=texto_add_br(alerta)
					alerta=alerta & "O n�mero de itens que est� sendo cadastrado (" & CStr(n) & ") excede o m�ximo permitido por pedido (" & CStr(MAX_ITENS) & ")!!"
					end if
				end if
			end if 'if isLojaHabilitadaProdCompostoECommerce(loja)
		end if 'if alerta = ""
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
	<title><%=TITULO_JANELA_MODULO_ORCAMENTO%></title>
	</head>



<script src="<%=URL_FILE__JQUERY%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__AJAX_JS%>" language="JavaScript" type="text/javascript"></script>

<script type="text/javascript">
	$(function() {
	    $("#divAjaxRunning").css('filter', 'alpha(opacity=60)'); // TRANSPAR�NCIA NO IE8
	    <%if isLojaHabilitadaProdCompostoECommerce(loja) then%>
	    fORC.submit();
        <%end if%>

	});

	//Every resize of window
	$(window).resize(function() {
		sizeDivAjaxRunning();
	});

	//Every scroll of window
	$(window).scroll(function() {
		sizeDivAjaxRunning();
	});

	//Dynamically assign height
	function sizeDivAjaxRunning() {
		var newTop = $(window).scrollTop() + "px";
		$("#divAjaxRunning").css("top", newTop);
	}
</script>

<script language="JavaScript" type="text/javascript">
var fCustoFinancFornecParcelamentoPopup;
var objAjaxCustoFinancFornecConsultaPreco;

function processaSelecaoCustoFinancFornecParcelamento(){};

function abreTabelaCustoFinancFornecParcelamento(intIndex){
var f, strUrl;
	f=fORC;
	if (trim(f.c_fabricante[intIndex].value)=="") {
		alert("Informe o c�digo do fabricante do produto!");
		f.c_fabricante[intIndex].focus();
		return;
		}
	if (trim(f.c_produto[intIndex].value)=="") {
		alert("Informe o c�digo do produto!");
		f.c_produto[intIndex].focus();
		return;
		}
		
	try
		{
	//  SE J� HOUVER UMA JANELA DE TABELA DE PARCELAMENTO ABERTA, GARANTE QUE ELA SER� FECHADA
	//  E UMA NOVA SER� CRIADA (EVITA PROBLEMAS C/ O 'WINDOW.OPENER')
		fCustoFinancFornecParcelamentoPopup.close();
		}
	catch (e) {
	 // NOP
		}
	processaSelecaoCustoFinancFornecParcelamento=trataSelecaoCustoFinancFornecParcelamento;
	strUrl="../Global/AjaxCustoFinancFornecParcelamentoPopup.asp";
	strUrl=strUrl+"?fabricante="+trim(f.c_fabricante[intIndex].value)+"&produto="+trim(f.c_produto[intIndex].value)+"&loja="+trim(f.c_loja.value)+"&tipoParcelamento="+f.c_custoFinancFornecTipoParcelamento.value+"&qtdeParcelas="+f.c_custoFinancFornecQtdeParcelas.value;
	try
	{
		fCustoFinancFornecParcelamentoPopup=window.open(strUrl, "AjaxCustoFinancFornecParcelamentoPopup", "status=1,toolbar=0,location=0,menubar=0,directories=0,resizable=1,scrollbars=1,width=800,height=675,left=0,top=0");
	}
	catch (e) {
		alert("Falha ao ativar o painel com a tabela de pre�os!!\n"+e.message);
		}
	
	try
	{
		fCustoFinancFornecParcelamentoPopup.focus();
	}
	catch (e) {
	 // NOP
		}
}

function trataSelecaoCustoFinancFornecParcelamento(strTipoParcelamento, strPrecoLista, intQtdeParcelas, strFabricante, strProduto) {
var f,i,blnAlterou;
	f=fORC;
//  Percorre o la�o at� o final para o caso do usu�rio ter digitado o mesmo produto em v�rias linhas
//	(apesar de que isso n�o ser� aceito pelas consist�ncias que ser�o feitas).
	for (i=0; i<f.c_fabricante.length; i++) {
		if ((f.c_fabricante[i].value==strFabricante)&&(f.c_produto[i].value==strProduto)) {
			f.c_preco_lista[i].value=strPrecoLista;
			f.c_preco_lista[i].style.color="black";
			}
		}
	blnAlterou=false;
	if (f.c_custoFinancFornecTipoParcelamento.value!=strTipoParcelamento) blnAlterou=true;
	if (!blnAlterou) {
		if ((strTipoParcelamento==COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__COM_ENTRADA)||(strTipoParcelamento==COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__SEM_ENTRADA)) {
			if (converte_numero(intQtdeParcelas)!=converte_numero(f.c_custoFinancFornecQtdeParcelas.value)) blnAlterou=true;
			}
		}
//  Memoriza sele��o atual
	f.c_custoFinancFornecTipoParcelamento.value=strTipoParcelamento;
	f.c_custoFinancFornecQtdeParcelas.value=intQtdeParcelas;
	
	if (blnAlterou) {
		f.c_custoFinancFornecParcelamentoDescricao.value=descricaoCustoFinancFornecTipoParcelamento(strTipoParcelamento);
		if (strTipoParcelamento==COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__COM_ENTRADA) {
			f.c_custoFinancFornecParcelamentoDescricao.value += " (1+" + intQtdeParcelas + ")";
			}
		else if (strTipoParcelamento==COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__SEM_ENTRADA) {
			f.c_custoFinancFornecParcelamentoDescricao.value += " (0+" + intQtdeParcelas + ")";
			}
	
		// Houve altera��o no tipo de parcelamento, portanto, � necess�rio atualizar os 
		// pre�os de lista de todos os produtos
		atualizaPrecos(strFabricante, strProduto);
		}
	window.status="Conclu�do";
}

function trataRespostaAjaxCustoFinancFornecSincronizaPrecos() {
var f, strResp, i, j, xmlDoc, oNodes;
var strFabricante,strProduto, strStatus, strPrecoLista, strDescricao, strMsgErro;
	f=fORC;
	if (objAjaxCustoFinancFornecConsultaPreco.readyState==AJAX_REQUEST_IS_COMPLETE) {
		strResp=objAjaxCustoFinancFornecConsultaPreco.responseText;
		if (strResp=="") {
			alert("Falha ao consultar o pre�o!!");
			window.status="Conclu�do";
			$("#divAjaxRunning").hide();
			return;
			}
		
		if (strResp!="") {
			try
				{
				xmlDoc=objAjaxCustoFinancFornecConsultaPreco.responseXML.documentElement;
				for (i=0; i < xmlDoc.getElementsByTagName("ItemConsulta").length; i++) {
				//  Fabricante
					oNodes=xmlDoc.getElementsByTagName("fabricante")[i];
					if (oNodes.childNodes.length > 0) strFabricante=oNodes.childNodes[0].nodeValue; else strFabricante="";
					if (strFabricante==null) strFabricante="";
				//  Produto
					oNodes=xmlDoc.getElementsByTagName("produto")[i];
					if (oNodes.childNodes.length > 0) strProduto=oNodes.childNodes[0].nodeValue; else strProduto="";
					if (strProduto==null) strProduto="";
				//  Status
					oNodes=xmlDoc.getElementsByTagName("status")[i];
					if (oNodes.childNodes.length > 0) strStatus=oNodes.childNodes[0].nodeValue; else strStatus="";
					if (strStatus==null) strStatus="";
					if (strStatus=="OK") {
					//  Descri��o
						oNodes=xmlDoc.getElementsByTagName("descricao")[i];
						if (oNodes.childNodes.length > 0) strDescricao=oNodes.childNodes[0].nodeValue; else strDescricao="";
						if (strDescricao==null) strDescricao="";
					//  Pre�o
						oNodes=xmlDoc.getElementsByTagName("precoLista")[i];
						if (oNodes.childNodes.length > 0) strPrecoLista=oNodes.childNodes[0].nodeValue; else strPrecoLista="";
						if (strPrecoLista==null) strPrecoLista="";
					//  Atualiza o pre�o
						if (strPrecoLista=="") {
							alert("Falha na consulta do pre�o do produto " + strProduto + "!!\n" + strMsgErro);
							}
						else {
							for (j=0; j<f.c_fabricante.length; j++) {
								if ((f.c_fabricante[j].value==strFabricante)&&(f.c_produto[j].value==strProduto)) {
								//  Percorre o la�o at� o final para o caso do usu�rio ter digitado o mesmo produto em v�rias linhas
								//	(apesar de que isso n�o ser� aceito pelas consist�ncias que ser�o feitas).
									f.c_preco_lista[j].value=strPrecoLista;
									f.c_descricao[j].value=strDescricao;
									f.c_preco_lista[j].style.color="black";
									}
								}
							}
						}
					else {
					//  Mensagem de Erro
						oNodes=xmlDoc.getElementsByTagName("msg_erro")[i];
						if (oNodes.childNodes.length > 0) strMsgErro=oNodes.childNodes[0].nodeValue; else strMsgErro="";
						if (strMsgErro==null) strMsgErro="";
						for (j=0; j<f.c_fabricante.length; j++) {
						//  Percorre o la�o at� o final para o caso do usu�rio ter digitado o mesmo produto em v�rias linhas
						//	(apesar de que isso n�o ser� aceito pelas consist�ncias que ser�o feitas).
							if ((f.c_fabricante[j].value==strFabricante)&&(f.c_produto[j].value==strProduto)) {
								f.c_preco_lista[j].style.color=COR_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__INEXISTENTE;
								}
							}
						alert("Falha ao consultar o pre�o do produto " + strProduto + "!!\n" + strMsgErro);
						}
					}
				}
			catch (e)
				{
				alert("Falha na consulta do pre�o!!\n"+e.message);
				}
			}
		window.status="Conclu�do";
		$("#divAjaxRunning").hide();
		}
}

function atualizaPrecos(strFabricanteSelecionado, strProdutoSelecionado) {
var f, i, strListaProdutos, strUrl;
	f=fORC;
	objAjaxCustoFinancFornecConsultaPreco=GetXmlHttpObject();
	if (objAjaxCustoFinancFornecConsultaPreco==null) {
		alert("O browser N�O possui suporte ao AJAX!!");
		return;
		}
		
	strListaProdutos="";
	for (i=0; i<f.c_fabricante.length; i++) {
		if ((trim(f.c_fabricante[i].value)!="")&&(trim(f.c_produto[i].value)!="")) {
		//  N�o atualiza o pre�o do produto que acabou de ser consultado atrav�s da tabela de pre�os.
		//  Atualiza somente os demais produtos, se houver.
			if ((strFabricanteSelecionado!=trim(f.c_fabricante[i].value))||(strProdutoSelecionado!=trim(f.c_produto[i].value))) {
				if (strListaProdutos!="") strListaProdutos+=";";
				strListaProdutos += f.c_fabricante[i].value + "|" + f.c_produto[i].value;
				}
			}
		}
	if (strListaProdutos=="") return;
	
	window.status="Aguarde, consultando pre�os ...";
	$("#divAjaxRunning").show();
		
	strUrl = "../Global/AjaxCustoFinancFornecConsultaPrecoBD.asp";
	strUrl+="?tipoParcelamento="+f.c_custoFinancFornecTipoParcelamento.value;
	strUrl+="&qtdeParcelas="+f.c_custoFinancFornecQtdeParcelas.value;
	strUrl+="&loja="+f.c_loja.value;
	strUrl+="&listaProdutos="+strListaProdutos;
//  Prevents server from using a cached file
	strUrl=strUrl+"&sid="+Math.random()+Math.random();
	objAjaxCustoFinancFornecConsultaPreco.onreadystatechange=trataRespostaAjaxCustoFinancFornecSincronizaPrecos;
	objAjaxCustoFinancFornecConsultaPreco.open("GET",strUrl,true);
	objAjaxCustoFinancFornecConsultaPreco.send(null);
}

function trataRespostaAjaxCustoFinancFornecConsultaPreco() {
var f, strResp, i, j, xmlDoc, oNodes;
var strFabricante,strProduto, strStatus, strPrecoLista, strDescricao, strMsgErro;
	f=fORC;
	if (objAjaxCustoFinancFornecConsultaPreco.readyState==AJAX_REQUEST_IS_COMPLETE) {
		strResp=objAjaxCustoFinancFornecConsultaPreco.responseText;
		if (strResp=="") {
			alert("Falha ao consultar o pre�o!!");
			window.status="Conclu�do";
			$("#divAjaxRunning").hide();
			return;
			}
		
		if (strResp!="") {
			try
				{
				xmlDoc=objAjaxCustoFinancFornecConsultaPreco.responseXML.documentElement;
				for (i=0; i < xmlDoc.getElementsByTagName("ItemConsulta").length; i++) {
				//  Fabricante
					oNodes=xmlDoc.getElementsByTagName("fabricante")[i];
					if (oNodes.childNodes.length > 0) strFabricante=oNodes.childNodes[0].nodeValue; else strFabricante="";
					if (strFabricante==null) strFabricante="";
				//  Produto
					oNodes=xmlDoc.getElementsByTagName("produto")[i];
					if (oNodes.childNodes.length > 0) strProduto=oNodes.childNodes[0].nodeValue; else strProduto="";
					if (strProduto==null) strProduto="";
				//  Descri��o
					oNodes=xmlDoc.getElementsByTagName("descricao")[i];
					if (oNodes.childNodes.length > 0) strDescricao=oNodes.childNodes[0].nodeValue; else strDescricao="";
					if (strDescricao==null) strDescricao="";
					if (strDescricao!="") {
						for (j=0; j<f.c_fabricante.length; j++) {
							if ((f.c_fabricante[j].value==strFabricante)&&(f.c_produto[j].value==strProduto)) {
							//  Percorre o la�o at� o final para o caso do usu�rio ter digitado o mesmo produto em v�rias linhas
							//	(apesar de que isso n�o ser� aceito pelas consist�ncias que ser�o feitas).
								f.c_descricao[j].value=strDescricao;
								}
							}
						}
				//  Status
					oNodes=xmlDoc.getElementsByTagName("status")[i];
					if (oNodes.childNodes.length > 0) strStatus=oNodes.childNodes[0].nodeValue; else strStatus="";
					if (strStatus==null) strStatus="";
					if (strStatus=="OK") {
					//  Pre�o
						oNodes=xmlDoc.getElementsByTagName("precoLista")[i];
						if (oNodes.childNodes.length > 0) strPrecoLista=oNodes.childNodes[0].nodeValue; else strPrecoLista="";
						if (strPrecoLista==null) strPrecoLista="";
					//  Atualiza o pre�o
						if (strPrecoLista=="") {
							alert("Falha na consulta do pre�o do produto " + strProduto + "\n" + strMsgErro);
							}
						else {
							for (j=0; j<f.c_fabricante.length; j++) {
								if ((f.c_fabricante[j].value==strFabricante)&&(f.c_produto[j].value==strProduto)) {
								//  Percorre o la�o at� o final para o caso do usu�rio ter digitado o mesmo produto em v�rias linhas
								//	(apesar de que isso n�o ser� aceito pelas consist�ncias que ser�o feitas).
									f.c_preco_lista[j].value=strPrecoLista;
									f.c_preco_lista[j].style.color="black";
									}
								}
							}
						}
					else {
					//  Mensagem de Erro
						oNodes=xmlDoc.getElementsByTagName("msg_erro")[i];
						if (oNodes.childNodes.length > 0) strMsgErro=oNodes.childNodes[0].nodeValue; else strMsgErro="";
						if (strMsgErro==null) strMsgErro="";
						for (j=0; j<f.c_fabricante.length; j++) {
						//  Percorre o la�o at� o final para o caso do usu�rio ter digitado o mesmo produto em v�rias linhas
						//	(apesar de que isso n�o ser� aceito pelas consist�ncias que ser�o feitas).
							if ((f.c_fabricante[j].value==strFabricante)&&(f.c_produto[j].value==strProduto)) {
								f.c_preco_lista[j].style.color=COR_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__INEXISTENTE;
								}
							}
						alert("Falha ao consultar o pre�o do produto " + strProduto + "\n" + strMsgErro);
						}
					}
				}
			catch (e)
				{
				alert("Falha na consulta do pre�o!!\n"+e.message);
				}
			}
		window.status="Conclu�do";
		$("#divAjaxRunning").hide();
		}
}

function consultaPreco(intIndice) {
var f, i, strProdutoSelecionado, strUrl;
	f=fORC;
	if (trim(f.c_fabricante[intIndice].value)=="") return;
	if (trim(f.c_produto[intIndice].value)=="") return;
	
	objAjaxCustoFinancFornecConsultaPreco=GetXmlHttpObject();
	if (objAjaxCustoFinancFornecConsultaPreco==null) {
		alert("O browser N�O possui suporte ao AJAX!!");
		return;
		}
		
	strProdutoSelecionado=f.c_fabricante[intIndice].value + "|" + f.c_produto[intIndice].value;
	
	window.status="Aguarde, consultando pre�o ...";
	$("#divAjaxRunning").show();
		
	strUrl = "../Global/AjaxCustoFinancFornecConsultaPrecoBD.asp";
	strUrl+="?tipoParcelamento="+f.c_custoFinancFornecTipoParcelamento.value;
	strUrl+="&qtdeParcelas="+f.c_custoFinancFornecQtdeParcelas.value;
	strUrl+="&loja="+f.c_loja.value;
	strUrl+="&listaProdutos="+strProdutoSelecionado;
//  Prevents server from using a cached file
	strUrl=strUrl+"&sid="+Math.random()+Math.random();
	objAjaxCustoFinancFornecConsultaPreco.onreadystatechange=trataRespostaAjaxCustoFinancFornecConsultaPreco;
	objAjaxCustoFinancFornecConsultaPreco.open("GET",strUrl,true);
	objAjaxCustoFinancFornecConsultaPreco.send(null);
}

function trataLimpaLinha(intIndice) {
var f;
	f=fORC;
	if ((trim(f.c_fabricante[intIndice].value)=="")&&(trim(f.c_produto[intIndice].value)=="")) {
		f.c_qtde[intIndice].value="";
		f.c_descricao[intIndice].value="";
		f.c_preco_lista[intIndice].value="";
		}
}

function fORCConfirma( f ) {
var i, b, ha_item, strMsgErro;
	ha_item=false;
	for (i=0; i < f.c_produto.length; i++) {
		b=false;
		if (trim(f.c_fabricante[i].value)!="") b=true;
		if (trim(f.c_produto[i].value)!="") b=true;
		if (trim(f.c_qtde[i].value)!="") b=true;
		
		if (b) {
			ha_item=true;
			if (trim(f.c_fabricante[i].value)=="") {
				alert("Informe o c�digo do fabricante!!");
				f.c_fabricante[i].focus();
				return;
				}
			if (trim(f.c_produto[i].value)=="") {
				alert("Informe o c�digo do produto!!");
				f.c_produto[i].focus();
				return;
				}
			if (trim(f.c_qtde[i].value)=="") {
				alert("Informe a quantidade!!");
				f.c_qtde[i].focus();
				return;
				}
			if (parseInt(f.c_qtde[i].value)<=0) {
				alert("Quantidade inv�lida!!");
				f.c_qtde[i].focus();
				return;
				}
			}
		}
	
	if (!ha_item) {
		alert("N�o h� produtos na lista!!");
		f.c_fabricante[0].focus();
		return;
		}

/*	if (trim(f.midia.value)=='') {
		alert('Indique a forma pela qual o cliente conheceu a Bonshop!!');
		f.midia.focus();
		return;
		}
*/
		
	
	if (trim(f.vendedor.value)=='') {
		alert("Indique um vendedor!!");
		f.vendedor.focus();
		return;
		}

	if (trim(f.c_custoFinancFornecTipoParcelamento.value)=="") {
		alert('N�o foi informada a forma de pagamento!');
		return;
		}
	
	if ((f.c_custoFinancFornecTipoParcelamento.value==COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__COM_ENTRADA)||
		(f.c_custoFinancFornecTipoParcelamento.value==COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__SEM_ENTRADA)) {
		if (converte_numero(f.c_custoFinancFornecQtdeParcelas.value)==0) {
			alert('N�o foi informada a quantidade de parcelas da forma de pagamento!');
			return;
			}
		}
	
	strMsgErro="";
	for (i=0; i < f.c_produto.length; i++) {
		if (trim(f.c_produto[i].value)!="") {
			if (f.c_preco_lista[i].style.color.toLowerCase()==COR_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__INEXISTENTE.toLowerCase()) {
				strMsgErro+="\n" + f.c_produto[i].value + " - " + f.c_descricao[i].value;
				}
			}
		}
	if (strMsgErro!="") {
		strMsgErro="A forma de pagamento " + KEY_ASPAS + f.c_custoFinancFornecParcelamentoDescricao.value.toLowerCase() + KEY_ASPAS + " n�o est� dispon�vel para o(s) produto(s):"+strMsgErro;
		alert(strMsgErro);
		return;
		}
		
	dCONFIRMA.style.visibility="hidden";
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

<link href="<%=URL_FILE__E_CSS%>" rel="stylesheet" type="text/css">
<link href="<%=URL_FILE__EPRINTER_CSS%>" rel="stylesheet" type="text/css" media="print">

<style type="text/css">
#divAjaxRunning
{
	position:absolute;
	top:0;
	left:0;
	width:100%;
	height:100%;
	z-index:1001;
	background-color:grey;
	opacity: .6;
}
.AjaxImgLoader
{
	position: absolute;
	left: 50%;
	top: 50%;
	margin-left: -128px; /* -1 * image width / 2 */
	margin-top: -128px;  /* -1 * image height / 2 */
	display: block;
}
</style>



<% if alerta <> "" then %>
<!-- ************************************************************ -->
<!-- **********  P�GINA PARA EXIBIR MENSAGENS DE ERRO  ********** -->
<!-- ************************************************************ -->
<body>
<center>
<br>
<!--  T E L A  -->
<p class="T">A V I S O</p>
<div class="MtAlerta" style="width:600px;font-weight:bold;" align="center"><p style='margin:5px 2px 5px 2px;'><%=alerta%></p></div>
<% if s_tabela_municipios_IBGE <> "" then %>
	<br /><br />
	<%=s_tabela_municipios_IBGE%>
<% end if %>
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
<body onload="if (trim(fORC.c_fabricante[0].value)=='') fORC.c_fabricante[0].focus();">
<center>

<form id="fORC" name="fORC" method="post" action="OrcamentoNovoConsiste.asp">
<input type="hidden" name="c_loja" id="c_loja" value='<%=loja%>'>
<input type="hidden" name="cliente_selecionado" id="cliente_selecionado" value='<%=cliente_selecionado%>'>
<input type="hidden" name="vendedor" id="vendedor" value='<%=vendedor%>'>
<input type="hidden" name="rb_end_entrega" id="rb_end_entrega" value='<%=rb_end_entrega%>'>
<input type="hidden" name="EndEtg_endereco" id="EndEtg_endereco" value="<%=EndEtg_endereco%>">
<input type="hidden" name="EndEtg_endereco_numero" id="EndEtg_endereco_numero" value="<%=EndEtg_endereco_numero%>">
<input type="hidden" name="EndEtg_endereco_complemento" id="EndEtg_endereco_complemento" value="<%=EndEtg_endereco_complemento%>">
<input type="hidden" name="EndEtg_bairro" id="EndEtg_bairro" value="<%=EndEtg_bairro%>">
<input type="hidden" name="EndEtg_cidade" id="EndEtg_cidade" value="<%=EndEtg_cidade%>">
<input type="hidden" name="EndEtg_uf" id="EndEtg_uf" value="<%=EndEtg_uf%>">
<input type="hidden" name="EndEtg_cep" id="EndEtg_cep" value="<%=EndEtg_cep%>">
<input type="hidden" name="EndEtg_obs" id="EndEtg_obs" value='<%=EndEtg_obs%>'>
<input type="hidden" name="c_custoFinancFornecTipoParcelamento" id="c_custoFinancFornecTipoParcelamento" value='<%=COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__A_VISTA%>'>
<input type="hidden" name="c_custoFinancFornecQtdeParcelas" id="c_custoFinancFornecQtdeParcelas" value='0'>

<!-- AJAX EM ANDAMENTO -->
<div id="divAjaxRunning" style="display:none;"><img src="../Imagem/ajax_loader_gray_256.gif" class="AjaxImgLoader"/></div>


<!--  I D E N T I F I C A � � O   D O   O R � A M E N T O -->
<table width="749" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">

<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Pr�-Pedido Novo</span></td>
</tr>
</table>
<br>

<!--  R E L A � � O   D E   P R O D U T O S  -->
<table class="Qx" cellspacing="0" <%if isLojaHabilitadaProdCompostoECommerce(loja) then Response.Write "style='display:none;'"%>>
	<tr bgcolor="#FFFFFF">
	<td class="MB" align="left"><span class="PLTe">Fabr</span></td>
	<td class="MB" align="left"><span class="PLTe">Produto</span></td>
	<td class="MB" align="right"><span class="PLTd">Qtde</span></td>
	<td class="MB" align="left"><span class="PLTe">Descri��o</span></td>
	<td class="MB" align="right"><span class="PLTd">VL Unit</span></td>
	</tr>
<% intIdxProduto = LBound(vProduto)-1 %>
<% for i=1 to MAX_ITENS
		intIdxProduto = intIdxProduto + 1
		s_fabricante = ""
		s_produto = ""
		s_qtde = ""
		s_preco_lista = ""
		s_descricao = ""
		if isLojaHabilitadaProdCompostoECommerce(loja) then
			if intIdxProduto <= Ubound(vProduto) then
				if Trim("" & vProduto(intIdxProduto).produto) <> "" then
					with vProduto(intIdxProduto)
						s_fabricante = .fabricante
						s_produto = .produto
						s_qtde = CStr(.qtde)
						s_preco_lista = formata_moeda(.preco_lista)
						s_descricao = .descricao
						end with
					end if
				end if
			end if
%>
	<tr>
	<td class="MDBE" align="left">
		<input name="c_fabricante" id="c_fabricante" class="PLLe" maxlength="4" style="width:30px;" onkeypress="if (digitou_enter(true)&&(tem_info(this.value)||(<%=Cstr(i)%>!=1))) if (trim(this.value)=='') fORC.midia.focus(); else fORC.c_produto[<%=Cstr(i-1)%>].focus(); filtra_fabricante();" onblur="this.value=normaliza_codigo(this.value,TAM_MIN_FABRICANTE);trataLimpaLinha(<%=Cstr(i-1)%>);"
			<% 'A DECLARA��O DA PROPRIEDADE VALUE APENAS SE HOUVER VALOR EVITA QUE O CAMPO SEJA LIMPO AP�S A SEGUNDA CHAMADA DO HISTORY.BACK() NA TELA SEGUINTE QUANDO OCORRE ERRO DE CONSIST�NCIA E DESEJA-SE RETORNAR � ESTA TELA %>
			<% if s_fabricante <> "" then %>
			value="<%=s_fabricante%>"
			<% end if %>
			/>
	</td>
	<td class="MDB" align="left">
		<input name="c_produto" id="c_produto" class="PLLe" maxlength="8" style="width:60px;" onkeypress="if (digitou_enter(true)) fORC.c_qtde[<%=Cstr(i-1)%>].focus(); filtra_produto();" onblur="this.value=normaliza_produto(this.value);consultaPreco(<%=Cstr(i-1)%>);trataLimpaLinha(<%=Cstr(i-1)%>);"
			<% if s_produto <> "" then %>
			value="<%=s_produto%>"
			<% end if %>
			/>
	</td>
	<td class="MDB" align="right">
		<input name="c_qtde" id="c_qtde" class="PLLd" maxlength="4" style="width:30px;" onkeypress="if (digitou_enter(true)) {if (<%=Cstr(i)%>==fORC.c_qtde.length) fORC.midia.focus(); else fORC.c_fabricante[<%=Cstr(i)%>].focus();} filtra_numerico();"
			<% if s_qtde <> "" then %>
			value="<%=s_qtde%>"
			<% end if %>
			/>
	</td>
	<td class="MDB" align="left">
		<input name="c_descricao" id="c_descricao" class="PLLe" style="width:377px;" readonly tabindex=-1
			<% if s_descricao <> "" then %>
			value="<%=s_descricao%>"
			<% end if %>
			/>
	</td>
	<td class="MDB" align="right">
		<input name="c_preco_lista" id="c_preco_lista" class="PLLd" style="width:62px;" readonly tabindex=-1
			<% if s_preco_lista <> "" then %>
			value="<%=s_preco_lista%>"
			<% end if %>
			/>
	</td>
	<td align="left">
		&nbsp;<button type="button" name="bCustoFinancFornecParcelamento" id="bCustoFinancFornecParcelamento" style='width:50px;font-size:8pt;font-weight:bold;color:black;margin-bottom:1px;' class="Botao" onclick="abreTabelaCustoFinancFornecParcelamento(<%=i-1%>);"><%=SIMBOLO_MONETARIO%></button>
	</td>
	</tr>
<% next %>
</table>

<br>
<div <%if isLojaHabilitadaProdCompostoECommerce(loja) then Response.Write "style='display:none;'"%>>
    <span class="PLLe">Forma de pagamento: </span>
    <input name="c_custoFinancFornecParcelamentoDescricao" id="c_custoFinancFornecParcelamentoDescricao" class="PLLe" style="width:115px;color:#0000CD;font-weight:bold;"
	    value="<%=descricaoCustoFinancFornecTipoParcelamento(COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__A_VISTA)%>">
</div>
<br>

<!-- ************   M�DIA (INATIVO)  ************ -->
<!-- <table cellspacing="0" style="width:375px" style="display:none">
	<tr>
	<td width="100%" align="left"><p class="R">FORMA PELA QUAL CONHECEU A BONSHOP</p><p class="C">
		<select id="midia" name="midia" style="margin-top:4pt; margin-bottom:4pt;width:370px;">
			<%'=midia_monta_itens_select(r_cliente.midia)%>
		</select>
		</p></td>
	</tr>
</table>
//-->


<!-- ************   SEPARADOR   ************ -->
<table width="749" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black;display:none">
<tr><td class="Rc" align="left">&nbsp;</td></tr>
</table>
<br>


<table class="notPrint" width="749" cellspacing="0" <%if isLojaHabilitadaProdCompostoECommerce(loja) then Response.Write "style='display:none;'"%>>
<tr>
	<% if isLojaHabilitadaProdCompostoECommerce(loja) then %>
	<td align="left"><a name="bCANCELA" id="bCANCELA" href="javascript:history.back();" title="volta para a p�gina anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
	<% else %>
	<td align="left"><a name="bCANCELA" id="bCANCELA" href="Resumo.asp" title="cancela o novo pr�-pedido">
		<img src="../botao/cancelar.gif" width="176" height="55" border="0"></a></td>
	<% end if %>
	<td align="right"><div name="dCONFIRMA" id="dCONFIRMA">
		<a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fORCConfirma(fORC)" title="vai para a p�gina de confirma��o">
		<img src="../botao/proximo.gif" width="176" height="55" border="0"></a></div>
	</td>
</tr>
</table>
</form>
    <%if isLojaHabilitadaProdCompostoECommerce(loja) then%>
<!-- Aguarde //-->
<table width="749" class="notPrint">
    <tr>
        <td style="text-align:right;vertical-align:middle;width:50%;">
            <img src="../IMAGEM/aguarde.gif" />
        </td>
        <td style="text-align:left;vertical-align:middle;width:50%">
            <span class="C">Redirecionando...</span>
        </td>
    </tr>
</table>
    <% end if %>
</center>
</body>

<% end if %>

</html>


<%
	if rs.State <> 0 then rs.Close
	set rs = nothing

	if rs2.State <> 0 then rs2.Close
	set rs2 = nothing

'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>