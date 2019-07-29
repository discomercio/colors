<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<% Response.Expires=-1 %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<%
'     =========================================
'	  AjaxConsultaDadosProdutoBD.asp
'     =========================================
'
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


' _____________________________________________________________________________________________
'
'			I N I C I A L I Z A     P Á G I N A     A S P    N O    S E R V I D O R
' _____________________________________________________________________________________________

	On Error GoTo 0
	Err.Clear

	class cl_RespConsultaDadosProdutoBD
		dim fabricante
		dim produto
		dim status
		dim precoLista
		dim descricao
		dim descricao_html
		dim tabela_origem
		dim codigo_erro
		dim msg_erro
		end class
	
'	OBTEM O ID
	dim strSql, strResp, msg_erro
	
'	CONECTA COM O BANCO DE DADOS
	dim cn, rs, rs2
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	If Not cria_recordset_otimista(rs2, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

	dim strLoja, strListaProdutos, vProdutos, vAux, intCounter, blnFabricanteInformado, blnProdutoCompostoProcessadoOk, vl_prod_composto_preco_lista_loja, blnProdCompostoPrecoListaItensOk
	strLoja = Trim(Request("loja"))
	strListaProdutos = Trim(Request("ListaProdutos"))
	
	if strListaProdutos = "" then
		Response.End
	elseif converte_numero(strLoja) = 0 then
		Response.End
		end if

	dim vResp
	redim vResp(0)
	set vResp(0) = new cl_RespConsultaDadosProdutoBD
	vResp(0).produto = ""
	
	vProdutos = Split(strListaProdutos, ";")
	for intCounter=Lbound(vProdutos) to UBound(vProdutos)
		if Trim(vProdutos(intCounter)) <> "" then
			if vResp(Ubound(vResp)).produto <> "" then
				redim preserve vResp(Ubound(vResp)+1)
				set vResp(Ubound(vResp)) = new cl_RespConsultaDadosProdutoBD
				end if
			vAux = Split(vProdutos(intCounter), "|")
			vResp(Ubound(vResp)).fabricante = vAux(0)
			vResp(Ubound(vResp)).produto = vAux(1)
			end if
		next
	
	for intCounter=Lbound(vResp) to Ubound(vResp)
		if (Trim(vResp(intCounter).produto) <> "") And _
		   (Trim(vResp(intCounter).status) = "") then
			
			blnProdutoCompostoProcessadoOk = False

		'	CONSULTA NO CADASTRO DE PRODUTOS COMPOSTOS (E-COMMERCE)
			if isLojaHabilitadaProdCompostoECommerce(strLoja) then
				strSql = "SELECT " & _
							"*" & _
						" FROM t_EC_PRODUTO_COMPOSTO" & _
						" WHERE" & _
							" (produto_composto = '" & vResp(intCounter).produto & "')"
				if Trim(vResp(intCounter).fabricante) <> "" then
					strSql = strSql & _
							" AND (fabricante_composto = '" & vResp(intCounter).fabricante & "')"
					end if
				if rs.State <> 0 then rs.Close
				rs.open strSql, cn
				if Not rs.Eof then
					vl_prod_composto_preco_lista_loja = 0
				'	PRODUTO ESTÁ CADASTRADO NA TABELA DE PRODUTO COMPOSTO (E-COMMERCE), ENTÃO É FEITA A LOCALIZAÇÃO DOS PRODUTOS QUE O COMPÕEM E O CÁLCULO DO PREÇO
					strSql = "SELECT" & _
								" fabricante_item," & _
								" produto_item," & _
								" qtde," & _
								" Coalesce(preco_lista_loja,-1) AS preco_lista_loja" & _
							" FROM (" & _
								"SELECT" & _
									" fabricante_item," & _
									" produto_item," & _
									" qtde," & _
									" (" & _
										"SELECT" & _
											" preco_lista" & _
										" FROM t_PRODUTO_LOJA tPL" & _
										" WHERE" & _
											" (CONVERT(smallint,loja) = " & strLoja & ")" & _
											" AND (tPL.fabricante=tECPCI.fabricante_item)" & _
											" AND (tPL.produto=tECPCI.produto_item)" & _
									") AS preco_lista_loja" & _
								" FROM t_EC_PRODUTO_COMPOSTO_ITEM tECPCI" & _
								" WHERE" & _
									" (fabricante_composto = '" & Trim("" & rs("fabricante_composto")) & "')" & _
									" AND (produto_composto = '" & Trim("" & rs("produto_composto")) & "')" & _
									" AND (excluido_status = 0)" & _
								") t"
					if rs2.State <> 0 then rs2.Close
					rs2.open strSql, cn
					if Not rs2.Eof then
						blnProdCompostoPrecoListaItensOk = True
						do while Not rs2.Eof
							if rs2("preco_lista_loja") < 0 then
								blnProdCompostoPrecoListaItensOk = False
							else
								vl_prod_composto_preco_lista_loja = vl_prod_composto_preco_lista_loja + rs2("qtde") * rs2("preco_lista_loja")
								end if
							rs2.MoveNext
							loop

						if blnProdCompostoPrecoListaItensOk then
							blnProdutoCompostoProcessadoOk = True
							if Not blnFabricanteInformado then vResp(intCounter).fabricante = Trim("" & rs("fabricante_composto"))
							vResp(intCounter).descricao = Trim("" & rs("descricao"))
							vResp(intCounter).tabela_origem = "t_EC_PRODUTO_COMPOSTO"
							if vResp(intCounter).status = "" then
								vResp(intCounter).status = "OK"
								vResp(intCounter).precoLista = formata_moeda(vl_prod_composto_preco_lista_loja)
								end if
							end if
						end if
					end if
				end if ' if isLojaHabilitadaProdCompostoECommerce(strLoja)
			
		'	SE O CÓDIGO SE REFERE A UM PRODUTO COMPOSTO (E-COMMERCE) E FOI PROCESSADO CORRETAMENTE, NÃO DEVE FAZER O PROCESSAMENTO A SEGUIR
			if Not blnProdutoCompostoProcessadoOk then
				strSql = _
					"SELECT " & _
						"*" & _
					" FROM t_PRODUTO" & _
						" INNER JOIN t_PRODUTO_LOJA" & _
							" ON (t_PRODUTO.fabricante=t_PRODUTO_LOJA.fabricante) AND (t_PRODUTO.produto=t_PRODUTO_LOJA.produto)" & _
					" WHERE" & _
						" (CONVERT(smallint,loja) = " & strLoja & ")" & _
						" AND (t_PRODUTO.produto = '" & vResp(intCounter).produto & "')"
				if Trim(vResp(intCounter).fabricante) <> "" then
					blnFabricanteInformado = True
					strSql = strSql & _
						" AND (t_PRODUTO.fabricante = '" & vResp(intCounter).fabricante & "')"
				else
					blnFabricanteInformado = False
					strSql = strSql & _
						" AND (t_PRODUTO.excluido_status = 0)" & _
						" AND (t_PRODUTO_LOJA.excluido_status = 0)"
					end if
				if rs.State <> 0 then rs.Close
				rs.open strSql, cn
				if rs.Eof then
				'	CONSULTA NO CADASTRO DE PRODUTOS COMPOSTOS (E-COMMERCE)
					if isLojaHabilitadaProdCompostoECommerce(strLoja) then
						strSql = "SELECT " & _
									"*" & _
								" FROM t_EC_PRODUTO_COMPOSTO" & _
								" WHERE" & _
									" (produto_composto = '" & vResp(intCounter).produto & "')"
						if Trim(vResp(intCounter).fabricante) <> "" then
							strSql = strSql & _
									" AND (fabricante_composto = '" & vResp(intCounter).fabricante & "')"
							end if
						if rs.State <> 0 then rs.Close
						rs.open strSql, cn
						if rs.Eof then
							vResp(intCounter).status = "ERR"
							vResp(intCounter).codigo_erro = "1"
							vResp(intCounter).msg_erro = "Produto " & vResp(intCounter).produto & " não localizado."
						else
							if Not blnFabricanteInformado then vResp(intCounter).fabricante = Trim("" & rs("fabricante_composto"))
							vResp(intCounter).descricao = Trim("" & rs("descricao"))
							vResp(intCounter).tabela_origem = "t_EC_PRODUTO_COMPOSTO"
							if vResp(intCounter).status = "" then
								vResp(intCounter).status = "OK"
								end if
							end if
					else
						vResp(intCounter).status = "ERR"
						vResp(intCounter).codigo_erro = "2"
						vResp(intCounter).msg_erro = "Produto " & vResp(intCounter).produto & " não localizado para a loja " & strLoja & "."
						end if
				else
					if Not blnFabricanteInformado then vResp(intCounter).fabricante = Trim("" & rs("fabricante"))
					vResp(intCounter).descricao = Trim("" & rs("descricao"))
					vResp(intCounter).descricao_html = produto_formata_descricao_em_html(Trim("" & rs("descricao_html")))
					vResp(intCounter).tabela_origem = "t_PRODUTO"
					if vResp(intCounter).status = "" then
						vResp(intCounter).status = "OK"
						vResp(intCounter).precoLista = formata_moeda(rs("preco_lista"))
						end if
					if Not blnFabricanteInformado then
						rs.MoveNext
						if Not rs.Eof then
							vResp(intCounter).status = "ERR"
							vResp(intCounter).codigo_erro = "3"
							vResp(intCounter).msg_erro = "Há mais de um produto cadastrado com o código " & vResp(intCounter).produto & "!! É necessário informar o código do fabricante para resolver a ambiguidade!!"
							end if
						end if
					end if
				end if
			end if 'if Not blnProdutoCompostoProcessadoOk
		next
	
'	MONTA A RESPOSTA
	strResp = ""
	
	for intCounter=Lbound(vResp) to Ubound(vResp)
		if Trim(vResp(intCounter).produto) <> "" then
			strResp = strResp & _
					  "<ItemConsulta>" & _
						"<fabricante>" & _
							vResp(intCounter).fabricante & _
						"</fabricante>" & _
						"<produto>" & _
							vResp(intCounter).produto & _
						"</produto>" & _
						"<status>" & _
							vResp(intCounter).status & _
						"</status>" & _
						"<precoLista>" & _
							vResp(intCounter).precoLista & _
						"</precoLista>" & _
						"<descricao>" & _
							vResp(intCounter).descricao & _
						"</descricao>" & _
						"<descricao_html>" & _
							"<![CDATA[" & _
							vResp(intCounter).descricao_html & _
							"]]>" & _
						"</descricao_html>" & _
						"<tabela_origem>" & _
							vResp(intCounter).tabela_origem & _
						"</tabela_origem>" & _
						"<codigo_erro>" & _
							vResp(intCounter).codigo_erro & _
						"</codigo_erro>" & _
						"<msg_erro>" & _
							vResp(intCounter).msg_erro & _
						"</msg_erro>" & _
					  "</ItemConsulta>"
			end if
		next


'	HÁ RESPOSTA?
	if strResp <> "" then 
		Response.ContentType="text/xml"
		strResp = "<?xml version='1.0' encoding='ISO-8859-1'?>" & _
				  "<TabelaResultado>" & _
				  strResp & _
				  "</TabelaResultado>"
		end if
		
	Response.Write strResp

'	FECHA CONEXAO COM O BANCO DE DADOS
	if rs.State <> 0 then rs.Close
	set rs = nothing

	if rs2.State <> 0 then rs2.Close
	set rs2 = nothing

	cn.Close
	set cn = nothing

%>
