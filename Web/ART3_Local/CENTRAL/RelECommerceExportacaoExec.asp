<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/estoque.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =================================================================
'	  RelECommerceExportacaoExec.asp
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
	
	const MSO_NUMBER_FORMAT_PERC = "\#\#0\.0%"
	const MSO_NUMBER_FORMAT_INTEIRO = "\#\#\#\,\#\#\#\,\#\#0"
	const MSO_NUMBER_FORMAT_MOEDA = "\#\#\#\,\#\#\#\,\#\#0\.00"
	const MSO_NUMBER_FORMAT_TEXTO = "\@"
	
	dim usuario
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO)

'	CONECTA COM O BANCO DE DADOS
	dim cn
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	if Not operacao_permitida(OP_CEN_REL_ECOMMERCE_EXPORTACAO, s_lista_operacoes_permitidas) then 
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	const OPCAO_UM_CODIGO = "UM"
	const OPCAO_FAIXA_CODIGOS = "FAIXA"
	
	dim alerta
	dim s, s_aux, s_filtro, lista_loja, s_filtro_loja, v_loja, v, i
	dim ckb_normais, ckb_compostos
	dim rb_fabricante, c_fabricante, c_fabricante_de, c_fabricante_ate
	dim rb_produto, c_produto, c_produto_de, c_produto_ate
	dim rb_grupo, c_grupo, c_grupo_de, c_grupo_ate, c_loja, rb_saida, c_qtde_corte_estoque, c_fabricante_ignorado, v_fabricante_ignorado
    dim rb_percentual_majoracao, c_percentual_majoracao

	alerta = ""

	rb_fabricante = Ucase(Trim(Request.Form("rb_fabricante")))
	c_fabricante = retorna_so_digitos(Trim(Request.Form("c_fabricante")))
	c_fabricante_de = retorna_so_digitos(Trim(Request.Form("c_fabricante_de")))
	c_fabricante_ate = retorna_so_digitos(Trim(Request.Form("c_fabricante_ate")))
	
	rb_produto = Ucase(Trim(Request.Form("rb_produto")))
	c_produto = Ucase(Trim(Request.Form("c_produto")))
	c_produto_de = Ucase(Trim(Request.Form("c_produto_de")))
	c_produto_ate = Ucase(Trim(Request.Form("c_produto_ate")))
	
	rb_grupo = Ucase(Trim(Request.Form("rb_grupo")))
	c_grupo = Ucase(Trim(Request.Form("c_grupo")))
	c_grupo_de = Ucase(Trim(Request.Form("c_grupo_de")))
	c_grupo_ate = Ucase(Trim(Request.Form("c_grupo_ate")))

	ckb_normais = Trim(Request.Form("ckb_normais"))
	ckb_compostos = Trim(Request.Form("ckb_compostos"))
    c_loja = Trim(Request.Form("c_loja"))

    rb_saida = Request.Form("rb_saida")
    c_qtde_corte_estoque = Trim(Request.Form("c_qtde_corte_estoque"))
    c_fabricante_ignorado = Trim(Request.Form("c_fabricante_ignorado"))

    rb_percentual_majoracao = Trim(Request.Form("rb_percentual_majoracao"))
    c_percentual_majoracao = Trim(Request.Form("c_percentual_majoracao"))
    
    if c_qtde_corte_estoque = "" then c_qtde_corte_estoque = 0
    if c_percentual_majoracao = "" then c_percentual_majoracao = 0


	if alerta = "" then
        call set_default_valor_texto_bd(usuario, "RelECommerceExportacao|c_loja", c_loja)
        call set_default_valor_texto_bd(usuario, "RelECommerceExportacao|c_fabricante_ignorado", c_fabricante_ignorado)
        call set_default_valor_texto_bd(usuario, "RelECommerceExportacao|c_percentual_majoracao", c_percentual_majoracao)
        call set_default_valor_texto_bd(usuario, "RelECommerceExportacao|rb_percentual_majoracao", rb_percentual_majoracao)
        if rb_saida <> 1 then call set_default_valor_texto_bd(usuario, "RelECommerceExportacao|c_qtde_corte_estoque", c_qtde_corte_estoque)        

		Response.ContentType = "application/csv"
		Response.AddHeader "Content-Disposition", "attachment; filename=ECommerce_" & formata_data_yyyymmdd(Now) & "_" & formata_hora_hhnnss(Now) & ".csv"
		consulta_executa
		Response.End

		end if

    if (c_loja = "") then
        alerta = "Informe a LOJA!!"
    end if

' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________

' _____________________________________
' CONSULTA EXECUTA
'
sub consulta_executa
const NOME_CAMPO_FABRICANTE = "§FABR§"
const NOME_CAMPO_PRODUTO = "§PROD§"
dim r, t, rs, tPCI, msg_erro
dim s, s_sql, n_reg, n_reg_total, x
dim s_where, s_where_aux, s_where_prod_normal, s_where_prod_composto
dim vl_preco_lista_composto, qtde_estoque_venda_composto, qtde_estoque_venda_aux, coeficiente, fabricante_a, produto
dim blnPularProdutoComposto, preco_lista

	call cria_recordset_otimista(t, msg_erro)
	call cria_recordset_otimista(tPCI, msg_erro)
	call cria_recordset_otimista(rs, msg_erro)

'	CRITÉRIOS COMUNS
'	================
	s_where = ""
	
'	GRUPO DE PRODUTOS
	if rb_grupo = OPCAO_UM_CODIGO then
		if c_grupo <> "" then
			if s_where <> "" then s_where = s_where & " AND"
			s_where = s_where & " (tP.grupo = '" & c_grupo & "')"

			if s_where <> "" then s_where = s_where & " AND"
			s_where = s_where & " (tP.grupo IS NOT NULL)"
			end if
	elseif rb_grupo = OPCAO_FAIXA_CODIGOS then
		if (c_grupo_de<>"") Or (c_grupo_ate<>"") then
			s = ""
			if c_grupo_de<>"" then
				if s <> "" then s = s & " AND"
				s = s & " (tP.grupo >= '" & c_grupo_de & "')"
				end if
			if c_grupo_ate<>"" then
				if s <> "" then s = s & " AND"
				s = s & " (tP.grupo <= '" & c_grupo_ate & "')"
				end if
			if s <> "" then 
				s = " (" & s & ")"
				if s_where <> "" then s_where = s_where & " AND"
				s_where = s_where & s
				end if
			end if
		end if

'	CRITÉRIOS P/ PEDIDOS DE VENDA NORMAIS E P/ DEVOLUÇÕES
'	=====================================================
	s_where_aux = ""

'	FABRICANTE
	if rb_fabricante = OPCAO_UM_CODIGO then
		if c_fabricante <> "" then
			if s_where_aux <> "" then s_where_aux = s_where_aux & " AND"
			s_where_aux = s_where_aux & " (" & NOME_CAMPO_FABRICANTE & " = '" & c_fabricante & "')"
			end if
	elseif rb_fabricante = OPCAO_FAIXA_CODIGOS then
		if (c_fabricante_de<>"") Or (c_fabricante_ate<>"") then
			s = ""
			if c_fabricante_de <> "" then
				if s <> "" then s = s & " AND"
				s = s & " (" & NOME_CAMPO_FABRICANTE & " >= '" & c_fabricante_de & "')"
				end if
			if c_fabricante_ate <> "" then
				if s <> "" then s = s & " AND"
				s = s & " (" & NOME_CAMPO_FABRICANTE & " <= '" & c_fabricante_ate & "')"
				end if
			if s <> "" then 
				s = " (" & s & ")"
				if s_where_aux <> "" then s_where_aux = s_where_aux & " AND"
				s_where_aux = s_where_aux & s
				end if
			end if
		end if

'   FABRICANTES IGNORADOS
    if c_fabricante_ignorado <> "" then
        v_fabricante_ignorado = split(c_fabricante_ignorado, ", ")
        s = ""
        for i = LBound(v_fabricante_ignorado) to UBound(v_fabricante_ignorado)
            if s <> "" then s = s & ", "
            s = s & "'" & v_fabricante_ignorado(i) & "'"
        next
        s = " (" & NOME_CAMPO_FABRICANTE & " NOT IN (" & s & "))"
        if s_where_aux <> "" then s_where_aux = s_where_aux & " AND"
        s_where_aux = s_where_aux & s
    end if


'	PRODUTO
	if rb_produto = OPCAO_UM_CODIGO then
		if c_produto <> "" then
			if s_where_aux <> "" then s_where_aux = s_where_aux & " AND"
			s_where_aux = s_where_aux & " (" & NOME_CAMPO_PRODUTO & " = '" & c_produto & "')"
			end if
	elseif rb_produto = OPCAO_FAIXA_CODIGOS then
		if (c_produto_de<>"") Or (c_produto_ate<>"") then
			s = ""
			if c_produto_de <> "" then
				if s <> "" then s = s & " AND"
				s = s & " (" & NOME_CAMPO_PRODUTO & " >= '" & c_produto_de & "')"
				end if
			if c_produto_ate <> "" then
				if s <> "" then s = s & " AND"
				s = s & " (" & NOME_CAMPO_PRODUTO & " <= '" & c_produto_ate & "')"
				end if
			if s <> "" then 
				s = " (" & s & ")"
				if s_where_aux <> "" then s_where_aux = s_where_aux & " AND"
				s_where_aux = s_where_aux & s
				end if
			end if
		end if


	s_where_prod_normal = s_where_aux
	s_where_prod_composto = s_where_aux
	if s_where_aux <> "" then
		s_where_prod_normal = replace(s_where_prod_normal, NOME_CAMPO_FABRICANTE, "tP.fabricante")
		s_where_prod_composto = replace(s_where_prod_composto, NOME_CAMPO_FABRICANTE, "tP.fabricante")

		s_where_prod_normal = replace(s_where_prod_normal, NOME_CAMPO_PRODUTO, "tP.produto")
		s_where_prod_composto = replace(s_where_prod_composto, NOME_CAMPO_PRODUTO, "tP.produto")
		end if

    fabricante_a = "XXXXXXXXXXX"

'	MONTA CONSULTA
'	==============

    x = ""
    if rb_saida <> 1 then      
            x = "sku;qty;is_in_stock;price" & vbcrlf
        end if    

'	INÍCIO TRATAMENTO PRODUTOS NORMAIS
'	OBS: NO TRECHO DO SQL EM QUE SE EXCLUEM OS PRODUTOS QUE FAÇAM PARTE DE UM PRODUTO COMPOSTO, ASSUME-SE QUE
'	==== O CÓDIGO DO PRODUTO (CAMPO 'PRODUTO') É ÚNICO, NÃO HAVENDO A NECESSIDADE DO USO CONJUNTO DO CAMPO 'FABRICANTE'
	if ckb_normais = "1" then
		s = s_where
		if (s <> "") And (s_where_prod_normal <> "") then s = s & " AND"
		s = s & s_where_prod_normal
		if s <> "" then s = " AND" & s
		s_sql = "SELECT" & _
					" tP.fabricante," & _
					" tP.produto," & _
					" tP.descricao," & _
					" tF.nome AS nome_fabricante," & _
                    " tPL.preco_lista," & _
					" Coalesce((SELECT TOP 1 vl_custo2 FROM t_ESTOQUE_ITEM tEI WHERE (tEI.fabricante=tP.fabricante) AND (tEI.produto=tP.produto) AND (vl_custo2 > 0) ORDER BY id_estoque DESC), 0) AS vl_custo2," & _
					" Coalesce((SELECT Sum(qtde-qtde_utilizada) FROM t_ESTOQUE_ITEM tEI WHERE (tEI.fabricante=tP.fabricante) AND (tEI.produto=tP.produto) AND ((qtde-qtde_utilizada)>0)), 0) AS qtde_estoque_venda" & _
				" FROM t_PRODUTO_LOJA tPL" & _
					" INNER JOIN t_PRODUTO tP ON (tPL.fabricante = tP.fabricante) AND (tPL.produto = tP.produto)" & _
					" INNER JOIN t_FABRICANTE tF ON (tP.fabricante = tF.fabricante)" & _
				" WHERE" & _
					" (Upper(Coalesce(vendavel, '')) = 'S')" & _
					" AND (Upper(Coalesce(descontinuado, '')) <> 'S')" & _
					" AND (loja = '" & c_loja & "')" & _
					" AND (preco_lista > 0)" & _
					" AND (descricao <> '.')" & _
					" AND (descricao <> 'CÓDIGO VAGO' COLLATE Latin1_General_CI_AI)" & _
					" AND (descricao <> 'Renegociação' COLLATE Latin1_General_CI_AI)" & _
					" AND (tP.produto NOT IN (SELECT DISTINCT produto_composto AS produto FROM t_EC_PRODUTO_COMPOSTO))" & _
					" AND (tP.produto NOT IN (SELECT DISTINCT produto_item AS produto FROM t_EC_PRODUTO_COMPOSTO_ITEM))" & _
					s
		
		s_sql = "SELECT " & _
					"*" & _
				" FROM (" & s_sql & ") t" & _
				" WHERE" & _
					" (preco_lista > 0)" & _
				" ORDER BY" & _
					" fabricante," & _
					" produto"
		
		n_reg = 0
		n_reg_total = 0
        coeficiente = 1
            
	    
		set r = cn.execute(s_sql)
		do while Not r.Eof
		
            if Not r.EOF then
                if fabricante_a <> r("fabricante") then
                    rs.Open "SELECT coeficiente FROM t_PERCENTUAL_CUSTO_FINANCEIRO_FORNECEDOR WHERE" & _
                                " fabricante='" & r("fabricante") & "'" & _
                                " AND tipo_parcelamento='SE'" & _
                                " AND qtde_parcelas='10'", cn

                        if Not rs.Eof then
                            coeficiente = rs("coeficiente")
                        else
                            coeficiente = 1
                        end if
                    end if
            end if
		  ' CONTAGEM
			n_reg = n_reg + 1
			n_reg_total = n_reg_total + 1

		 '> CÓDIGO DO PRODUTO
            if rb_saida <> 1 then
                produto = Trim("" & r("produto"))
                for i = 0 to Len(produto)
                    if Mid(produto, 1, 1) = "0" then
                        produto = Right(produto, Len(produto)-1)
                    else
                        exit for
                    end if
                    i = i + 1
                next
			    x = x & produto & ";"                
            else
			    x = x & Trim("" & r("produto")) & ";"
            end if

		 '> VALOR CUSTO
		'	EXPORTAR VALOR UTILIZANDO '.' COMO SEPARADOR DECIMAL
		'	s = bd_formata_moeda(r("vl_custo2"))
		'	x = x & s & ";"

         '> PREÇO LISTA
		'	EXPORTAR VALOR UTILIZANDO '.' COMO SEPARADOR DECIMAL
            if rb_saida = 1 then
			    s = bd_formata_moeda(r("preco_lista") * coeficiente)
			    x = x & s & ";"
	        end if		

		 '> SALDO ESTOQUE
			s = Cstr(r("qtde_estoque_venda"))
			x = x & s & ";"
			
		 '> DESCRIÇÃO
            if rb_saida = 1 then
			    s = Trim("" & r("descricao"))
			    s = substitui_caracteres(s, ";", ",")
			    x = x & s & ";"
            end if
			
		 '> NOME DO FABRICANTE
            if rb_saida = 1 then            
			    s = Trim("" & r("nome_fabricante"))
			    s = substitui_caracteres(s, ";", ",")
			    x = x & s
            end if

          '> FLAG IS IN STOCK
            if rb_saida <> 1 then
                if ((r("qtde_estoque_venda") >= CInt(c_qtde_corte_estoque)) And (r("qtde_estoque_venda") > 0)) then
                    s = "1"
                else
                    s = "0"
                end if
                x = x & s & ";"
            end if

            if rb_saida <> 1 then
                preco_lista = r("preco_lista") * coeficiente
                if rb_percentual_majoracao = "1" then
                        s = bd_formata_moeda(((c_percentual_majoracao /100) * preco_lista) + preco_lista)
                    elseif rb_percentual_majoracao = "2" then
                        s = bd_formata_moeda(preco_lista - ((c_percentual_majoracao /100) * preco_lista))
                    else
                        s = bd_formata_moeda(preco_lista)                    
                    end if
                
                x = x & s
            end if
                    
			
			x = x & vbcrlf
			
			if (n_reg_total mod 100) = 0 then
				Response.Write x
				x = ""
				end if

			if c_fabricante = "" then
		        if rs.State <> 0 then rs.Close
            end if
            fabricante_a = r("fabricante")
			r.MoveNext
			loop
			
        
		if r.State <> 0 then r.Close
		set r=nothing
		end if


'	INÍCIO TRATAMENTO PRODUTOS COMPOSTOS
'	OBS: NO TRECHO DO SQL EM QUE SE EXCLUEM OS PRODUTOS INDISPONÍVEIS, ASSUME-SE QUE O CÓDIGO DO
'	==== PRODUTO (CAMPO 'PRODUTO') É ÚNICO, NÃO HAVENDO A NECESSIDADE DO USO CONJUNTO DO CAMPO 'FABRICANTE'
	if ckb_compostos = "2" then
		s = s_where
		if (s <> "") And (s_where_prod_composto <> "") then s = s & " AND"
		s = s & s_where_prod_composto
		if s <> "" then s = " AND" & s
		s_sql = "SELECT" & _
					" tECPC.fabricante_composto," & _
					" tECPC.produto_composto," & _
					" tECPC.descricao," & _
					" tF.nome AS nome_fabricante" & _
				" FROM t_EC_PRODUTO_COMPOSTO tECPC" & _
					" LEFT JOIN t_FABRICANTE tF ON (tECPC.fabricante_composto = tF.fabricante)" & _
				" ORDER BY" & _
					" tECPC.fabricante_composto," & _
					" tECPC.produto_composto"

        
		
		n_reg = 0
		set r = cn.execute(s_sql)
		do while Not r.Eof
			blnPularProdutoComposto = False
			vl_preco_lista_composto = 0
			qtde_estoque_venda_composto = -1
			
			s_sql = "SELECT " & _
						" fabricante_item," & _
						" produto_item," & _
						" qtde" & _
					" FROM t_EC_PRODUTO_COMPOSTO_ITEM" & _
					" WHERE" & _
						" (fabricante_composto = '" & Trim("" & r("fabricante_composto")) & "')" & _
						" AND (produto_composto = '" & Trim("" & r("produto_composto")) & "')" & _
					" ORDER BY" & _
						" fabricante_item," & _
						" produto_item"
			if tPCI.State <> 0 then tPCI.Close
			tPCI.Open s_sql, cn
			do while Not tPCI.Eof
				s = s_where
				if (s <> "") And (s_where_prod_composto <> "") then s = s & " AND"
				s = s & s_where_prod_composto
				if s <> "" then s = " AND" & s
				s_sql = "SELECT" & _
							" tP.fabricante," & _
							" tP.produto," & _
                            " tPL.preco_lista," & _
							" Coalesce((SELECT TOP 1 vl_custo2 FROM t_ESTOQUE_ITEM tEI WHERE (tEI.fabricante=tP.fabricante) AND (tEI.produto=tP.produto) AND (vl_custo2 > 0) ORDER BY id_estoque DESC), 0) AS vl_custo2," & _
							" Coalesce((SELECT Sum(qtde-qtde_utilizada) FROM t_ESTOQUE_ITEM tEI WHERE (tEI.fabricante=tP.fabricante) AND (tEI.produto=tP.produto) AND ((qtde-qtde_utilizada)>0)), 0) AS qtde_estoque_venda" & _
						" FROM t_PRODUTO_LOJA tPL" & _
							" INNER JOIN t_PRODUTO tP ON (tPL.fabricante = tP.fabricante) AND (tPL.produto = tP.produto)" & _
						" WHERE" & _
							" (Upper(Coalesce(vendavel, '')) = 'S')" & _
							" AND (Upper(Coalesce(descontinuado, '')) <> 'S')" & _
							" AND (loja = '" & c_loja & "')" & _
							" AND (preco_lista > 0)" & _
							" AND (descricao <> '.')" & _
							" AND (descricao <> 'CÓDIGO VAGO' COLLATE Latin1_General_CI_AI)" & _
							" AND (descricao <> 'Renegociação' COLLATE Latin1_General_CI_AI)" & _
							" AND (tP.fabricante = '" & Trim("" & tPCI("fabricante_item")) & "')" & _
							" AND (tP.produto = '" & Trim("" & tPCI("produto_item")) & "')" & _
							s
				
				s_sql = "SELECT " & _
							"*" & _
						" FROM (" & s_sql & ") t" & _
						" WHERE" & _
							" (preco_lista > 0)" & _
						" ORDER BY" & _
							" fabricante," & _
							" produto"
				if t.State <> 0 then t.Close
				t.Open s_sql, cn
                if Not t.EOF then
                    if fabricante_a <> t("fabricante") then
                        if rs.State <> 0 then rs.Close
                        rs.Open "SELECT coeficiente FROM t_PERCENTUAL_CUSTO_FINANCEIRO_FORNECEDOR WHERE" & _
                                    " fabricante='" & t("fabricante") & "'" & _
                                    " AND tipo_parcelamento='SE'" & _
                                    " AND qtde_parcelas='10'", cn
        
        
                            if Not rs.Eof then
                                coeficiente = rs("coeficiente")
                            else
                                coeficiente = 1
                            end if
                     end if
                end if

				if t.Eof then
					blnPularProdutoComposto = True
				else
					vl_preco_lista_composto = vl_preco_lista_composto + ((tPCI("qtde") * t("preco_lista"))) * coeficiente
					qtde_estoque_venda_aux = t("qtde_estoque_venda") \ tPCI("qtde")
					if qtde_estoque_venda_composto = -1 then
						qtde_estoque_venda_composto = qtde_estoque_venda_aux
					else
						if qtde_estoque_venda_aux < qtde_estoque_venda_composto then
							qtde_estoque_venda_composto = qtde_estoque_venda_aux
							end if
						end if
					end if
                if c_fabricante = "" then
				    if rs.State <> 0 then rs.Close
                end if
				if blnPularProdutoComposto then exit do
				fabricante_a = t("fabricante")
				tPCI.MoveNext
				loop
			
            

			if Not blnPularProdutoComposto then
			'	CONTAGEM
				n_reg = n_reg + 1
				n_reg_total = n_reg_total + 1

			 '> CÓDIGO DO PRODUTO
		    if rb_saida <> 1 then
                produto = Trim("" & r("produto_composto"))
                for i = 0 to Len(produto)
                    if Mid(produto, 1, 1) = "0" then
                        produto = Right(produto, Len(produto)-1)
                    else
                        exit for
                    end if
                    i = i + 1
                next
			    x = x & produto & ";"                
            else
			    x = x & Trim("" & r("produto_composto")) & ";"
            end if

			 '> VALOR CUSTO
			'	EXPORTAR VALOR UTILIZANDO '.' COMO SEPARADOR DECIMAL
                if rb_saida = 1 then                
				    s = bd_formata_moeda(vl_preco_lista_composto)
				    x = x & s & ";"
                end if
				
			 '> SALDO ESTOQUE
				s = Cstr(qtde_estoque_venda_composto)
				x = x & s & ";"
				
			 '> DESCRIÇÃO
                if rb_saida = 1 then                                
				    s = Trim("" & r("descricao"))
				    s = substitui_caracteres(s, ";", ",")
				    x = x & s & ";"
                end if
				
			 '> NOME DO FABRICANTE
                if rb_saida = 1 then                                
				    s = Trim("" & r("nome_fabricante"))
				    s = substitui_caracteres(s, ";", ",")
				    x = x & s
                end if

              '> FLAG IS IN STOCK
                if rb_saida <> 1 then
                    if ((qtde_estoque_venda_composto >= CInt(c_qtde_corte_estoque) And qtde_estoque_venda_composto > 0)) then
                        s = "1"
                    else
                        s = "0"
                    end if
                    x = x & s & ";"
                end if

                if rb_saida <> 1 then
                   if rb_percentual_majoracao = "1" then
                        s = bd_formata_moeda(((c_percentual_majoracao /100) * vl_preco_lista_composto) + vl_preco_lista_composto)
                    elseif rb_percentual_majoracao = "2" then
                        s = bd_formata_moeda(vl_preco_lista_composto - ((c_percentual_majoracao /100) * vl_preco_lista_composto))
                    else
                        s = bd_formata_moeda(vl_preco_lista_composto)
                        
                    end if

                   x = x & s
                end if
				
				x = x & vbcrlf
				
				if (n_reg_total mod 100) = 0 then
					Response.Write x
					x = ""
					end if
				
				end if
			
			r.MoveNext
			loop
		end if

'	MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	if n_reg_total = 0 then
		x = x & "NENHUM PRODUTO ENCONTRADO"
		end if
	
	Response.write x
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



<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
window.status='Aguarde, executando a consulta ...';
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
<!-- *************************************************************************** -->
<!-- **********  PÁGINA PARA EXIBIR RESULTADO   (apenas para testes)  ********** -->
<!-- *************************************************************************** -->
<body onload="window.status='Concluído';">

<center>

<form id="fREL" name="fREL" method="post">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="rb_fabricante" id="rb_fabricante" value="<%=rb_fabricante%>">
<input type="hidden" name="c_fabricante" id="c_fabricante" value="<%=c_fabricante%>">
<input type="hidden" name="c_fabricante_de" id="c_fabricante_de" value="<%=c_fabricante_de%>">
<input type="hidden" name="c_fabricante_ate" id="c_fabricante_ate" value="<%=c_fabricante_ate%>">
<input type="hidden" name="rb_produto" id="rb_produto" value="<%=rb_produto%>">
<input type="hidden" name="c_produto" id="c_produto" value="<%=c_produto%>">
<input type="hidden" name="c_produto_de" id="c_produto_de" value="<%=c_produto_de%>">
<input type="hidden" name="c_produto_ate" id="c_produto_ate" value="<%=c_produto_ate%>">
<input type="hidden" name="rb_grupo" id="rb_grupo" value="<%=rb_grupo%>">
<input type="hidden" name="c_grupo" id="c_grupo" value="<%=c_grupo%>">
<input type="hidden" name="c_grupo_de" id="c_grupo_de" value="<%=c_grupo_de%>">
<input type="hidden" name="c_grupo_ate" id="c_grupo_ate" value="<%=c_grupo_ate%>">


<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">E-Commerce: Exportação da Tabela de Produtos</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>

<!-- FILTROS -->
<% 
	s_filtro = "<table width='649' cellpadding='0' cellspacing='0' style='border-bottom:1px solid black' border='0'>" & chr(13)
	
'	FABRICANTE
	s = ""
	if rb_fabricante = OPCAO_UM_CODIGO then
		s = c_fabricante
	elseif rb_fabricante = OPCAO_FAIXA_CODIGOS then
		if (c_fabricante_de<>"") Or (c_fabricante_ate<>"") then
			s_aux = c_fabricante_de
			if s_aux = "" then s_aux = "N.I."
			s = s & s_aux & " a "
			s_aux = c_fabricante_ate
			if s_aux = "" then s_aux = "N.I."
			s = s & s_aux
			end if
		end if
		
	if s = "" then s = "todos"
	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' nowrap><span class='N'>Fabricante:&nbsp;</span></td>" & chr(13) & _
				"		<td align='left' valign='top'><span class='N'>" & s & "</span></td>" & chr(13) & _
				"	</tr>" & chr(13)

'	PRODUTO
	s = ""
	if rb_produto = OPCAO_UM_CODIGO then
		s = c_produto
	elseif rb_produto = OPCAO_FAIXA_CODIGOS then
		if (c_produto_de<>"") Or (c_produto_ate<>"") then
			s_aux = c_produto_de
			if s_aux = "" then s_aux = "N.I."
			s = s & s_aux & " a "
			s_aux = c_produto_ate
			if s_aux = "" then s_aux = "N.I."
			s = s & s_aux
			end if
		end if

	if s = "" then s = "todos"
	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' nowrap><span class='N'>Produto:&nbsp;</span></td>" & chr(13) & _
				"		<td align='left' valign='top'><span class='N'>" & s & "</span></td>" & chr(13) & _
				"	</tr>" & chr(13)
	
'	GRUPO DE PRODUTOS
	s = ""
	if rb_grupo = OPCAO_UM_CODIGO then
		s = c_grupo
	elseif rb_grupo = OPCAO_FAIXA_CODIGOS then
		if (c_grupo_de<>"") Or (c_grupo_ate<>"") then
			s_aux = c_grupo_de
			if s_aux = "" then s_aux = "N.I."
			s = s & s_aux & " a "
			s_aux = c_grupo_ate
			if s_aux = "" then s_aux = "N.I."
			s = s & s_aux
			end if
		end if

	if s = "" then s = "todos"
	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' nowrap><span class='N'>Grupo de Produtos:&nbsp;</span></td>" & chr(13) & _
				"		<td align='left' valign='top'><span class='N'>" & s & "</span></td>" & chr(13) & _
				"	</tr>" & chr(13)

	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' nowrap><span class='N'>Emissão:&nbsp;</span></td>" & chr(13) & _
				"		<td align='left' valign='top' width='99%'><span class='N'>" & formata_data_hora(Now) & "</span></td>" & chr(13) & _
				"	</tr>" & chr(13)

	s_filtro = s_filtro & "</table>" & chr(13)
	Response.Write s_filtro
%>

<!--  RELATÓRIO  -->
<br>
<% consulta_executa %>

<!-- ************   SEPARADOR   ************ -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc" align='left'>&nbsp;</td></tr>
</table>
<br>


<table class="notPrint" width="649" cellspacing="0">
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
'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>
