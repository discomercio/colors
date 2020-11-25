<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ======================================================
'	  RelEstoqueEcommerceExec.asp
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

	const MSO_NUMBER_FORMAT_PERC = "\#\#0\.0%"
	const MSO_NUMBER_FORMAT_INTEIRO = "\#\#\#\,\#\#\#\,\#\#0"
	const MSO_NUMBER_FORMAT_MOEDA = "\#\#\#\,\#\#\#\,\#\#0\.00"
	const MSO_NUMBER_FORMAT_TEXTO = "\@"
    Const VENDA_SHOW_ROOM = "VENDA_SHOW_ROOM"

	dim usuario
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	CONECTA COM O BANCO DE DADOS
	dim cn, rs, msg_erro,tPCI,t
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	If Not cria_recordset_otimista(tPCI, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
    If Not cria_recordset_otimista(t, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	if (Not operacao_permitida(OP_CEN_REL_ESTOQUE2, s_lista_operacoes_permitidas)) then 
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	dim alerta
	dim s, c_fabricante, c_produto, rb_estoque, rb_detalhe, rb_saida,ckb_normais,ckb_compostos,c_loja, c_empresa
	dim cod_fabricante, cod_produto
	dim s_nome_fabricante, s_nome_produto, s_nome_produto_html
	dim blnSaidaExcel
    dim s_where_aux,s_where,s_where_prod_normal,s_where_prod_composto
    dim v_fabricante(),v_codigo(),v_descricao(),v_qtde(),v_valor(),v_produtos()
    dim cont,blnPularProdutoComposto,v_preco_lista(),colspan,classe_cab,classe_td,preco_lista
    dim qtde_estoque_venda_aux,n_reg_total,vRelat(),vl_custo2_composto,qtde_estoque_venda_composto,i
    i = 0
    cont = 0
    colspan=0
    classe_cab=""
    classe_td= ""
    preco_lista=0
    redim v_codigo(0)
   
	c_fabricante = Trim(Request.Form("c_fabricante"))
	c_produto = UCase(Trim(Request.Form("c_produto")))
	rb_estoque = Trim(Request.Form("rb_estoque"))
	rb_detalhe = Trim(Request.Form("rb_detalhe"))
	rb_saida = Trim(Request.Form("rb_saida"))
    ckb_normais =  Trim(Request.Form("ckb_normais"))
    ckb_compostos = Trim(Request.Form("ckb_compostos"))
    c_loja = Trim(Request.Form("c_loja"))
    c_empresa = Trim(Request.Form("c_empresa"))

	blnSaidaExcel = False
	if rb_saida = "XLS" then
		blnSaidaExcel = True
		end if
	
	alerta = ""
	if (c_produto<>"") And (Not IsEAN(c_produto)) then
		if c_fabricante = "" then alerta = "PARA CONSULTAR PELO CÓDIGO INTERNO DE PRODUTO É NECESSÁRIO ESPECIFICAR O FABRICANTE."
		end if
	if alerta = "" then
        if c_loja <> "" then
            s = "SELECT loja FROM t_PRODUTO_LOJA WHERE (loja='" & c_loja & "')"
            if rs.State <> 0 then rs.Close
			rs.open s, cn
            if rs.Eof then
                alerta = "Número da loja incorreto"
            end if       
        end if
    end if
  
	dim s_where_compostos,s_where_normais, v_fabricantes, v_grupos,s_where_normais2
	s_nome_fabricante = ""
	s_where_normais = ""
    s_where_compostos = ""
	v_fabricantes = ""
    s_where_normais2 = ""
	if alerta = "" then   
		if c_fabricante <> "" then		    
		    v_fabricantes = split(c_fabricante, ", ")
            if ckb_compostos <> "" then
		        for cont = 0 to Ubound(v_fabricantes)
                        s = "SELECT fabricante from t_FABRICANTE where (nome = '" & v_fabricantes(cont) & "')"
                        if rs.State <> 0 then rs.Close
                        rs.Open s,cn
                        if s_where_compostos <> "" then s_where_compostos = s_where_compostos & " OR " 	                        
		                s_where_compostos = s_where_compostos & _
			            " (fabricante_composto ='" & rs("fabricante") & "') "
                                           
	            next
            end if
            if ckb_normais <> "" then
                for cont = 0 to Ubound(v_fabricantes)
                        s = "SELECT fabricante from t_FABRICANTE where (nome = '" & v_fabricantes(cont) & "')"
                        if rs.State <> 0 then rs.Close
                        rs.Open s,cn
                        if s_where_normais <> "" then s_where_normais = s_where_normais & " OR " 	                        
		                s_where_normais = s_where_normais & _
			            " (t_ESTOQUE_ITEM.fabricante ='" & rs("fabricante") & "') "
                                           
	            next
                if rb_estoque = VENDA_SHOW_ROOM then
                    for cont = 0 to Ubound(v_fabricantes)
                        s = "SELECT fabricante from t_FABRICANTE where (nome = '" & v_fabricantes(cont) & "')"
                        if rs.State <> 0 then rs.Close
                        rs.Open s,cn
                        if s_where_normais2 <> "" then s_where_normais2 = s_where_normais2 & " OR " 	                        
		                s_where_normais2 = s_where_normais2 & _
			            " (t_ESTOQUE_MOVIMENTO.fabricante ='" & rs("fabricante") & "') "
                                           
	                next
                end if    
            end if       
		end if
    end if
    
    if alerta = "" then
		call set_default_valor_texto_bd(usuario, "RelEstoqueEcommerceFiltro|c_fabricante", c_fabricante)
		call set_default_valor_texto_bd(usuario, "RelEstoqueEcommerceFiltro|rb_estoque", rb_estoque)
		call set_default_valor_texto_bd(usuario, "RelEstoqueEcommerceFiltro|rb_detalhe", rb_detalhe)
		call set_default_valor_texto_bd(usuario, "RelEstoqueEcommerceFiltro|rb_saida", rb_saida)
		call set_default_valor_texto_bd(usuario, "RelEstoqueEcommerceFiltro|ckb_normais", ckb_normais)
		call set_default_valor_texto_bd(usuario, "RelEstoqueEcommerceFiltro|ckb_compostos", ckb_compostos)
        call set_default_valor_texto_bd(usuario, "RelEstoqueEcommerceFiltro|c_loja", c_loja)
		end if	
	if alerta = "" then
	'	DEFAULT
		cod_produto = c_produto
		cod_fabricante = c_fabricante
		
		if IsEAN(c_produto) then
			s = "SELECT fabricante, produto, ean FROM t_PRODUTO WHERE (ean='" & c_produto & "')"
			if rs.State <> 0 then rs.Close
			rs.open s, cn
			if rs.Eof then
				alerta = "Produto com código EAN " & c_produto & " não está cadastrado."
			else
				if c_fabricante <> "" then
					if c_fabricante <> Trim("" & rs("fabricante")) then 
						alerta = "Produto " & Trim("" & rs("produto")) & " (EAN: " & _
								 Trim("" & rs("ean")) & ") não pertence ao fabricante " & c_fabricante & "."
						end if
					end if
				
				if alerta = "" then
				'	OBTÉM O CÓDIGO INTERNO DE PRODUTO
					cod_fabricante = Trim("" & rs("fabricante"))
					cod_produto = Trim("" & rs("produto"))
					end if
				end if
			end if
		end if
    '-----------------
	if alerta = "" then
		if cod_fabricante <> "" then          
		    s_nome_fabricante = fabricante_descricao(cod_fabricante)            
		else
			s_nome_fabricante = ""
		end if
				
		if cod_produto <> "" then
            t.Open "   Select descricao from t_EC_PRODUTO_COMPOSTO where produto_composto = " & cod_produto &"",cn   
            if not t.Eof then
                s_nome_produto =  t("descricao") 
                s_nome_produto_html = t("descricao") 
                
            else
			    s_nome_produto = produto_descricao(cod_fabricante, cod_produto)
			    s_nome_produto_html = produto_formata_descricao_em_html(produto_descricao_html(cod_fabricante, cod_produto))
            end if
            if t.State <> 0 then t.Close
		else
			s_nome_produto = ""
			s_nome_produto_html = ""
			end if
            
		end if

		if blnSaidaExcel then
			Response.ContentType = "application/vnd.ms-excel"
			Response.AddHeader "Content-Disposition", "attachment; filename=EstoqueEcommerce_" & formata_data_yyyymmdd(Now) & "_" & formata_hora_hhnnss(Now) & ".xls"
			Response.Write "<h2>Relatório de Estoque (E-Commerce) </h2>"

			select case rb_estoque
				case ID_ESTOQUE_VENDA:			s = "VENDA"
                case VENDA_SHOW_ROOM:           s = "VENDA + SHOW-ROOM" 
				case else						s = ""
				end select
			Response.Write "Estoque de Interesse: " & s
			Response.Write "<br>"

			select case rb_detalhe
				case "INTERMEDIARIO":	s = "INTERMEDIÁRIO (VALOR MÉDIO)"
				case else				s = ""
				end select
			Response.Write "Tipo de Detalhamento: " & s
			Response.Write "<br>"

			s = ""
			if cod_fabricante <> ""  then 
				s = cod_fabricante
'				if (s<>"") And (s_nome_fabricante<>"") and Ubound(v_fabricantes) < 1 then s = s & " - " & s_nome_fabricante
				end if
			if s <> "" then
				Response.Write "Fabricante: " & s
				Response.Write "<br>"
				end if

			s = c_empresa
			if s = "" then 
				s = "todas"
			else
				s = obtem_apelido_empresa_NFe_emitente(c_empresa)
				end if
			Response.Write "Empresa: " & s
			Response.Write "<br>"

			s = ""
			if cod_produto <> "" then
				s = cod_produto
				if (s<>"") And (s_nome_produto_html<>"") then s = s & " - " & s_nome_produto_html
				end if
			if s <> "" then
				Response.Write "Produto: " & s
				Response.Write "<br>"
				end if
            
            s = ""
            if c_loja <> "" then
                s = c_loja
            end if
            if s <> "" then
				Response.Write "Loja: " & s
				Response.Write "<br>"
			end if

            s = ""
            if ckb_normais <> "" Or ckb_compostos <> "" then
               if (ckb_normais<>"")  then 
                    s = s & ckb_normais
                    if (ckb_compostos<>"") then s = s & ", " & ckb_compostos
                    else if (ckb_compostos<>"") then s = s & ckb_compostos                
                end if  
            end if
            if s <> "" then
				Response.Write "Opções de Exportação: " & s
				Response.Write "<br>"
			end if
			
			s = "Emissão: " & formata_data_hora(Now)
			Response.Write s
			Response.Write "<br><br>"
			
			if rb_estoque = ID_ESTOQUE_VENDA then
				select case rb_detalhe					
					case "INTERMEDIARIO"
						consulta_estoque_venda_detalhe_intermediario					
					end select
            elseif rb_estoque = VENDA_SHOW_ROOM then
                select case rb_detalhe
                    case "SINTETICO"
                        consulta_estoque_venda_show_room_detalhe_sintetico
                    case "INTERMEDIARIO"
                        consulta_estoque_venda_show_room_detalhe_intermediario
            end select
			else
				select case rb_detalhe
					
					case "INTERMEDIARIO"
						consulta_estoque_detalhe_intermediario
					end select
				end if

			Response.End
			end if





' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________

' ________________________________________________
' CONSULTA ESTOQUE VENDA DETALHE INTERMEDIARIO, PRODUTOS NORMAIS E COMPOSTOS
'

sub consulta_estoque_venda_detalhe_intermediario
   
dim r
dim s, s_aux, s_sql, s_nbsp, x, cab_table, cab, fabricante_a
dim n_reg, n_saldo_total, n_saldo_parcial, qtde_fabricantes
dim vl, vl_total_geral, vl_sub_total
cont = 0
     if ckb_normais <> "" then
        if c_loja <> "" then
	      s_sql =" SELECT t_ESTOQUE_ITEM.fabricante, t_ESTOQUE_ITEM.produto, descricao, descricao_html" & _
			    ", Sum(qtde-qtde_utilizada) AS saldo, Sum((qtde-qtde_utilizada)*t_ESTOQUE_ITEM.vl_custo2) AS preco_total,preco_lista" & _
			    " FROM t_ESTOQUE_ITEM LEFT JOIN t_PRODUTO ON" & _
			     " ((t_ESTOQUE_ITEM.fabricante=t_PRODUTO.fabricante) AND (t_ESTOQUE_ITEM.produto=t_PRODUTO.produto))" & _
                 " LEFT JOIN t_ESTOQUE ON (t_ESTOQUE_ITEM.id_estoque=t_ESTOQUE.id_estoque)" & _
                 " INNER JOIN t_PRODUTO_LOJA ON (t_ESTOQUE_ITEM.fabricante = t_PRODUTO_LOJA.fabricante) AND (t_ESTOQUE_ITEM.produto = t_PRODUTO_LOJA.produto) " & _
			     " WHERE ((qtde-qtde_utilizada) > 0)" & _
			     " AND ( t_ESTOQUE_ITEM.produto NOT IN (SELECT DISTINCT produto_composto AS produto FROM t_EC_PRODUTO_COMPOSTO))" & _
			     " AND (t_ESTOQUE_ITEM.produto NOT IN (SELECT DISTINCT produto_item AS produto FROM t_EC_PRODUTO_COMPOSTO_ITEM))" & _
                 " AND (t_PRODUTO_LOJA.loja='" & c_loja & "')"
            if cod_fabricante <> "" then
	            s_sql = s_sql & "AND"
	            s_sql = s_sql & "(" & s_where_normais & ")"
            end if
            if c_empresa <> "" then
                s_sql = s_sql & " AND (t_ESTOQUE.id_nfe_emitente='" & c_empresa & "')"
            end if
            if cod_produto <> "" then
		        s_sql = s_sql & " AND (t_ESTOQUE_ITEM.produto='" & cod_produto & "')"
	        end if
	        s_sql = s_sql & " GROUP BY t_ESTOQUE_ITEM.fabricante, t_ESTOQUE_ITEM.produto, descricao, descricao_html,preco_lista" & _
					    " ORDER BY t_ESTOQUE_ITEM.fabricante, t_ESTOQUE_ITEM.produto, descricao, descricao_html"
        else
            s_sql =" SELECT t_ESTOQUE_ITEM.fabricante, t_ESTOQUE_ITEM.produto, descricao, descricao_html" & _
			    ", Sum(qtde-qtde_utilizada) AS saldo, Sum((qtde-qtde_utilizada)*t_ESTOQUE_ITEM.vl_custo2) AS preco_total" & _
			    " FROM t_ESTOQUE_ITEM LEFT JOIN t_PRODUTO ON" & _
			     " ((t_ESTOQUE_ITEM.fabricante=t_PRODUTO.fabricante) AND (t_ESTOQUE_ITEM.produto=t_PRODUTO.produto))" & _            
                 " LEFT JOIN t_ESTOQUE ON (t_ESTOQUE_ITEM.id_estoque=t_ESTOQUE.id_estoque)" & _     
			     " WHERE ((qtde-qtde_utilizada) > 0)" & _
			     " AND ( t_ESTOQUE_ITEM.produto NOT IN (SELECT DISTINCT produto_composto AS produto FROM t_EC_PRODUTO_COMPOSTO))" & _
			     " AND (t_ESTOQUE_ITEM.produto NOT IN (SELECT DISTINCT produto_item AS produto FROM t_EC_PRODUTO_COMPOSTO_ITEM))"                 
            if cod_fabricante <> "" then
	            s_sql = s_sql & "AND"
	            s_sql = s_sql & "(" & s_where_normais & ")"
            end if
            if c_empresa <> "" then
                s_sql = s_sql & " AND (t_ESTOQUE.id_nfe_emitente='" & c_empresa & "')"
            end if
            if cod_produto <> "" then
		        s_sql = s_sql & " AND (t_ESTOQUE_ITEM.produto='" & cod_produto & "')"
	        end if
	            s_sql = s_sql & " GROUP BY t_ESTOQUE_ITEM.fabricante, t_ESTOQUE_ITEM.produto, descricao, descricao_html " & _
					        " ORDER BY t_ESTOQUE_ITEM.fabricante, t_ESTOQUE_ITEM.produto, descricao, descricao_html"
        end if

    
        set r = cn.execute(s_sql)
        
	    do while Not r.Eof

            redim preserve v_fabricante(cont)
            redim preserve v_codigo(cont)
            redim preserve v_descricao(cont)
            redim preserve v_qtde(cont)
            redim preserve v_valor(cont)
            redim preserve v_preco_lista(cont)               
            v_fabricante(cont) =  r("fabricante")
            v_codigo(cont)  =  r("produto")
            v_descricao(cont) = r("descricao")
            v_qtde(cont) = r("saldo")
            v_valor(cont) = r("preco_total")
            if c_loja <> "" then 
                v_preco_lista(cont) = r("preco_lista")
            else
                v_preco_lista(cont) = 0
            end if      
            cont = cont + 1
            r.MoveNext
        loop
        if r.State <> 0 then r.Close
        set r=nothing       
     end if

     if ckb_compostos <> "" then
                
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
			vl_custo2_composto = 0
            preco_lista = 0
			qtde_estoque_venda_composto = -1
			
			s_sql =         "SELECT " & _
						        " fabricante_item," & _
						        " produto_item," & _
						        " qtde" & _
			                    " FROM t_EC_PRODUTO_COMPOSTO_ITEM"  & _
                                " WHERE" & _
						        " (fabricante_composto = '" & Trim("" & r("fabricante_composto")) & "')" & _
						        " AND (produto_composto = '" & Trim("" & r("produto_composto")) & "')" 

            if cod_fabricante <> "" then
	            s_sql = s_sql & " AND "
	            s_sql = s_sql & " (" & s_where_compostos & ") "
            end if

            if cod_produto <> "" then
		        s_sql = s_sql & " AND (produto_composto='" & cod_produto & "')"
	        end if 
			s_sql = s_sql & " ORDER BY" & _
						    " fabricante_item," & _
						    " produto_item"

			if tPCI.State <> 0 then tPCI.Close
			tPCI.Open s_sql, cn
			do while Not tPCI.Eof
            
				s_sql = " SELECT" & _
							" tP.fabricante," & _
							" tP.produto," & _
							" Coalesce((SELECT Sum((tEI.qtde-qtde_utilizada)* vl_custo2) AS saldo FROM t_ESTOQUE_ITEM tEI INNER JOIN t_ESTOQUE tE ON (tEI.id_estoque=tE.id_estoque) WHERE (tEI.fabricante=tP.fabricante) AND (tEI.produto=tP.produto)"
                if c_empresa <> "" then s_sql = s_sql & " AND (tE.id_nfe_emitente = '" & c_empresa & "')"
                s_sql = s_sql & "), 0) AS vl_custo2," & _
							" Coalesce((SELECT Sum(qtde-qtde_utilizada) FROM t_ESTOQUE_ITEM tEI INNER JOIN t_ESTOQUE tE ON (tEI.id_estoque=tE.id_estoque) WHERE (tEI.fabricante=tP.fabricante) AND (tEI.produto=tP.produto) AND ((qtde-qtde_utilizada)>0)"
                if c_empresa <> "" then s_sql = s_sql & " AND (tE.id_nfe_emitente = '" & c_empresa & "')"    
                s_sql = s_sql & "), 0) AS qtde_estoque_venda," & _
                            " t_PRODUTO_LOJA.preco_lista " & _                    
						" FROM t_PRODUTO tPL" & _
							" INNER JOIN t_PRODUTO tP ON (tPL.fabricante = tP.fabricante) AND (tPL.produto = tP.produto)" & _
                            " INNER JOIN t_PRODUTO_LOJA on (tPL.fabricante = t_PRODUTO_LOJA.fabricante) AND (tPL.produto = t_PRODUTO_LOJA.produto)" & _                                            
						" WHERE " & _
                        " (tP.fabricante = '" & Trim("" & tPCI("fabricante_item")) & "')" & _
                       	" AND (tP.produto = '" & Trim("" & tPCI("produto_item")) & "') "   
                if c_loja <> "" then                         								
				      s_sql = s_sql + " AND (t_PRODUTO_LOJA.loja = '" & c_loja &"')"
                end if  

				s_sql = " SELECT " & _
							"*" & _
						" FROM (" & s_sql & ") t" & _
						" WHERE" & _
							" (vl_custo2 > 0) " & _
						" ORDER BY" & _
							" fabricante," & _
							" produto"
				if t.State <> 0 then t.Close
				t.Open s_sql, cn
				if t.Eof then
				    blnPularProdutoComposto = true
              
				else
                    if c_loja <> "" then
                        preco_lista = preco_lista + t("preco_lista") 
                    else
                        preco_lista = 0 
                    end if                   
					vl_custo2_composto = vl_custo2_composto + (tPCI("qtde") * (t("vl_custo2") / t("qtde_estoque_venda")))
					qtde_estoque_venda_aux = t("qtde_estoque_venda") \ tPCI("qtde")
					if qtde_estoque_venda_composto = -1 then
						qtde_estoque_venda_composto = qtde_estoque_venda_aux
					else
						if qtde_estoque_venda_aux < qtde_estoque_venda_composto then
							qtde_estoque_venda_composto = qtde_estoque_venda_aux
						end if
					end if   
				end if

				    if blnPularProdutoComposto then exit do
				    tPCI.MoveNext
				loop
			
			if qtde_estoque_venda_composto > 0 then
               if Not blnPularProdutoComposto then
			     '> CÓDIGO DO PRODUTO
                    redim preserve v_codigo(cont)
				    v_codigo(cont) =  Trim("" & r("produto_composto")) 

			     '> VALOR CUSTO
			    '	EXPORTAR VALOR UTILIZANDO '.' COMO SEPARADOR DECIMAL
                    redim preserve v_valor(cont)
				    v_valor(cont) = formata_moeda(vl_custo2_composto * qtde_estoque_venda_composto )
				
			     '> SALDO ESTOQUE
                    redim preserve v_qtde(cont)
				    v_qtde(cont) = Cstr(qtde_estoque_venda_composto)			
			     '> DESCRIÇÃO
                    redim preserve v_descricao(cont)
				    v_descricao(cont) = Trim("" & r("descricao"))			
			     '> NOME DO FABRICANTE
                    redim preserve v_fabricante(cont)
				    v_fabricante(cont) = Trim("" & r("fabricante_composto"))				                   
                 '> PREÇO LISTA
                    redim preserve v_preco_lista(cont)
				    v_preco_lista(cont) = preco_lista				
                    cont = cont + 1
		        end if
			end if

			r.MoveNext
			loop
           
      end if
		

        redim vRelat(0)
	    set vRelat(0) = New cl_SEIS_COLUNAS
	    with vRelat(0)
		    .c1 = ""
		    .c2 = ""
		    .c3 = ""
		    .c4 = ""
		    .c5 = ""
            .c6 = ""
		end with
    if v_codigo(Ubound(v_codigo)) <> "" then
        for cont = 0 to Ubound(v_codigo)
            if Trim(vRelat(ubound(vRelat)).c1) <> "" then
				redim preserve vRelat(ubound(vRelat)+1)
				set vRelat(ubound(vRelat)) = New cl_SEIS_COLUNAS
			end if
			with vRelat(ubound(vRelat))
				.c1 =  v_codigo(cont)
                .c2 =  v_fabricante(cont)
                .c3 =  v_descricao(cont)
                .c4 =  v_qtde(cont)
                .c5 =  v_valor(cont)
                .c6 =  v_preco_lista(cont)
			end with
        next
    end if

        ordena_cl_seis_colunas vRelat, 0, Ubound(vRelat)
            if c_loja <> "" then
                classe_cab="MD MB"
                classe_td= "MDB"
                colspan = 6 
            else              
                classe_cab="MB"
                classe_td= "MB"    
                colspan = 5
            end if
                  ' CABEÇALHO
	        cab_table = "<TABLE class='Q' cellSpacing=0>" & chr(13)
        '	Mantendo a padronização com as demais opções de emissão do relatório em Excel (sem bordas no cabeçalho)
	        if blnSaidaExcel then
		        cab = "	    <TR style='background:azure' NOWRAP>" & chr(13) & _
			          "		    <TD width='75' valign='bottom'><P class='R' style='font-weight:bold;'>PRODUTO</P></TD>" & chr(13) & _
			          "		    <TD width='350' valign='bottom'><P class='R' style='font-weight:bold;'>DESCRIÇÃO</P></TD>" & chr(13) & _
			          "		    <TD width='60' align='right' valign='bottom'><P class='Rd' style='font-weight:bold;'>QTDE</P></TD>" & chr(13) & _
			          "		    <TD width='100' align='right' valign='bottom'><P class='Rd' style='font-weight:bold;'>REFERÊNCIA<br style='mso-data-placement:same-cell;' />UNITÁRIO MÉDIO</P></TD>" & chr(13) & _
			          "		    <TD width='100' align='right' valign='bottom'><P class='Rd' style='font-weight:bold;'>REFERÊNCIA<br style='mso-data-placement:same-cell;' />TOTAL</P></TD>" & chr(13) 
			    if c_loja <> "" then
                    cab = cab & "   <TD width='100' valign='bottom' NOWRAP class='MB'><P class='Rd' style='font-weight:bold;'>PREÇO LISTA</P></TD>" & chr(13) 
                end if
		        cab = cab &"</TR>" & chr(13)
	        else
		        cab = "	    <TR style='background:azure' NOWRAP>" & chr(13) & _
			          "		    <TD width='60' valign='bottom' NOWRAP class='MD MB'><P class='R'>PRODUTO</P></TD>" & chr(13) & _
			          "		    <TD width='274' valign='bottom' class='MD MB'><P class='R'>DESCRIÇÃO</P></TD>" & chr(13) & _
			          "		    <TD width='60' valign='bottom' NOWRAP class='MD MB'><P class='Rd' style='font-weight:bold;'>QTDE</P></TD>" & chr(13) & _
			          "		    <TD width='100' valign='bottom' NOWRAP class='MD MB'><P class='Rd' style='font-weight:bold;'>REFERÊNCIA UNITÁRIO MÉDIO</P></TD>" & chr(13) & _
			          "		    <TD width='100' valign='bottom' NOWRAP class='" & classe_cab & "'><P class='Rd' style='font-weight:bold;'>REFERÊNCIA TOTAL</P></TD>" & chr(13)
            if c_loja <> "" then
                cab = cab & "   <TD width='100' valign='bottom' NOWRAP class='MB'><P class='Rd' style='font-weight:bold;'>PREÇO LISTA</P></TD>" & chr(13) 
            end if
		        cab = cab &"</TR>" & chr(13)
			          
		    end if
	
	        x = cab_table & cab
	        n_reg = 0
	        n_saldo_total = 0
	        n_saldo_parcial = 0
	        vl_total_geral = 0
	        vl_sub_total = 0
	        qtde_fabricantes = 0
	        fabricante_a = "XXXXX"
		    
        if vRelat(ubound(vRelat)).c1 <> "" then	
	        for cont=Lbound(vRelat) to Ubound(vRelat)
                
	        '	MUDOU DE FABRICANTE?
		        if vRelat(cont).c2 <> fabricante_a then
		        '	SUB-TOTAL POR FORNECEDOR
			        if n_reg > 0 then
				        x = x & "	<TR NOWRAP>" & chr(13) & _
						        "		<TD class='MB' colspan='2' align='right' valign='bottom'><P class='Cd' style='font-weight:bold;'>Total:</P></TD>" & chr(13) & _
						        "		<TD class='MB' valign='bottom'><P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(n_saldo_parcial) & "</P></TD>" & chr(13) & _
						        "		<TD class='MB' valign='bottom'>&nbsp;</TD>" & chr(13) & _
						        "		<TD class='MB' valign='bottom'><P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl_sub_total) & "</P></TD>" & chr(13) 
                                
                        if c_loja <> "" then
                            x = x &"    <TD class='MB' valign='bottom'>&nbsp;</TD>" 
                        end if
						x = x & "	</TR>" & chr(13)
				
				        if blnSaidaExcel then
					        x = x & _
						        "	<TR NOWRAP>" & chr(13) & _
						        "		<TD colspan='5' class='MB'>&nbsp;</TD>" & _
						        "	</TR>" & chr(13) & _
						        "	<TR NOWRAP>" & chr(13) & _
						        "		<TD colspan='5'>&nbsp;</TD>" & _
						        "	</TR>" & chr(13)
				        else
					        x = x & _
						        "	<TR NOWRAP>" & chr(13) & _
						        "		<TD colspan='" & colspan &"' class='MB'>&nbsp;</TD>" & _
						        "	</TR>" & chr(13)
					        end if
				        end if
			        qtde_fabricantes = qtde_fabricantes + 1
			        fabricante_a = vRelat(cont).c2
			        s = vRelat(cont).c2
			        s_aux = ucase(x_fabricante(s))
			        if (s<>"") And (s_aux<>"") then s = s & " - "
			            s = s & s_aux
			        x = x & "	<TR NOWRAP>" & chr(13) & _
					        "		<TD class='MB' align='center' valign='bottom' colspan='" & colspan & "' style='background: honeydew'><P class='C' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & s & "</P></TD>" & chr(13) & _
					        "	</TR>" & chr(13)
			        n_saldo_parcial = 0
			        vl_sub_total = 0
			        end if
                
	          ' CONTAGEM
		        n_reg = n_reg + 1

		        x = x & "	<TR NOWRAP>" & chr(13)

	         '> PRODUTO
		        x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='C' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & vRelat(cont).c1 & "</P></TD>" & chr(13)

	         '> DESCRIÇÃO
		        x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='C' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & produto_formata_descricao_em_html(vRelat(cont).c3) & "</P></TD>" & chr(13)

	         '> SALDO
		        x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(vRelat(cont).c4) & "</P></TD>" & chr(13)
	
	         '> CUSTO DE ENTRADA UNITÁRIO (MÉDIA)
		        if vRelat(cont).c4 = 0 then vl = 0 else vl = vRelat(cont).c5/vRelat(cont).c4
		        x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl) & "</P></TD>" & chr(13)

	         '> CUSTO DE ENTRADA TOTAL
		        vl = vRelat(cont).c5
		        x = x & "		<TD class='" & classe_td & "' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl) & "</P></TD>" & chr(13)
		     '> PRECO LISTA POR LOJA
                if c_loja <> "" then                  		            
		            x = x & "		<TD class='MB' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vRelat(cont).c6) & "</P></TD>" & chr(13)
		        end if    
                vl_total_geral = vl_total_geral + vl
		        vl_sub_total = vl_sub_total + vl
		
		        n_saldo_total = n_saldo_total + vRelat(cont).c4
		        n_saldo_parcial = n_saldo_parcial + vRelat(cont).c4
		
		        x = x & "	</TR>" & chr(13)

		        if (n_reg mod 100) = 0 then
			        Response.Write x
			        x = ""
			        end if			
		        
		    next
          end if
        
        
        ' MOSTRA TOTAL
	    if n_reg <> 0 then 
	    '	TOTAL DO ÚLTIMO FORNECEDOR
		    x = x & "	<TR NOWRAP>" & chr(13) & _
						    "		<TD colspan='2' align='right' valign='bottom'><P class='Cd' style='font-weight:bold;'>Total:</P></TD>" & chr(13) & _
						    "		<TD valign='bottom'><P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(n_saldo_parcial) & "</P></TD>" & chr(13) & _
						    "		<TD>&nbsp;</TD>" & chr(13) & _
						    "		<TD valign='bottom'><P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl_sub_total) & "</P></TD>" & chr(13) & _
				    "	</TR>" & chr(13)

		    if qtde_fabricantes > 1 then
		    '	TOTAL GERAL
			    x = x & "	<TR NOWRAP><TD COLSPAN='" & colspan & "' class='MC'>&nbsp;</TD></TR>" & chr(13) & _
					    "	<TR NOWRAP style='background: #FFFFDD'>" & chr(13) & _
					    "		<TD COLSPAN='2' align='right' valign='bottom' NOWRAP class='MC'><P class='Cd' style='font-weight:bold;'>TOTAL GERAL:</P></TD>" & chr(13) & _
					    "		<TD NOWRAP class='MC' valign='bottom'><P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(n_saldo_total) & "</P></TD>" & chr(13) & _
					    "		<TD class='MC'>&nbsp;</TD>" & chr(13) & _
					    "		<TD NOWRAP class='MC' valign='bottom'><P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl_total_geral) & "</P></TD>" & chr(13) 
                if c_loja <> "" then
                  x = x & "		<TD class='MC'>&nbsp;</TD>" & chr(13) 
                end if
				x = x &"	</TR>" & chr(13)
					    
			    end if
		    end if

        ' MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	    if n_reg = 0 then
		    x = cab_table & cab & _
			    "	<TR NOWRAP>" & chr(13) & _
			    "		<TD colspan='" & colspan & "' align='center'><P class='ALERTA'>NENHUM PRODUTO DO ESTOQUE SATISFAZ AS CONDIÇÕES ESPECIFICADAS</P></TD>" & chr(13) & _
			    "	</TR>" & chr(13)
		    end if

        ' FECHA TABELA
	    x = x & "</TABLE>" & chr(13)
	
	    Response.write x
end sub

' ________________________________________________________
' CONSULTA ESTOQUE VENDA + SHOW-ROOM DETALHE INTERMEDIARIO
'
sub consulta_estoque_venda_show_room_detalhe_intermediario
dim r
dim s, s_aux, s_sql, s_nbsp, x, cab_table, cab, fabricante_a
dim n_reg, n_saldo_total, n_saldo_parcial, qtde_fabricantes
dim vl, vl_total_geral, vl_sub_total
cont = 0
if ckb_normais <> "" then
    if c_loja <> "" then
	    s_sql = " SELECT fabricante, produto, descricao, descricao_html, SUM(saldo) AS saldo, SUM(preco_total) AS preco_total,preco_lista FROM ( " & _
                   "  SELECT fabricante, produto, descricao, descricao_html, SUM(saldo) AS saldo, preco_total,preco_lista FROM (SELECT SUM(t_ESTOQUE_ITEM.qtde-t_ESTOQUE_ITEM.qtde_utilizada) AS saldo, " & _
                    " t_ESTOQUE_ITEM.fabricante, t_ESTOQUE_ITEM.produto, t_PRODUTO.descricao, t_PRODUTO.descricao_html, Sum((qtde-qtde_utilizada)*t_ESTOQUE_ITEM.vl_custo2) AS preco_total,t_PRODUTO_LOJA.preco_lista as preco_lista " & _
                    " FROM t_ESTOQUE_ITEM LEFT JOIN t_PRODUTO ON ((t_ESTOQUE_ITEM.fabricante=t_PRODUTO.fabricante) AND (t_ESTOQUE_ITEM.produto=t_PRODUTO.produto)) " & _
                    " LEFT JOIN t_ESTOQUE ON (t_ESTOQUE_ITEM.id_estoque=t_ESTOQUE.id_estoque)" & _
                    " INNER JOIN t_PRODUTO_LOJA ON (t_ESTOQUE_ITEM.fabricante = t_PRODUTO_LOJA.fabricante) AND (t_ESTOQUE_ITEM.produto = t_PRODUTO_LOJA.produto) " & _
                    " WHERE ((t_ESTOQUE_ITEM.qtde-t_ESTOQUE_ITEM.qtde_utilizada) > 0) " & _
                    " AND ( t_ESTOQUE_ITEM.produto NOT IN (SELECT DISTINCT produto_composto AS produto FROM t_EC_PRODUTO_COMPOSTO)) " & _
				    " AND (t_ESTOQUE_ITEM.produto NOT IN (SELECT DISTINCT produto_item AS produto FROM t_EC_PRODUTO_COMPOSTO_ITEM)) " & _
                    " AND (t_PRODUTO_LOJA.loja='" & c_loja & "')"

        if cod_fabricante <> "" then
	        s_sql = s_sql & " AND "
	        s_sql = s_sql & " (" & s_where_normais & ") "
        end if           
                
        if c_empresa <> "" then
                s_sql = s_sql & " AND (t_ESTOQUE.id_nfe_emitente='" & c_empresa & "')"
        end if

        if cod_produto <> "" then
		    s_sql = s_sql & " AND (t_ESTOQUE_ITEM.produto='" & cod_produto & "')"
	    end if

        s_sql = s_sql & " GROUP BY t_ESTOQUE_ITEM.fabricante, t_ESTOQUE_ITEM.produto, descricao, descricao_html,preco_lista " & _
                        " UNION ALL SELECT SUM(t_ESTOQUE_MOVIMENTO.qtde) AS saldo, " & _
                        " t_ESTOQUE_MOVIMENTO.fabricante, t_ESTOQUE_MOVIMENTO.produto, t_PRODUTO.descricao, t_PRODUTO.descricao_html, Sum(t_ESTOQUE_MOVIMENTO.qtde*t_ESTOQUE_ITEM.vl_custo2) AS preco_total,t_PRODUTO_LOJA.preco_lista as preco_lista " & _
                        " FROM t_ESTOQUE_MOVIMENTO INNER JOIN t_ESTOQUE_ITEM ON ((t_ESTOQUE_MOVIMENTO.id_estoque=t_ESTOQUE_ITEM.id_estoque) " & _
                        " AND (t_ESTOQUE_MOVIMENTO.fabricante=t_ESTOQUE_ITEM.fabricante) AND (t_ESTOQUE_MOVIMENTO.produto=t_ESTOQUE_ITEM.produto)) " & _
                        " LEFT JOIN t_PRODUTO ON ((t_ESTOQUE_MOVIMENTO.fabricante=t_PRODUTO.fabricante) AND (t_ESTOQUE_MOVIMENTO.produto=t_PRODUTO.produto)) " & _
                        " LEFT JOIN t_ESTOQUE ON (t_ESTOQUE_ITEM.id_estoque=t_ESTOQUE.id_estoque)" & _
                        " INNER JOIN t_PRODUTO_LOJA ON (t_ESTOQUE_MOVIMENTO.fabricante = t_PRODUTO_LOJA.fabricante) AND (t_ESTOQUE_MOVIMENTO.produto = t_PRODUTO_LOJA.produto)  " & _
                        " WHERE (anulado_status=0 AND estoque='SHR') " & _
                        " AND ( t_ESTOQUE_ITEM.produto NOT IN (SELECT DISTINCT produto_composto AS produto FROM t_EC_PRODUTO_COMPOSTO)) " & _
				        " AND (t_ESTOQUE_ITEM.produto NOT IN (SELECT DISTINCT produto_item AS produto FROM t_EC_PRODUTO_COMPOSTO_ITEM)) " & _
                         " AND (t_PRODUTO_LOJA.loja='" & c_loja & "')"
                         
        if cod_fabricante <> "" then     
	        s_sql = s_sql & "AND"
	        s_sql = s_sql & "(" & s_where_normais2 & ")"
        end if 

        if c_empresa <> "" then
                s_sql = s_sql & " AND (t_ESTOQUE.id_nfe_emitente='" & c_empresa & "')"
        end if

        if cod_produto <> "" then
		    s_sql = s_sql & " AND (t_ESTOQUE_MOVIMENTO.produto='" & cod_produto & "')"
	    end if

        s_sql = s_sql & " GROUP BY t_ESTOQUE_MOVIMENTO.fabricante, t_ESTOQUE_MOVIMENTO.produto, descricao, descricao_html,preco_lista) tbl " & _
                        " GROUP BY tbl.fabricante, tbl.produto, tbl.descricao, tbl.descricao_html,tbl.preco_total,tbl.preco_lista " & _
                        " ) tbl2 "

        s_sql = s_sql & " GROUP BY tbl2.fabricante, tbl2.produto, tbl2.descricao, tbl2.descricao_html,tbl2.preco_lista " & _
                        " ORDER BY tbl2.fabricante, tbl2.produto, tbl2.descricao, tbl2.descricao_html "
    else
        s_sql = " SELECT fabricante, produto, descricao, descricao_html, SUM(saldo) AS saldo, SUM(preco_total) AS preco_total FROM ( " & _
                "  SELECT fabricante, produto, descricao, descricao_html, SUM(saldo) AS saldo, preco_total FROM (SELECT SUM(t_ESTOQUE_ITEM.qtde-t_ESTOQUE_ITEM.qtde_utilizada) AS saldo, " & _
                " t_ESTOQUE_ITEM.fabricante, t_ESTOQUE_ITEM.produto, t_PRODUTO.descricao, t_PRODUTO.descricao_html, Sum((qtde-qtde_utilizada)*t_ESTOQUE_ITEM.vl_custo2) AS preco_total " & _
                " FROM t_ESTOQUE_ITEM LEFT JOIN t_PRODUTO ON ((t_ESTOQUE_ITEM.fabricante=t_PRODUTO.fabricante) AND (t_ESTOQUE_ITEM.produto=t_PRODUTO.produto)) " & _  
                " LEFT JOIN t_ESTOQUE ON (t_ESTOQUE_ITEM.id_estoque=t_ESTOQUE.id_estoque)" & _              
                " WHERE ((t_ESTOQUE_ITEM.qtde-t_ESTOQUE_ITEM.qtde_utilizada) > 0) " & _
                " AND ( t_ESTOQUE_ITEM.produto NOT IN (SELECT DISTINCT produto_composto AS produto FROM t_EC_PRODUTO_COMPOSTO)) " & _
				" AND (t_ESTOQUE_ITEM.produto NOT IN (SELECT DISTINCT produto_item AS produto FROM t_EC_PRODUTO_COMPOSTO_ITEM)) " 

        if cod_fabricante <> "" then
	        s_sql = s_sql & " AND "
	        s_sql = s_sql & " (" & s_where_normais & ") "
        end if           
                
        if c_empresa <> "" then
                s_sql = s_sql & " AND (t_ESTOQUE.id_nfe_emitente='" & c_empresa & "')"
        end if

        if cod_produto <> "" then
		    s_sql = s_sql & " AND (t_ESTOQUE_ITEM.produto='" & cod_produto & "')"
	    end if

        s_sql = s_sql & " GROUP BY t_ESTOQUE_ITEM.fabricante, t_ESTOQUE_ITEM.produto, descricao, descricao_html " & _
                        " UNION ALL SELECT SUM(t_ESTOQUE_MOVIMENTO.qtde) AS saldo, " & _
                        " t_ESTOQUE_MOVIMENTO.fabricante, t_ESTOQUE_MOVIMENTO.produto, t_PRODUTO.descricao, t_PRODUTO.descricao_html, Sum(t_ESTOQUE_MOVIMENTO.qtde*t_ESTOQUE_ITEM.vl_custo2) AS preco_total" & _
                        " FROM t_ESTOQUE_MOVIMENTO INNER JOIN t_ESTOQUE_ITEM ON ((t_ESTOQUE_MOVIMENTO.id_estoque=t_ESTOQUE_ITEM.id_estoque) " & _
                        " AND (t_ESTOQUE_MOVIMENTO.fabricante=t_ESTOQUE_ITEM.fabricante) AND (t_ESTOQUE_MOVIMENTO.produto=t_ESTOQUE_ITEM.produto)) " & _
                        " LEFT JOIN t_PRODUTO ON ((t_ESTOQUE_MOVIMENTO.fabricante=t_PRODUTO.fabricante) AND (t_ESTOQUE_MOVIMENTO.produto=t_PRODUTO.produto)) " & _
                        " LEFT JOIN t_ESTOQUE ON (t_ESTOQUE_ITEM.id_estoque=t_ESTOQUE.id_estoque)" & _
                        " WHERE (anulado_status=0 AND estoque='SHR') " & _
                        " AND ( t_ESTOQUE_ITEM.produto NOT IN (SELECT DISTINCT produto_composto AS produto FROM t_EC_PRODUTO_COMPOSTO)) " & _
				        " AND (t_ESTOQUE_ITEM.produto NOT IN (SELECT DISTINCT produto_item AS produto FROM t_EC_PRODUTO_COMPOSTO_ITEM)) " 
                
                         
        if cod_fabricante <> "" then     
	        s_sql = s_sql & "AND"
	        s_sql = s_sql & "(" & s_where_normais2 & ")"
        end if 

        if c_empresa <> "" then
                s_sql = s_sql & " AND (t_ESTOQUE.id_nfe_emitente='" & c_empresa & "')"
        end if

        if cod_produto <> "" then
		    s_sql = s_sql & " AND (t_ESTOQUE_MOVIMENTO.produto='" & cod_produto & "')"
	    end if

        s_sql = s_sql & " GROUP BY t_ESTOQUE_MOVIMENTO.fabricante, t_ESTOQUE_MOVIMENTO.produto, descricao, descricao_html) tbl " & _
                        " GROUP BY tbl.fabricante, tbl.produto, tbl.descricao, tbl.descricao_html, tbl.preco_total " & _
                        " ) tbl2 "

        s_sql = s_sql & " GROUP BY tbl2.fabricante, tbl2.produto, tbl2.descricao, tbl2.descricao_html " & _
                        " ORDER BY tbl2.fabricante, tbl2.produto, tbl2.descricao, tbl2.descricao_html "
    end if
        
    set r = cn.execute(s_sql)
	do while Not r.Eof
        redim preserve v_fabricante(cont)
        redim preserve v_codigo(cont)
        redim preserve v_descricao(cont)
        redim preserve v_qtde(cont)
        redim preserve v_valor(cont)
        redim preserve v_preco_lista(cont)             
        v_fabricante(cont) =  r("fabricante")
        v_codigo(cont)  =  r("produto")
        v_descricao(cont) = r("descricao")
        v_qtde(cont) = r("saldo")
        v_valor(cont) = r("preco_total")
        if c_loja <> "" then
            v_preco_lista(cont) = r("preco_lista")
        else
            v_preco_lista(cont) = 0 
        end if    
        cont = cont + 1
        r.MoveNext
    loop
    if r.State <> 0 then r.Close
    set r=nothing 
end if

 if ckb_compostos <> "" then
                
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
			vl_custo2_composto = 0
            preco_lista = 0
			qtde_estoque_venda_composto = -1
			
			s_sql =         "SELECT " & _
						        " fabricante_item," & _
						        " produto_item," & _
						        " qtde" & _                 
			                    " FROM t_EC_PRODUTO_COMPOSTO_ITEM" & _            
                                " WHERE" & _
						        " (fabricante_composto = '" & Trim("" & r("fabricante_composto")) & "')" & _
						        " AND (produto_composto = '" & Trim("" & r("produto_composto")) & "')" 

            if cod_fabricante <> "" then
	            s_sql = s_sql & " AND "
	            s_sql = s_sql & " (" & s_where_compostos & ") "
            end if

            if cod_produto <> "" then
		        s_sql = s_sql & " AND (produto_composto='" & cod_produto & "')"
	        end if 
			s_sql = s_sql & " ORDER BY" & _
						    " fabricante_item," & _
						    " produto_item"
			if tPCI.State <> 0 then tPCI.Close
			tPCI.Open s_sql, cn
			do while Not tPCI.Eof

				s_sql = " SELECT " & _
							" tP.fabricante, " & _
							" tP.produto, " & _
							" Coalesce((SELECT  Sum((tEM.qtde)*tEI.vl_custo2) AS saldo FROM t_ESTOQUE_MOVIMENTO tEM " & _
							" INNER JOIN t_ESTOQUE_ITEM tEI ON ((tEM.id_estoque=tEI.id_estoque)" & _
							" AND (tEM.fabricante=tEI.fabricante) AND (tEM.produto=tEI.produto))" & _
							" LEFT JOIN t_PRODUTO ON ((tEM.fabricante=t_PRODUTO.fabricante) AND (tEM.produto=t_PRODUTO.produto))" & _
                            " INNER JOIN t_ESTOQUE tE ON (tEM.id_estoque=tE.id_estoque)" & _
							" WHERE (tEI.fabricante=tP.fabricante) AND (tEI.produto=tP.produto) AND (anulado_status=0 ) AND (estoque='SHR') AND (tEI.vl_custo2 > 0)"
                if c_empresa <> "" then s_sql = s_sql & " AND (tE.id_nfe_emitente = '" & c_empresa & "')"
                s_sql = s_sql & "), 0) AS vl_custo2_m," & _
							" Coalesce((SELECT  Sum(tEM.qtde) AS saldo FROM t_ESTOQUE_MOVIMENTO tEM " & _
							" INNER JOIN t_ESTOQUE_ITEM tEI ON ((tEM.id_estoque=tEI.id_estoque)" & _
							" AND (tEM.fabricante=tEI.fabricante) AND (tEM.produto=tEI.produto))" & _
							" LEFT JOIN t_PRODUTO ON ((tEM.fabricante=t_PRODUTO.fabricante) AND (tEM.produto=t_PRODUTO.produto))" & _
                            " INNER JOIN t_ESTOQUE tE ON (tEM.id_estoque=tE.id_estoque)" & _
							" WHERE (tEI.fabricante=tP.fabricante) AND (tEI.produto=tP.produto) AND (anulado_status=0 ) AND (estoque='SHR')"
                if c_empresa <> "" then s_sql = s_sql & " AND (tE.id_nfe_emitente = '" & c_empresa & "')"    
                s_sql = s_sql & "), 0) AS qtde_estoque_venda_m," & _
							" Coalesce((SELECT Sum((tEI.qtde-qtde_utilizada)* vl_custo2) AS saldo FROM t_ESTOQUE_ITEM tEI INNER JOIN t_ESTOQUE tE ON (tEI.id_estoque=tE.id_estoque) WHERE (tEI.fabricante=tP.fabricante) AND (tEI.produto=tP.produto) AND (vl_custo2 > 0)"
                if c_empresa <> "" then s_sql = s_sql & " AND (tE.id_nfe_emitente = '" & c_empresa & "')"
                s_sql = s_sql & "), 0) AS vl_custo2_i," & _
							" Coalesce((SELECT Sum(qtde-qtde_utilizada) FROM t_ESTOQUE_ITEM tEI INNER JOIN t_ESTOQUE tE ON (tEI.id_estoque=tE.id_estoque) WHERE (tEI.fabricante=tP.fabricante) AND (tEI.produto=tP.produto) AND ((qtde-qtde_utilizada)>0)"
                if c_empresa <> "" then s_sql = s_sql & " AND (tE.id_nfe_emitente = '" & c_empresa & "')"    
                s_sql = s_sql & "), 0) AS qtde_estoque_venda_i," & _
                            " t_PRODUTO_LOJA.preco_lista" & _
						    " FROM t_PRODUTO tPL" & _
							" INNER JOIN t_PRODUTO tP ON (tPL.fabricante = tP.fabricante) AND (tPL.produto = tP.produto)" & _
                            " INNER JOIN t_PRODUTO_LOJA on (tPL.fabricante = t_PRODUTO_LOJA.fabricante) AND (tPL.produto = t_PRODUTO_LOJA.produto)" & _ 
						    " WHERE (tP.fabricante = '" & Trim("" & tPCI("fabricante_item")) & "')" & _               	 
							" AND (tP.produto = '" & Trim("" & tPCI("produto_item")) & "')" 
                if c_loja <> "" then
                    s_sql = s_sql + " AND (loja = " & c_loja & ")"
                end if

				s_sql = "SELECT " & _
							"*" & _
						" FROM (" & s_sql & ") t" & _
						" WHERE" & _
							" ((vl_custo2_i + vl_custo2_m) > 0)" & _
						" ORDER BY" & _
							" fabricante," & _
							" produto"
				if t.State <> 0 then t.Close
				t.Open s_sql, cn
				if t.Eof then
				blnPularProdutoComposto = true	
				else
                    if c_loja <> "" then
                        preco_lista = preco_lista + t("preco_lista")
                    else
                        preco_lista = 0
                    end if
					vl_custo2_composto = vl_custo2_composto + (tPCI("qtde") * ((t("vl_custo2_m") + t("vl_custo2_i")) / (t("qtde_estoque_venda_m") + t("qtde_estoque_venda_i"))))
					qtde_estoque_venda_aux = (t("qtde_estoque_venda_m") + t("qtde_estoque_venda_i")) \ tPCI("qtde")
					if qtde_estoque_venda_composto = -1 then
						qtde_estoque_venda_composto = qtde_estoque_venda_aux
					else
						if qtde_estoque_venda_aux < qtde_estoque_venda_composto then
							qtde_estoque_venda_composto = qtde_estoque_venda_aux
							end if
						end if
					end if
				
				
				if blnPularProdutoComposto then exit do
				tPCI.MoveNext
				loop
			
			if qtde_estoque_venda_composto > 0 then
                if Not blnPularProdutoComposto then
			     '> CÓDIGO DO PRODUTO
                    redim preserve v_codigo(cont)
				    v_codigo(cont) =  Trim("" & r("produto_composto")) 

			     '> VALOR CUSTO
			    '	EXPORTAR VALOR UTILIZANDO '.' COMO SEPARADOR DECIMAL
                    redim preserve v_valor(cont)
				    v_valor(cont) = formata_moeda(vl_custo2_composto * qtde_estoque_venda_composto )
				
			     '> SALDO ESTOQUE
                    redim preserve v_qtde(cont)
				    v_qtde(cont) = Cstr(qtde_estoque_venda_composto)			
			     '> DESCRIÇÃO
                    redim preserve v_descricao(cont)
				    v_descricao(cont) = Trim("" & r("descricao"))			
			     '> NOME DO FABRICANTE
                    redim preserve v_fabricante(cont)
				    v_fabricante(cont) = Trim("" & r("fabricante_composto"))
                 '> PREÇO LISTA POR LOJA
                    redim preserve v_preco_lista(cont)
				    v_preco_lista(cont) = preco_lista				
                    cont = cont + 1
		        end if
			end if

			r.MoveNext
			loop
           
      end if
    ' ORDENAÇÃO
    redim vRelat(0)
	    set vRelat(0) = New cl_SEIS_COLUNAS
	    with vRelat(0)
		    .c1 = ""
		    .c2 = ""
		    .c3 = ""
		    .c4 = ""
		    .c5 = ""
            .c6 = ""
		end with
        
    if v_codigo(Ubound(v_codigo)) <> "" then
        for cont = 0 to Ubound(v_codigo)
            if Trim(vRelat(ubound(vRelat)).c1) <> "" then
				redim preserve vRelat(ubound(vRelat)+1)
				set vRelat(ubound(vRelat)) = New cl_SEIS_COLUNAS
			end if
			with vRelat(ubound(vRelat))
				.c1 =  v_codigo(cont)
                .c2 =  v_fabricante(cont)
                .c3 =  v_descricao(cont)
                .c4 =  v_qtde(cont)
                .c5 =  v_valor(cont)
                .c6 =  v_preco_lista(cont)
			end with
        next
    end if       
    ordena_cl_seis_colunas vRelat, 0, Ubound(vRelat)


  if c_loja <> "" then
                classe_cab="MD MB"
                classe_td= "MDB"
                colspan = 6 
            else              
                classe_cab="MB"
                classe_td= "MB"    
                colspan = 5
            end if
                  ' CABEÇALHO
	        cab_table = "<TABLE class='Q' cellSpacing=0>" & chr(13)
        '	Mantendo a padronização com as demais opções de emissão do relatório em Excel (sem bordas no cabeçalho)
	        if blnSaidaExcel then
		        cab = "	    <TR style='background:azure' NOWRAP>" & chr(13) & _
			          "		    <TD width='75' valign='bottom'><P class='R' style='font-weight:bold;'>PRODUTO</P></TD>" & chr(13) & _
			          "		    <TD width='350' valign='bottom'><P class='R' style='font-weight:bold;'>DESCRIÇÃO</P></TD>" & chr(13) & _
			          "		    <TD width='60' align='right' valign='bottom'><P class='Rd' style='font-weight:bold;'>QTDE</P></TD>" & chr(13) & _
			          "		    <TD width='100' align='right' valign='bottom'><P class='Rd' style='font-weight:bold;'>REFERÊNCIA<br style='mso-data-placement:same-cell;' />UNITÁRIO MÉDIO</P></TD>" & chr(13) & _
			          "		    <TD width='100' align='right' valign='bottom'><P class='Rd' style='font-weight:bold;'>REFERÊNCIA<br style='mso-data-placement:same-cell;' />TOTAL</P></TD>" & chr(13) 
			    if c_loja <> "" then
                    cab = cab & "    <TD width='100' align='right' valign='bottom'><P class='Rd' style='font-weight:bold;'>PREÇO<br style='mso-data-placement:same-cell;' />LISTA</P></TD>" & chr(13) 
                end if
		        cab = cab &"</TR>" & chr(13)
	        else
		        cab = "	    <TR style='background:azure' NOWRAP>" & chr(13) & _
			          "		    <TD width='60' valign='bottom' NOWRAP class='MD MB'><P class='R'>PRODUTO</P></TD>" & chr(13) & _
			          "		    <TD width='274' valign='bottom' class='MD MB'><P class='R'>DESCRIÇÃO</P></TD>" & chr(13) & _
			          "		    <TD width='60' valign='bottom' NOWRAP class='MD MB'><P class='Rd' style='font-weight:bold;'>QTDE</P></TD>" & chr(13) & _
			          "		    <TD width='100' valign='bottom' NOWRAP class='MD MB'><P class='Rd' style='font-weight:bold;'>REFERÊNCIA UNITÁRIO MÉDIO</P></TD>" & chr(13) & _
			          "		    <TD width='100' valign='bottom' NOWRAP class='" & classe_cab & "'><P class='Rd' style='font-weight:bold;'>REFERÊNCIA TOTAL</P></TD>" & chr(13)
            if c_loja <> "" then
                cab = cab & "   <TD width='100' valign='bottom' NOWRAP class='MB'><P class='Rd' style='font-weight:bold;'>PREÇO LISTA</P></TD>" & chr(13) 
            end if
		        cab = cab &"</TR>" & chr(13)
			          
		    end if
	
	        x = cab_table & cab
	        n_reg = 0
	        n_saldo_total = 0
	        n_saldo_parcial = 0
	        vl_total_geral = 0
	        vl_sub_total = 0
	        qtde_fabricantes = 0
	        fabricante_a = "XXXXX"
		    
        if vRelat(ubound(vRelat)).c1 <> "" then	
	        for cont=Lbound(vRelat) to Ubound(vRelat)
                
	        '	MUDOU DE FABRICANTE?
		        if vRelat(cont).c2 <> fabricante_a then
		        '	SUB-TOTAL POR FORNECEDOR
			        if n_reg > 0 then
				        x = x & "	<TR NOWRAP>" & chr(13) & _
						        "		<TD class='MB' colspan='2' align='right' valign='bottom'><P class='Cd' style='font-weight:bold;'>Total:</P></TD>" & chr(13) & _
						        "		<TD class='MB' valign='bottom'><P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(n_saldo_parcial) & "</P></TD>" & chr(13) & _
						        "		<TD class='MB' valign='bottom'>&nbsp;</TD>" & chr(13) & _
						        "		<TD class='MB' valign='bottom'><P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl_sub_total) & "</P></TD>" & chr(13) 
                                
                        if c_loja <> "" then
                            x = x &"    <TD class='MB' valign='bottom'>&nbsp;</TD>" 
                        end if
						x = x & "	</TR>" & chr(13)
				
				        if blnSaidaExcel then
					        x = x & _
						        "	<TR NOWRAP>" & chr(13) & _
						        "		<TD colspan='5' class='MB'>&nbsp;</TD>" & _
						        "	</TR>" & chr(13) & _
						        "	<TR NOWRAP>" & chr(13) & _
						        "		<TD colspan='5'>&nbsp;</TD>" & _
						        "	</TR>" & chr(13)
				        else
					        x = x & _
						        "	<TR NOWRAP>" & chr(13) & _
						        "		<TD colspan='" & colspan &"' class='MB'>&nbsp;</TD>" & _
						        "	</TR>" & chr(13)
					        end if
				        end if
			        qtde_fabricantes = qtde_fabricantes + 1
			        fabricante_a = vRelat(cont).c2
			        s = vRelat(cont).c2
			        s_aux = ucase(x_fabricante(s))
			        if (s<>"") And (s_aux<>"") then s = s & " - "
			            s = s & s_aux
			        x = x & "	<TR NOWRAP>" & chr(13) & _
					        "		<TD class='MB' align='center' valign='bottom' colspan='" & colspan & "' style='background: honeydew'><P class='C' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & s & "</P></TD>" & chr(13) & _
					        "	</TR>" & chr(13)
			        n_saldo_parcial = 0
			        vl_sub_total = 0
			        end if
                
	          ' CONTAGEM
		        n_reg = n_reg + 1

		        x = x & "	<TR NOWRAP>" & chr(13)

	         '> PRODUTO
		        x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='C' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & vRelat(cont).c1 & "</P></TD>" & chr(13)

	         '> DESCRIÇÃO
		        x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='C' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_TEXTO & chr(34) & ";'>" & s_nbsp & produto_formata_descricao_em_html(vRelat(cont).c3) & "</P></TD>" & chr(13)

	         '> SALDO
		        x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(vRelat(cont).c4) & "</P></TD>" & chr(13)
	
	         '> CUSTO DE ENTRADA UNITÁRIO (MÉDIA)
		        if vRelat(cont).c4 = 0 then vl = 0 else vl = vRelat(cont).c5/vRelat(cont).c4
		        x = x & "		<TD class='MDB' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl) & "</P></TD>" & chr(13)

	         '> CUSTO DE ENTRADA TOTAL
		        vl = vRelat(cont).c5
		        x = x & "		<TD class='" & classe_td & "' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl) & "</P></TD>" & chr(13)
		     '> PRECO LISTA POR LOJA
                if c_loja <> "" then		            
		            x = x & "		<TD class='MB' valign='middle' NOWRAP><P class='Cd' style='mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vRelat(cont).c6) & "</P></TD>" & chr(13)
		        end if    
                vl_total_geral = vl_total_geral + vl
		        vl_sub_total = vl_sub_total + vl
		
		        n_saldo_total = n_saldo_total + vRelat(cont).c4
		        n_saldo_parcial = n_saldo_parcial + vRelat(cont).c4
		
		        x = x & "	</TR>" & chr(13)

		        if (n_reg mod 100) = 0 then
			        Response.Write x
			        x = ""
			        end if					        
		    next
          end if      
        
        ' MOSTRA TOTAL
	    if n_reg <> 0 then 
	    '	TOTAL DO ÚLTIMO FORNECEDOR
		    x = x & "	<TR NOWRAP>" & chr(13) & _
						    "		<TD colspan='2' align='right' valign='bottom'><P class='Cd' style='font-weight:bold;'>Total:</P></TD>" & chr(13) & _
						    "		<TD valign='bottom'><P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(n_saldo_parcial) & "</P></TD>" & chr(13) & _
						    "		<TD>&nbsp;</TD>" & chr(13) & _
						    "		<TD valign='bottom'><P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl_sub_total) & "</P></TD>" & chr(13) & _
				    "	</TR>" & chr(13)

		    if qtde_fabricantes > 1 then
		    '	TOTAL GERAL
			    x = x & "	<TR NOWRAP><TD COLSPAN='" & colspan & "' class='MC'>&nbsp;</TD></TR>" & chr(13) & _
					    "	<TR NOWRAP style='background: #FFFFDD'>" & chr(13) & _
					    "		<TD COLSPAN='2' align='right' valign='bottom' NOWRAP class='MC'><P class='Cd' style='font-weight:bold;'>TOTAL GERAL:</P></TD>" & chr(13) & _
					    "		<TD NOWRAP class='MC' valign='bottom'><P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_INTEIRO & chr(34) & ";'>" & formata_inteiro(n_saldo_total) & "</P></TD>" & chr(13) & _
					    "		<TD class='MC'>&nbsp;</TD>" & chr(13) & _
					    "		<TD NOWRAP class='MC' valign='bottom'><P class='Cd' style='font-weight:bold;mso-number-format:" & chr(34) & MSO_NUMBER_FORMAT_MOEDA & chr(34) & ";'>" & formata_moeda(vl_total_geral) & "</P></TD>" & chr(13) 
                if c_loja <> "" then
                  x = x & "		<TD class='MC'>&nbsp;</TD>" & chr(13) 
                end if
				   x = x &"	</TR>" & chr(13)
			    end if
		    end if

        ' MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	    if n_reg = 0 then
		    x = cab_table & cab & _
			    "	<TR NOWRAP>" & chr(13) & _
			    "		<TD colspan='" & colspan & "' align='center'><P class='ALERTA'>NENHUM PRODUTO DO ESTOQUE SATISFAZ AS CONDIÇÕES ESPECIFICADAS</P></TD>" & chr(13) & _
			    "	</TR>" & chr(13)
		    end if

        ' FECHA TABELA
	    x = x & "</TABLE>" & chr(13)
	
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

<style type="text/css">
#rb_estoque_aux {
	margin: 0pt 2pt 1pt 15pt;
	vertical-align: top;
	}
#rb_detalhe_aux {
	margin: 0pt 2pt 1pt 15pt;
	vertical-align: top;
	}
</style>

<% if rb_detalhe = "SINTETICO" then %>
<style type="text/css">
P.C { font-size:10pt; }
P.Cc { font-size:10pt; }
P.Cd { font-size:10pt; }
P.F { font-size:11pt; }
</style>
<% end if %>


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

<form id="fESTOQ" name="fESTOQ" method="post">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="rb_estoque" id="rb_estoque" value="<%=rb_estoque%>">
<input type="hidden" name="rb_detalhe" id="rb_detalhe" value="<%=rb_detalhe%>">
<input type="hidden" name="c_fabricante" id="c_fabricante" value="<%=c_fabricante%>">
<input type="hidden" name="c_produto" id="c_produto" value="<%=c_produto%>">
<input type="hidden" name="c_produto" id="ckb_normais" value="<%=ckb_normais%>">
<input type="hidden" name="c_produto" id="ckb_compostos" value="<%=ckb_compostos%>">


<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Estoque (E-Commerce)</span>
	<br>
	<%	s = "<span class='N'>Emissão:&nbsp;" & formata_data_hora(Now) & "</span>"
		Response.Write s
	%>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>
<br>

<!--  PARÂMETROS DA CONSULTA  -->
<table class="Qx" cellspacing="0">
<!--  ESTOQUE  -->
	<tr bgcolor="#FFFFFF">
		<% select case rb_estoque
				case ID_ESTOQUE_VENDA:			s = "VENDA"
                case VENDA_SHOW_ROOM:           s = "VENDA + SHOW-ROOM"				
				case else						s = ""
				end select
		%>
	<td class="MT" nowrap><span class="PLTe">Estoque de Interesse</span>
		<br><p class="C" style="width:230px;cursor:default;"><%=s%></p></td>

<!--  TIPO DE DETALHAMENTO  -->
		<% select case rb_detalhe		
			case "INTERMEDIARIO":	s = "INTERMEDIÁRIO (VALOR MÉDIO)"			
			case else				s = ""
			end select
		%>
	<td class="MT" style="border-left:0px;" nowrap><span class="PLTe">Tipo de Detalhamento</span>
		<br><p class="C" style="width:230px;cursor:default;"><%=s%></p></td>
	</tr>

<!--  FABRICANTE  -->
	<% if cod_fabricante <> "" then %>
		<tr bgcolor="#FFFFFF">
		<td class="MDBE"  colspan="2" style="word-wrap:break-word;max-width:500px;margin-left:2pt;"><span class="PLTe">Fabricante</span>
			<%	
             s = ""
             if Ubound(v_fabricantes) > 0 then
                for cont = Lbound(v_fabricantes) to Ubound(v_fabricantes)

                    s = s & v_fabricantes(cont) 
                    i = i + 1
                   if not cont = Ubound(v_fabricantes)  then s = s & ", "
                               
                next
            else
                s = cod_fabricante
				if (s<>"") And (s_nome_fabricante<>"") and Ubound(v_fabricantes) < 1 then s = s 
            end if%>
			<br><span class="PLLe"  style="margin-left:2pt;"><%=s%></span> </td>
		</tr>
	<% end if %>
	
<!--  EMPRESA  -->
	<%
    s = c_empresa
	if s = "" then 
		s = "todas"
	else
		s = obtem_apelido_empresa_NFe_emitente(c_empresa)
		end if
	%>
		<tr bgcolor="#FFFFFF">
		<td class="MDBE" nowrap colspan="2"><span class="PLTe">Empresa</span>
			<br>
			<span class="PLLe" style="width:460px;margin-left:2pt;"><%=s%></span>
			
		</td>
		</tr>

<!--  PRODUTO  -->
	<% if cod_produto <> "" then %>
		<tr bgcolor="#FFFFFF">
		<td class="MDBE" nowrap colspan="2"><span class="PLTe">Produto</span>
			<%	s = cod_produto
				if (s<>"") And (s_nome_produto_html<>"") then s = s & " - " & s_nome_produto_html %>
			<br>
			<span class="PLLe" style="width:460px;margin-left:2pt;"><%=s%></span>
			<%	s = cod_produto
				if (s<>"") And (s_nome_produto<>"") then s = s & " - " & s_nome_produto %>
			<input type="hidden" name="c_produto_aux" id="c_produto_aux" value="<%=s%>">
		</td>
		</tr>
	<% end if %>
    <% if c_loja <> "" then %>
		<tr bgcolor="#FFFFFF">
		<td class="MDBE" nowrap colspan="2"><span class="PLTe">Loja</span>
			<%	s = ""
                s = c_loja                                                                        
            %>     
			<br>
			<span class="PLLe" style="width:460px;margin-left:2pt;"><%=s%></span>
			
		</td>
		</tr>
	<% end if %>
    <% if ckb_normais <> "" Or ckb_compostos <> "" then %>
		<tr bgcolor="#FFFFFF">
		<td class="MDBE" nowrap colspan="2"><span class="PLTe">Opções de Exportação</span>
			<%	s = ""
				if (ckb_normais<>"")  then 
                    s = s & ckb_normais
                    if (ckb_compostos<>"") then s = s & ", " & ckb_compostos
                    else if (ckb_compostos<>"") then s = s & ckb_compostos                
                end if                
            %>     
			<br>
			<span class="PLLe" style="width:460px;margin-left:2pt;"><%=s%></span>
			
		</td>
		</tr>
	<% end if %>
	
</table>

<!--  RELATÓRIO  -->
<br>
<%	
	if rb_estoque = ID_ESTOQUE_VENDA then
		select case rb_detalhe
			
			case "INTERMEDIARIO"
				consulta_estoque_venda_detalhe_intermediario			
			end select
    elseif rb_estoque = VENDA_SHOW_ROOM then
        select case rb_detalhe           
            case "INTERMEDIARIO"
                consulta_estoque_venda_show_room_detalhe_intermediario
        end select
	else
		select case rb_detalhe
			
			case "INTERMEDIARIO"
				consulta_estoque_detalhe_intermediario			
			end select
		end if
%>

<!-- ************   SEPARADOR   ************ -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
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
	if rs.State <> 0 then rs.Close
	set rs = nothing

'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>
