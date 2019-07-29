<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ======================================================
'	  RelEstoqueVendaCmvPvExec.asp
'     ======================================================
'     A página foi renomeada em 24/01/2018, anteriormente chamava-se RelPosicaoEstoqueCmvPvExec.asp
'     Este relatório foi duplicado da Loja p/ a Central, mas como na Central já havia uma página c/ o mesmo nome, optou-se por renomear p/ que este relatório mantivesse o mesmo nome na Loja e na Central.
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

	dim usuario, loja
	usuario = Trim(Session("usuario_atual"))
	loja = Trim(Session("loja_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 
	If (loja = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	CONECTA COM O BANCO DE DADOS
	dim cn, rs, msg_erro,tPCI,t
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
    If Not cria_recordset_otimista(tPCI, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
    If Not cria_recordset_otimista(t, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	
	
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	if Not ( _
			operacao_permitida(OP_LJA_REL_ESTOQUE_VENDA_SINTETICO_CMVPV, s_lista_operacoes_permitidas) Or _
			operacao_permitida(OP_LJA_REL_ESTOQUE_VENDA_INTERMEDIARIO_CMVPV, s_lista_operacoes_permitidas) _
			) then 
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	dim alerta
	dim s, c_fabricante, c_produto, rb_estoque, rb_detalhe, c_empresa
	dim cod_fabricante, cod_produto,cont 
	dim s_nome_fabricante, s_nome_produto, s_nome_produto_html,rb_exportacao
    dim v_fabricante(),v_codigo(),v_descricao(),v_qtde(),v_valor(),v_produtos()
    dim qtde_estoque_venda_aux,n_reg_total,vRelat(),vl_custo2_composto,qtde_estoque_venda_composto,i,blnPularProdutoComposto
    dim s_where_compostos

    redim v_codigo(0)
    cont = 0
	c_fabricante = retorna_so_digitos(Request.Form("c_fabricante"))
	if c_fabricante <> "" then c_fabricante = normaliza_codigo(c_fabricante, TAM_MIN_FABRICANTE)
	c_produto = UCase(Trim(Request.Form("c_produto")))
	rb_estoque = Trim(Request.Form("rb_estoque"))
	rb_detalhe = Trim(Request.Form("rb_detalhe"))
	rb_exportacao = Trim(Request.Form("rb_exportacao"))
    c_empresa = Trim(Request.Form("c_empresa"))

	alerta = ""
	if (c_produto<>"") And (Not IsEAN(c_produto)) then
		if c_fabricante = "" then alerta = "PARA CONSULTAR PELO CÓDIGO INTERNO DE PRODUTO É NECESSÁRIO ESPECIFICAR O FABRICANTE."
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
    
    
	if alerta = "" then
		if cod_fabricante <> "" then
			s_nome_fabricante = fabricante_descricao(cod_fabricante)
            if rb_exportacao = "Compostos" then s_where_compostos = " fabricante_composto ='" & cod_fabricante & "'"
		else
			s_nome_fabricante = ""
			end if
				
        
		if cod_produto <> "" then
			s_nome_produto = produto_descricao(cod_fabricante, cod_produto)
			s_nome_produto_html = produto_formata_descricao_em_html(produto_descricao_html(cod_fabricante, cod_produto))
		else
			s_nome_produto = ""
			s_nome_produto_html = ""
			end if
		end if

    if alerta = "" then
        if rb_exportacao = "Compostos" AND c_produto <> "" then
            s ="SELECT descricao FROM t_EC_PRODUTO_COMPOSTO WHERE produto_composto = " & c_produto &""
            if rs.State <> 0 then rs.Close
			    rs.open s, cn

            if not rs.Eof then
                s_nome_produto =  rs("descricao") 
                s_nome_produto_html = rs("descricao") 
            end if
        end if
    end if

    if alerta = "" then
		call set_default_valor_texto_bd(usuario, "RelEstoqueVendaCmvPvLoja|rb_detalhe", rb_detalhe)
		call set_default_valor_texto_bd(usuario, "RelEstoqueVendaCmvPvLoja|rb_exportacao", rb_exportacao)
		end if	


' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________

' ___________________________________________
' CONSULTA ESTOQUE VENDA DETALHE SINTETICO
' 
sub consulta_estoque_venda_detalhe_sintetico
dim r
dim s, s_aux, s_sql, x, cab_table, cab, fabricante_a
dim n_reg, n_saldo_total, n_saldo_parcial, qtde_fabricantes
cont = 0


    if rb_exportacao = "Normais" then

	    s_sql = "SELECT" & _
				    " t_ESTOQUE_ITEM.fabricante," & _
				    " t_ESTOQUE_ITEM.produto," & _
				    " descricao," & _
				    " descricao_html," & _
				    " Sum(qtde-qtde_utilizada) AS saldo" & _
			    " FROM t_ESTOQUE_ITEM" & _
				    " LEFT JOIN t_PRODUTO ON" & _
					    " ((t_ESTOQUE_ITEM.fabricante=t_PRODUTO.fabricante) AND (t_ESTOQUE_ITEM.produto=t_PRODUTO.produto))" & _
                    " LEFT JOIN t_ESTOQUE ON (t_ESTOQUE_ITEM.id_estoque=t_ESTOQUE.id_estoque)" & _
			    " WHERE" & _
				    " ((qtde-qtde_utilizada) > 0)"

	    if cod_fabricante <> "" then
		    s_sql = s_sql & " AND (t_ESTOQUE_ITEM.fabricante='" & cod_fabricante & "')" 
		    end if

	    if cod_produto <> "" then
		    s_sql = s_sql & " AND (t_ESTOQUE_ITEM.produto='" & cod_produto & "')"
		    end if

        if c_empresa <> "" then
            s_sql = s_sql & " AND (t_ESTOQUE.id_nfe_emitente='" & c_empresa & "')"
        end if
	
	    s_sql = s_sql & _
			    " GROUP BY" & _
				    " t_ESTOQUE_ITEM.fabricante," & _
				    " t_ESTOQUE_ITEM.produto," & _
				    " descricao," & _
				    " descricao_html" & _
			    " ORDER BY" & _
				    " t_ESTOQUE_ITEM.fabricante," & _
				    " t_ESTOQUE_ITEM.produto," & _
				    " descricao," & _
				    " descricao_html"      
	
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
                cont = cont + 1
                r.MoveNext
			
		    loop
    end if

            
    ' CONSULTA COMPOSTOS		
    if rb_exportacao = "Compostos" then

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
							" Coalesce((SELECT Sum(qtde-qtde_utilizada) FROM t_ESTOQUE_ITEM tEI LEFT JOIN t_ESTOQUE tE ON (tEI.id_estoque=tE.id_estoque) WHERE (tEI.fabricante=tP.fabricante) AND (tEI.produto=tP.produto) AND ((qtde-qtde_utilizada)>0)"
                if c_empresa <> "" then s_sql = s_sql & " AND (tE.id_nfe_emitente = '" & c_empresa & "')"    
                s_sql = s_sql & "), 0) AS qtde_estoque_venda," & _
                            " Coalesce((SELECT Sum((tEI.qtde-qtde_utilizada)* vl_custo2) AS saldo FROM t_ESTOQUE_ITEM tEI LEFT JOIN t_ESTOQUE tE ON (tEI.id_estoque=tE.id_estoque) WHERE (tEI.fabricante=tP.fabricante) AND (tEI.produto=tP.produto)"
                if c_empresa <> "" then s_sql = s_sql & " AND (tE.id_nfe_emitente = '" & c_empresa & "')"
                s_sql = s_sql & "), 0) AS vl_custo2" & _                                                                        
						" FROM t_PRODUTO tPL" & _
							" INNER JOIN t_PRODUTO tP ON (tPL.fabricante = tP.fabricante) AND (tPL.produto = tP.produto)" & _                                                                  
						" WHERE " & _
                        " (tP.fabricante = '" & Trim("" & tPCI("fabricante_item")) & "')" & _
                       	" AND (tP.produto = '" & Trim("" & tPCI("produto_item")) & "') "   
                

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
				
			     '> SALDO ESTOQUE
                    redim preserve v_qtde(cont)
				    v_qtde(cont) = Cstr(qtde_estoque_venda_composto)	
    		
			     '> DESCRIÇÃO
                    redim preserve v_descricao(cont)
				    v_descricao(cont) = Trim("" & r("descricao"))	
    		
			     '> NOME DO FABRICANTE
                    redim preserve v_fabricante(cont)
				    v_fabricante(cont) = Trim("" & r("fabricante_composto"))	
    
                    cont = cont + 1    			                                   
		        end if
			end if
			r.MoveNext
			loop

        s_sql = "SELECT" & _
				    " t_ESTOQUE_ITEM.fabricante," & _
				    " t_ESTOQUE_ITEM.produto," & _
				    " descricao," & _
				    " descricao_html," & _
				    " Sum(qtde-qtde_utilizada) AS saldo" & _
			    " FROM t_ESTOQUE_ITEM" & _
				    " LEFT JOIN t_PRODUTO ON" & _
					    " ((t_ESTOQUE_ITEM.fabricante=t_PRODUTO.fabricante) AND (t_ESTOQUE_ITEM.produto=t_PRODUTO.produto))" & _
                    " LEFT JOIN t_ESTOQUE ON (t_ESTOQUE_ITEM.id_estoque=t_ESTOQUE.id_estoque)" & _
			    " WHERE" & _
				    " ((qtde-qtde_utilizada) > 0)" & _
                    " AND (t_ESTOQUE_ITEM.produto NOT IN" & _
                    " (" & _
                    "SELECT produto_item FROM t_EC_PRODUTO_COMPOSTO_ITEM WHERE (fabricante_item=t_ESTOQUE_ITEM.fabricante)" & _
                    "))"

	    if cod_fabricante <> "" then
		    s_sql = s_sql & " AND (t_ESTOQUE_ITEM.fabricante='" & cod_fabricante & "')" 
		    end if

	    if cod_produto <> "" then
		    s_sql = s_sql & " AND (t_ESTOQUE_ITEM.produto='" & cod_produto & "')"
		    end if

        if c_empresa <> "" then
            s_sql = s_sql & " AND (t_ESTOQUE.id_nfe_emitente='" & c_empresa & "')"
        end if
	
	    s_sql = s_sql & _
			    " GROUP BY" & _
				    " t_ESTOQUE_ITEM.fabricante," & _
				    " t_ESTOQUE_ITEM.produto," & _
				    " descricao," & _
				    " descricao_html" & _
			    " ORDER BY" & _
				    " t_ESTOQUE_ITEM.fabricante," & _
				    " t_ESTOQUE_ITEM.produto," & _
				    " descricao," & _
				    " descricao_html"
        
        set r = cn.execute(s_sql)
	    do while Not r.Eof    

	            redim preserve v_fabricante(cont)
                redim preserve v_codigo(cont)
                redim preserve v_descricao(cont)
                redim preserve v_qtde(cont)         
                v_fabricante(cont) =  r("fabricante")
                v_codigo(cont)  =  r("produto")
                v_descricao(cont) = r("descricao")
                v_qtde(cont) = r("saldo")          
                cont = cont + 1
                r.MoveNext
			
		    loop
        
    end if

    ' ORDENA OS REGISTROS DE PRODUTOS NORMAIS E COMPOSTOS DE ACORDO COM SEU CODIGO.

    redim vRelat(0)
	    set vRelat(0) = New cl_SEIS_COLUNAS
	    with vRelat(0)
		    .c1 = ""
		    .c2 = ""
		    .c3 = ""
		    .c4 = ""
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
			end with
        next
    end if

     ordena_cl_seis_colunas vRelat, 0, Ubound(vRelat)


    ' CABEÇALHO
	    cab_table = "<TABLE class='Q' cellSpacing=0>" & chr(13)
	    cab = "	<TR style='background:azure' nowrap>" & chr(13) & _
		      "		<TD width='75' valign='bottom' nowrap class='MD MB'><P class='R'>PRODUTO</P></TD>" & chr(13) & _
		      "		<TD width='480' valign='bottom' nowrap class='MD MB'><P class='R'>DESCRIÇÃO</P></TD>" & chr(13) & _
		      "		<TD width='60' valign='bottom' nowrap class='MB'><P class='Rd' style='font-weight:bold;'>QTDE</P></TD>" & chr(13) & _
		      "	</TR>" & chr(13)
	
	    x = cab_table & cab
	    n_reg = 0
	    n_saldo_total = 0
	    n_saldo_parcial = 0
	    qtde_fabricantes = 0
	    fabricante_a = "XXXXX"



    if vRelat(ubound(vRelat)).c1 <> "" then	
	    for cont=Lbound(vRelat) to Ubound(vRelat)

    '	MUDOU DE FABRICANTE?
		    if vRelat(cont).c2 <> fabricante_a then
		    '	SUB-TOTAL POR FORNECEDOR
			    if n_reg > 0 then
				    x = x & "	<TR NOWRAP>" & chr(13) & _
						    "		<TD class='MB' colspan='2'><P class='Cd'>Total:</P></TD>" & chr(13) & _
						    "		<TD class='MB'><P class='Cd'>" & formata_inteiro(n_saldo_parcial) & "</P></TD>" & chr(13) & _
						    "	</TR>" & chr(13) & _
						    "	<TR NOWRAP>" & chr(13) & _
						    "		<TD colspan='3' class='MB'>&nbsp;</TD>" & _
						    "	</TR>" & chr(13)
				    end if
			    qtde_fabricantes = qtde_fabricantes + 1
			    fabricante_a = vRelat(cont).c2
			    s =  vRelat(cont).c2
			    s_aux = ucase(x_fabricante(s))
			    if (s<>"") And (s_aux<>"") then s = s & " - "
			    s = s & s_aux
			    x = x & "	<TR NOWRAP>" & chr(13) & _
					    "		<TD class='MB' align='center' colspan='3' style='background: honeydew'><P class='C'>&nbsp;" & s & "</P></TD>" & chr(13) & _
					    "	</TR>" & chr(13)
			    n_saldo_parcial = 0
			    end if

	      ' CONTAGEM
		    n_reg = n_reg + 1

		    x = x & "	<TR NOWRAP>" & chr(13)

	     '> PRODUTO
		    x = x & "		<TD class='MDB' valign='bottom' NOWRAP><P class='C'>&nbsp;" & vRelat(cont).c1 & "</P></TD>" & chr(13)

	     '> DESCRIÇÃO
		    x = x & "		<TD class='MDB' valign='bottom'><P class='C' NOWRAP>&nbsp;" & produto_formata_descricao_em_html(vRelat(cont).c3) & "</P></TD>" & chr(13)

	     '> SALDO  
		    x = x & "		<TD class='MB' valign='bottom' NOWRAP><P class='Cd'>&nbsp;" & formata_inteiro(vRelat(cont).c4) & "</P></TD>" & chr(13)

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
				"		<TD colspan='2'><P class='Cd'>Total:</P></TD>" & chr(13) & _
				"		<TD><P class='Cd'>" & formata_inteiro(n_saldo_parcial) & "</P></TD>" & chr(13) & _
				"	</TR>" & chr(13)
		
		if qtde_fabricantes > 1 then
		'	TOTAL GERAL
			x = x & "	<TR NOWRAP><TD COLSPAN='3' class='MC'>&nbsp;</TD></TR>" & chr(13) & _
					"	<TR NOWRAP style='background: #FFFFDD'>" & chr(13) & _
					"		<TD COLSPAN='2' NOWRAP class='MC'><P class='Cd'>" & "TOTAL GERAL:" & "</P></TD>" & chr(13) & _
					"		<TD NOWRAP class='MC'><P class='Cd'>" & formata_inteiro(n_saldo_total) & "</P></TD>" & chr(13) & _
					"	</TR>" & chr(13)
			end if
		end if

  ' MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	if n_reg = 0 then
		x = cab_table & cab & _
			"	<TR NOWRAP>" & chr(13) & _
			"		<TD colspan='3'><P class='ALERTA'>&nbsp;NENHUM PRODUTO DO ESTOQUE SATISFAZ AS CONDIÇÕES ESPECIFICADAS&nbsp;</P></TD>" & chr(13) & _
			"	</TR>" & chr(13)
		end if

  ' FECHA TABELA
	x = x & "</TABLE>" & chr(13)
	
	Response.write x

	if r.State <> 0 then r.Close
	set r=nothing

end sub




' ________________________________________________
' CONSULTA ESTOQUE VENDA DETALHE INTERMEDIARIO
' 
sub consulta_estoque_venda_detalhe_intermediario
dim r
dim s, s_aux, s_sql, x, cab_table, cab, fabricante_a
dim n_reg, n_saldo_total, n_saldo_parcial, qtde_fabricantes
dim vl, vl_total_geral, vl_sub_total

'	IMPORTANTE: O VALOR ATUAL DE CMV_PV ESTÁ EM T_PRODUTO.PRECO_FABRICANTE
'	==========  O HISTÓRICO DO VALOR DE CMV_PV ESTÁ EM T_PEDIDO_ITEM.PRECO_FABRICANTE (E T_PEDIDO_ITEM_DEVOLVIDO.PRECO_FABRICANTE)
'				O HISTÓRICO DO CUSTO REAL PAGO AO FABRICANTE ESTÁ EM T_ESTOQUE_ITEM.VL_CUSTO2

    if rb_exportacao = "Normais" then

	    s_sql = "SELECT" & _
				    " t_ESTOQUE_ITEM.fabricante," & _
				    " t_ESTOQUE_ITEM.produto," & _
				    " descricao," & _
				    " descricao_html," & _
				    " Sum(qtde-qtde_utilizada) AS saldo," & _
				    " Sum((qtde-qtde_utilizada)*t_ESTOQUE_ITEM.vl_custo2) AS preco_total" & _
			    " FROM t_ESTOQUE_ITEM" & _
				    " LEFT JOIN t_PRODUTO ON" & _
					    " ((t_ESTOQUE_ITEM.fabricante=t_PRODUTO.fabricante) AND (t_ESTOQUE_ITEM.produto=t_PRODUTO.produto))" & _
                    " LEFT JOIN t_ESTOQUE ON (t_ESTOQUE_ITEM.id_estoque=t_ESTOQUE.id_estoque)" & _
			    " WHERE" & _
				    " ((qtde-qtde_utilizada) > 0)"	


	    if cod_fabricante <> "" then
		    s_sql = s_sql & " AND (t_ESTOQUE_ITEM.fabricante='" & cod_fabricante & "')" 
		    end if

	    if cod_produto <> "" then
		    s_sql = s_sql & " AND (t_ESTOQUE_ITEM.produto='" & cod_produto & "')"
		    end if

        if c_empresa <> "" then
            s_sql = s_sql & " AND (t_ESTOQUE.id_nfe_emitente='" & c_empresa & "')"
        end if

	    s_sql = s_sql & _
			    " GROUP BY" & _
				    " t_ESTOQUE_ITEM.fabricante," & _
				    " t_ESTOQUE_ITEM.produto," & _
				    " descricao," & _
				    " descricao_html" & _
			    " ORDER BY" & _
				    " t_ESTOQUE_ITEM.fabricante," & _
				    " t_ESTOQUE_ITEM.produto," & _
				    " descricao," & _
				    " descricao_html"      
	
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
                cont = cont + 1
			
		    r.movenext
		    loop
    end if   
       
            
    ' CONSULTA COMPOSTOS
		
    if rb_exportacao = "Compostos" then

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
							" Coalesce((SELECT Sum(qtde-qtde_utilizada) FROM t_ESTOQUE_ITEM tEI LEFT JOIN t_ESTOQUE tE ON (tEI.id_estoque=tE.id_estoque) WHERE (tEI.fabricante=tP.fabricante) AND (tEI.produto=tP.produto) AND ((qtde-qtde_utilizada)>0)"
                if c_empresa <> "" then s_sql = s_sql & " AND (tE.id_nfe_emitente = '" & c_empresa & "')"    
                s_sql = s_sql & "), 0) AS qtde_estoque_venda," & _
                            " Coalesce((SELECT Sum((tEI.qtde-qtde_utilizada)* vl_custo2) AS saldo FROM t_ESTOQUE_ITEM tEI LEFT JOIN t_ESTOQUE tE ON (tEI.id_estoque=tE.id_estoque) WHERE (tEI.fabricante=tP.fabricante) AND (tEI.produto=tP.produto)"
                if c_empresa <> "" then s_sql = s_sql & " AND (tE.id_nfe_emitente = '" & c_empresa & "')"
                s_sql = s_sql & "), 0) AS vl_custo2" & _                                                               
						" FROM t_PRODUTO tPL" & _
							" INNER JOIN t_PRODUTO tP ON (tPL.fabricante = tP.fabricante) AND (tPL.produto = tP.produto)" & _                                                                  
						" WHERE " & _
                        " (tP.fabricante = '" & Trim("" & tPCI("fabricante_item")) & "')" & _
                       	" AND (tP.produto = '" & Trim("" & tPCI("produto_item")) & "') "   
                

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
				
			     '> SALDO ESTOQUE
                    redim preserve v_qtde(cont)
				    v_qtde(cont) = Cstr(qtde_estoque_venda_composto)	
                
                '> VALOR CUSTO
			    '	EXPORTAR VALOR UTILIZANDO '.' COMO SEPARADOR DECIMAL
                    redim preserve v_valor(cont)
				    v_valor(cont) = formata_moeda(vl_custo2_composto * qtde_estoque_venda_composto )
    		
			     '> DESCRIÇÃO
                    redim preserve v_descricao(cont)
				    v_descricao(cont) = Trim("" & r("descricao"))	
    		
			     '> NOME DO FABRICANTE
                    redim preserve v_fabricante(cont)
				    v_fabricante(cont) = Trim("" & r("fabricante_composto"))	
    
                    cont = cont + 1    			                                   
		        end if
			end if
			r.MoveNext
			loop

               s_sql = "SELECT" & _
				    " t_ESTOQUE_ITEM.fabricante," & _
				    " t_ESTOQUE_ITEM.produto," & _
				    " descricao," & _
				    " descricao_html," & _
				    " Sum(qtde-qtde_utilizada) AS saldo," & _
                    " Sum((qtde-qtde_utilizada)*t_ESTOQUE_ITEM.vl_custo2) AS preco_total" & _
			    " FROM t_ESTOQUE_ITEM" & _
				    " LEFT JOIN t_PRODUTO ON" & _
					    " ((t_ESTOQUE_ITEM.fabricante=t_PRODUTO.fabricante) AND (t_ESTOQUE_ITEM.produto=t_PRODUTO.produto))" & _
                    " LEFT JOIN t_ESTOQUE ON (t_ESTOQUE_ITEM.id_estoque=t_ESTOQUE.id_estoque)" & _
			    " WHERE" & _
				    " ((qtde-qtde_utilizada) > 0)" & _
                    " AND (t_ESTOQUE_ITEM.produto NOT IN" & _
                    " (" & _
                    "SELECT produto_item FROM t_EC_PRODUTO_COMPOSTO_ITEM WHERE (fabricante_item=t_ESTOQUE_ITEM.fabricante)" & _
                    "))"

	    if cod_fabricante <> "" then
		    s_sql = s_sql & " AND (t_ESTOQUE_ITEM.fabricante='" & cod_fabricante & "')" 
		    end if

	    if cod_produto <> "" then
		    s_sql = s_sql & " AND (t_ESTOQUE_ITEM.produto='" & cod_produto & "')"
		    end if

        if c_empresa <> "" then
            s_sql = s_sql & " AND (t_ESTOQUE.id_nfe_emitente='" & c_empresa & "')"
        end if
	
	    s_sql = s_sql & _
			    " GROUP BY" & _
				    " t_ESTOQUE_ITEM.fabricante," & _
				    " t_ESTOQUE_ITEM.produto," & _
				    " descricao," & _
				    " descricao_html" & _
			    " ORDER BY" & _
				    " t_ESTOQUE_ITEM.fabricante," & _
				    " t_ESTOQUE_ITEM.produto," & _
				    " descricao," & _
				    " descricao_html"
        
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
                cont = cont + 1

                r.MoveNext		
		    loop

    end if




    ' ORDENA OS REGISTROS DE PRODUTOS NORMAIS E COMPOSTOS DE ACORDO COM SEU CODIGO.
        redim vRelat(0)
	    set vRelat(0) = New cl_SEIS_COLUNAS
	    with vRelat(0)
		    .c1 = ""
		    .c2 = ""
		    .c3 = ""
		    .c4 = ""
		    .c5 = ""
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
			end with
        next
    end if

    ordena_cl_seis_colunas vRelat, 0, Ubound(vRelat)


    ' CABEÇALHO
	    cab_table = "<TABLE class='Q' cellSpacing=0>" & chr(13)
	    cab = "	<TR style='background:azure' nowrap>" & chr(13) & _
		      "		<TD width='60' valign='bottom' nowrap class='MD MB'><P class='R'>PRODUTO</P></TD>" & chr(13) & _
		      "		<TD width='274' valign='bottom' class='MD MB'><P class='R'>DESCRIÇÃO</P></TD>" & chr(13) & _
		      "		<TD width='60' valign='bottom' nowrap class='MD MB'><P class='Rd' style='font-weight:bold;'>QTDE</P></TD>" & chr(13) & _
		      "		<TD width='100' valign='bottom' nowrap class='MD MB'><P class='Rd' style='font-weight:bold;'>CUSTO ENTRADA UNITÁRIO MÉDIO</P></TD>" & chr(13) & _
		      "		<TD width='100' valign='bottom' nowrap class='MB'><P class='Rd' style='font-weight:bold;'>CUSTO ENTRADA TOTAL</P></TD>" & chr(13) & _
		      "	</TR>" & chr(13)
	
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
						    "		<TD class='MB' colspan='2'><P class='Cd'>Total:</P></TD>" & chr(13) & _
						    "		<TD class='MB'><P class='Cd'>" & formata_inteiro(n_saldo_parcial) & "</P></TD>" & chr(13) & _
						    "		<TD class='MB'>&nbsp;</TD>" & chr(13) & _
						    "		<TD class='MB'><P class='Cd'>" & formata_moeda(vl_sub_total) & "</P></TD>" & chr(13) & _
						    "	</TR>" & chr(13) & _
						    "	<TR NOWRAP>" & chr(13) & _
						    "		<TD colspan='5' class='MB'>&nbsp;</TD>" & _
						    "	</TR>" & chr(13)
				    end if
			    qtde_fabricantes = qtde_fabricantes + 1
			    fabricante_a = vRelat(cont).c2
			    s = vRelat(cont).c2
			    s_aux = ucase(x_fabricante(s))
			    if (s<>"") And (s_aux<>"") then s = s & " - "
			    s = s & s_aux
			    x = x & "	<TR NOWRAP>" & chr(13) & _
					    "		<TD class='MB' align='center' colspan='5' style='background: honeydew'><P class='C'>&nbsp;" & s & "</P></TD>" & chr(13) & _
					    "	</TR>" & chr(13)
			    n_saldo_parcial = 0
			    vl_sub_total = 0
			    end if

	      ' CONTAGEM
		    n_reg = n_reg + 1

		    x = x & "	<TR NOWRAP>" & chr(13)

	     '> PRODUTO
		    x = x & "		<TD class='MDB' valign='bottom' NOWRAP><P class='C'>&nbsp;" & vRelat(cont).c1 & "</P></TD>" & chr(13)

	     '> DESCRIÇÃO
		    x = x & "		<TD class='MDB' valign='bottom'><P class='C'>&nbsp;" & produto_formata_descricao_em_html(vRelat(cont).c3) & "</P></TD>" & chr(13)

	     '> SALDO
		    x = x & "		<TD class='MDB' valign='bottom' NOWRAP><P class='Cd'>&nbsp;" & formata_inteiro(vRelat(cont).c4) & "</P></TD>" & chr(13)
	
	     '> CUSTO DE ENTRADA UNITÁRIO (MÉDIA)
		    if vRelat(cont).c4 = 0 then vl = 0 else vl = vRelat(cont).c5/vRelat(cont).c4
		    x = x & "		<TD class='MDB' valign='bottom' NOWRAP><P class='Cd'>" & formata_moeda(vl) & "</P></TD>" & chr(13)

	     '> CUSTO DE ENTRADA TOTAL
		    vl = vRelat(cont).c5
		    x = x & "		<TD class='MB' valign='bottom' NOWRAP><P class='Cd'>" & formata_moeda(vl) & "</P></TD>" & chr(13)
		
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

    '''''''''''''''''''''

  ' MOSTRA TOTAL
	if n_reg <> 0 then 
	'	TOTAL DO ÚLTIMO FORNECEDOR
		x = x & "	<TR NOWRAP>" & chr(13) & _
						"		<TD colspan='2'><P class='Cd'>Total:</P></TD>" & chr(13) & _
						"		<TD><P class='Cd'>" & formata_inteiro(n_saldo_parcial) & "</P></TD>" & chr(13) & _
						"		<TD>&nbsp;</TD>" & chr(13) & _
						"		<TD><P class='Cd'>" & formata_moeda(vl_sub_total) & "</P></TD>" & chr(13) & _
				"	</TR>" & chr(13)

		if qtde_fabricantes > 1 then
		'	TOTAL GERAL
			x = x & "	<TR NOWRAP><TD COLSPAN='5' class='MC'>&nbsp;</TD></TR>" & chr(13) & _
					"	<TR NOWRAP style='background: #FFFFDD'>" & chr(13) & _
					"		<TD COLSPAN='2' NOWRAP class='MC'><P class='Cd'>TOTAL GERAL:</P></TD>" & chr(13) & _
					"		<TD NOWRAP class='MC'><P class='Cd'>" & formata_inteiro(n_saldo_total) & "</P></TD>" & chr(13) & _
					"		<TD class='MC'>&nbsp;</TD>" & chr(13) & _
					"		<TD NOWRAP class='MC'><P class='Cd'>" & formata_moeda(vl_total_geral) & "</P></TD>" & chr(13) & _
					"	</TR>" & chr(13)
			end if
		end if

  ' MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	if n_reg = 0 then
		x = cab_table & cab & _
			"	<TR NOWRAP>" & chr(13) & _
			"		<TD colspan='5'><P class='ALERTA'>&nbsp;NENHUM PRODUTO DO ESTOQUE SATISFAZ AS CONDIÇÕES ESPECIFICADAS&nbsp;</P></TD>" & chr(13) & _
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

<style TYPE="text/css">
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
<style TYPE="text/css">
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
<table cellSpacing="0">
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

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Estoque</span>
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
<table class="Qx" cellSpacing="0">
<!--  ESTOQUE  -->
	<tr bgColor="#FFFFFF">
		<% select case rb_estoque
				case ID_ESTOQUE_VENDA:			s = "VENDA"
				case ID_ESTOQUE_VENDIDO:		s = "VENDIDO"
				case ID_ESTOQUE_SHOW_ROOM:		s = "SHOW-ROOM"
				case ID_ESTOQUE_DANIFICADOS:	s = "PRODUTOS DANIFICADOS"
				case ID_ESTOQUE_DEVOLUCAO:		s = "DEVOLUÇÃO"
				case else						s = ""
				end select
		%>
	<td class="MT" NOWRAP><span class="PLTe">Estoque de Interesse</span>
		<br><p class="C" style="width:230px;cursor:default;"><%=s%></p></td>

<!--  TIPO DE DETALHAMENTO  -->
		<% select case rb_detalhe
			case "SINTETICO":		s = "SINTÉTICO (SEM CUSTOS)"
			case "INTERMEDIARIO":	s = "INTERMEDIÁRIO (CUSTOS MÉDIOS)"
			case "COMPLETO":		s = "COMPLETO (CUSTOS DIFERENCIADOS)"
			case else				s = ""
			end select
		%>
	<td class="MT" style="border-left:0px;" NOWRAP><span class="PLTe">Tipo de Detalhamento</span>
		<br><p class="C" style="width:230px;cursor:default;"><%=s%></p></td>
	</tr>

    <tr bgColor="#FFFFFF">
<!--  TIPO DE CONSULTA  -->
		<% select case rb_exportacao
			case "Normais":		s = "PRODUTOS NORMAIS"
			case "Compostos":	s = "PRODUTOS UNIFICADOS"
			case else				s = ""
			end select
		%>
	<td class="MDBE" NOWRAP colspan="2"><span class="PLTe">Tipo de Consulta</span>
		<br><p class="C" style="width:230px;cursor:default;"><%=s%></p></td>
	</tr>

<!--  FABRICANTE  -->
	<% if cod_fabricante <> "" then %>
		<tr bgColor="#FFFFFF">
		<td class="MDBE" NOWRAP colspan="2"><span class="PLTe">Fabricante</span>
			<%	s = cod_fabricante
				if (s<>"") And (s_nome_fabricante<>"") then s = s & " - " & s_nome_fabricante %>
			<br><input name="c_fabricante_aux" id="c_fabricante_aux" readonly tabindex=-1 class="PLLe" style="width:460px;margin-left:2pt;" 
					value="<%=s%>"></td>
		</tr>
	<% end if %>
	
<!--  PRODUTO  -->
	<% if cod_produto <> "" then %>
		<tr bgColor="#FFFFFF">
		<td class="MDBE" NOWRAP colspan="2"><span class="PLTe">Produto</span>
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

<!--  EMPRESA  -->
	<% if c_empresa = "" then
			s = "Todas"
		else
			s = obtem_apelido_empresa_NFe_emitente(c_empresa)
			end if
	%>
		<tr bgColor="#FFFFFF">
			<td class="MDBE" NOWRAP colspan="2"><span class="PLTe">Empresa</span>
				<br><p class="C" style="width:460px;margin-left:2pt;"><%=s%></p>
			</td>
		</tr>
	
</table>

<!--  RELATÓRIO  -->
<br>
<%	
	if rb_estoque = ID_ESTOQUE_VENDA then
		select case rb_detalhe
			case "SINTETICO"
				consulta_estoque_venda_detalhe_sintetico
			case "INTERMEDIARIO"
				consulta_estoque_venda_detalhe_intermediario
			end select
		end if
%>

<!-- ************   SEPARADOR   ************ -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>
<br>


<table class="notPrint" width="649" cellSpacing="0">
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
