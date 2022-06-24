<%
' =============================
' F U N Ç Õ E S  G L O B A I S 
' =============================

' _____________________________________________
' FABRICANTE_MONTA_ITENS_SELECT
'
function fabricante_monta_itens_select(byval id_default)
dim x, r, strResp, ha_default
	id_default = Trim("" & id_default)
	ha_default=False
	set r = cn.Execute("SELECT * FROM t_FABRICANTE ORDER BY fabricante")
	strResp = ""
	do while Not r.eof 
		x = Trim("" & r("fabricante"))
		if (id_default<>"") And (id_default=x) then
			strResp = strResp & "<option selected"
			ha_default=True
		else
			strResp = strResp & "<option"
			end if
		strResp = strResp & " value='" & x & "'>"
		strResp = strResp & Trim("" & r("fabricante")) & " - " & Trim("" & r("nome"))
		strResp = strResp & "</option>" & chr(13)
		r.MoveNext
		loop

	if Not ha_default then
		strResp = "<option selected value=''>&nbsp;</option>" & chr(13) & strResp
		end if
		
	fabricante_monta_itens_select = strResp
	r.close
	set r=nothing
end function




' _____________________________________________
' TRANSPORTADORA_MONTA_ITENS_SELECT
'
function transportadora_monta_itens_select(byval id_default)
dim x, r, strResp, ha_default
	id_default = Trim("" & id_default)
	ha_default=False
	set r = cn.Execute("SELECT * FROM t_TRANSPORTADORA ORDER BY id")
	strResp = ""
	do while Not r.eof 
		x = Trim("" & r("id"))
		if (id_default<>"") And (id_default=x) then
			strResp = strResp & "<option selected"
			ha_default=True
		else
			strResp = strResp & "<option"
			end if
		strResp = strResp & " value='" & x & "'>"
		strResp = strResp & Trim("" & r("id")) & " - " & Trim("" & r("nome"))
		strResp = strResp & "</option>" & chr(13)
		r.MoveNext
		loop

	if Not ha_default then
		strResp = "<option selected value=''>&nbsp;</option>" & chr(13) & strResp
		end if
		
	transportadora_monta_itens_select = strResp
	r.close
	set r=nothing
end function



' ________________________________________________________
' DESEMPENHO_NOTA_MONTA_ITENS_SELECT (TODOS)
'
function desempenho_nota_monta_itens_select(byval codigo_default)
dim i, x, strResp, ha_default
dim ListaCodigos
	ListaCodigos = Array("A", "B", "C", "D", "E")
	codigo_default = Trim("" & codigo_default)
	ha_default=False
	strResp = ""
	for i = Lbound(ListaCodigos) to Ubound(ListaCodigos)
		x = Trim("" & ListaCodigos(i))
		if (codigo_default<>"") And (codigo_default=x) then
			strResp = strResp & "<option selected"
			ha_default=True
		else
			strResp = strResp & "<option"
			end if
		strResp = strResp & " value='" & x & "'>"
		strResp = strResp & x
		strResp = strResp & "</option>" & chr(13)
		next

	if Not ha_default then
		strResp = "<option selected value=''>&nbsp;</option>" & chr(13) & strResp
		end if
		
	desempenho_nota_monta_itens_select = strResp
end function



' _____________________________________________
' FORMA_PAGTO_MONTA_ITENS_SELECT (TODOS)
'
function forma_pagto_monta_itens_select(byval id_default)
dim x, r, strResp, ha_default
	id_default = Trim("" & id_default)
	ha_default=False
	set r = cn.Execute("SELECT * FROM t_FORMA_PAGTO ORDER BY ordenacao")
	strResp = ""
	do while Not r.eof 
		x = Trim("" & r("id"))
		if (id_default<>"") And (id_default=x) then
			strResp = strResp & "<option selected"
			ha_default=True
		else
			strResp = strResp & "<option"
			end if
		strResp = strResp & " value='" & x & "'>"
		strResp = strResp & Trim("" & r("descricao"))
		strResp = strResp & "</option>" & chr(13)
		r.MoveNext
		loop

	if Not ha_default then
		strResp = "<option selected value=''>&nbsp;</option>" & chr(13) & strResp
		end if
		
	forma_pagto_monta_itens_select = strResp
	r.close
	set r=nothing
end function



' _____________________________________________
' FORMA_PAGTO_AV_MONTA_ITENS_SELECT (À VISTA)
'
function forma_pagto_av_monta_itens_select(byval id_default)
dim x, r, strResp, ha_default
	id_default = Trim("" & id_default)
	ha_default=False
	set r = cn.Execute("SELECT * FROM t_FORMA_PAGTO WHERE hab_a_vista=1 ORDER BY ordenacao")
	strResp = ""
	do while Not r.eof 
		x = Trim("" & r("id"))
		if (id_default<>"") And (id_default=x) then
			strResp = strResp & "<option selected"
			ha_default=True
		else
			strResp = strResp & "<option"
			end if
		strResp = strResp & " value='" & x & "'>"
		strResp = strResp & Trim("" & r("descricao"))
		strResp = strResp & "</option>" & chr(13)
		r.MoveNext
		loop

	if Not ha_default then
		strResp = "<option selected value=''>&nbsp;</option>" & chr(13) & strResp
		end if
		
	forma_pagto_av_monta_itens_select = strResp
	r.close
	set r=nothing
end function



' ________________________________________________________________
' FORMA_PAGTO_AV_MONTA_ITENS_SELECT_INCLUINDO_DEFAULT (À VISTA)
'
function forma_pagto_av_monta_itens_select_incluindo_default(byval id_default)
dim x, s, r, strResp, ha_default
	id_default = Trim("" & id_default)
	ha_default=False
	s = "SELECT " & _
			"*" & _
		" FROM t_FORMA_PAGTO" & _
		" WHERE" & _
			" (id = " & id_default & ")" & _
			" OR" & _
			" (hab_a_vista = 1)" & _
		" ORDER BY" & _
			" ordenacao"
	set r = cn.Execute(s)
	strResp = ""
	do while Not r.eof 
		x = Trim("" & r("id"))
		if (id_default<>"") And (id_default=x) then
			strResp = strResp & "<option selected"
			ha_default=True
		else
			strResp = strResp & "<option"
			end if
		strResp = strResp & " value='" & x & "'>"
		strResp = strResp & Trim("" & r("descricao"))
		strResp = strResp & "</option>" & chr(13)
		r.MoveNext
		loop

	if Not ha_default then
		strResp = "<option selected value=''>&nbsp;</option>" & chr(13) & strResp
		end if
		
	forma_pagto_av_monta_itens_select_incluindo_default = strResp
	r.close
	set r=nothing
end function

' _____________________________________________________________________________________
' FORMA_PAGTO_LIBERADA_AV_MONTA_ITENS_SELECT (À VISTA)
' VERIFICA NA TABELA t_ORCAMENTISTA_E_INDICADOR_RESTRICAO_FORMA_PAGTO SE HÁ
' RESTRIÇÕES P/ O REFERIDO 'INDICADOR' P/ O TIPO DE CLIENTE (PF/PJ) ESPECIFICADO.
' NO CASO DE NÃO SER INFORMADO O INDICADOR OU TIPO DE CLIENTE, TODAS AS FORMAS DE
' PAGAMENTO SÃO RETORNADAS.
function forma_pagto_liberada_av_monta_itens_select(byval id_default, byval id_orcamentista_e_indicador, byval tipo_cliente)
dim s, x, r, strResp, ha_default
	id_default = Trim("" & id_default)
	ha_default=False
	s = "SELECT " & _
			"*" & _
		" FROM t_FORMA_PAGTO tFP" & _
		" WHERE" & _
			" (hab_a_vista=1)" & _
			" AND " & _
			"(" & _
				"id NOT IN " & _
					"(" & _
					"SELECT" & _
						" id_forma_pagto" & _
					" FROM t_ORCAMENTISTA_E_INDICADOR_RESTRICAO_FORMA_PAGTO tOIRFP" & _
					" WHERE" & _ 
						"(" & _
							"(id_orcamentista_e_indicador = '" & Trim(id_orcamentista_e_indicador) & "')" & _
							" OR " & _
							"(id_orcamentista_e_indicador = '" & ID_ORCAMENTISTA_E_INDICADOR_RESTRICAO_FP_TODOS & "')" & _
						")" & _
						" AND (tipo_cliente = '" & Trim(tipo_cliente) & "')" & _
						" AND (st_restricao_ativa <> 0)" & _
					")" & _
			")" & _
		" ORDER BY" & _
			" ordenacao"
	set r = cn.Execute(s)
	strResp = ""
	do while Not r.eof 
		x = Trim("" & r("id"))
		if (id_default<>"") And (id_default=x) then
			strResp = strResp & "<option selected"
			ha_default=True
		else
			strResp = strResp & "<option"
			end if
		strResp = strResp & " value='" & x & "'>"
		strResp = strResp & Trim("" & r("descricao"))
		strResp = strResp & "</option>" & chr(13)
		r.MoveNext
		loop

	if Not ha_default then
		strResp = "<option selected value=''>&nbsp;</option>" & chr(13) & strResp
		end if
	
	forma_pagto_liberada_av_monta_itens_select = strResp
	r.close
	set r=nothing
end function



' _____________________________________________________________________________________
' FORMA_PAGTO_LIBERADA_AV_MONTA_ITENS_SELECT_INCLUINDO_DEFAULT (À VISTA)
' VERIFICA NA TABELA t_ORCAMENTISTA_E_INDICADOR_RESTRICAO_FORMA_PAGTO SE HÁ
' RESTRIÇÕES P/ O REFERIDO 'INDICADOR' P/ O TIPO DE CLIENTE (PF/PJ) ESPECIFICADO.
' NO CASO DE NÃO SER INFORMADO O INDICADOR OU TIPO DE CLIENTE, TODAS AS FORMAS DE
' PAGAMENTO SÃO RETORNADAS.
' ASSEGURA A INCLUSÃO DA OPÇÃO DEFAULT NA LISTA.
function forma_pagto_liberada_av_monta_itens_select_incluindo_default(byval id_default, byval id_orcamentista_e_indicador, byval tipo_cliente)
dim s, x, r, strResp, ha_default
	id_default = Trim("" & id_default)
	ha_default=False
	s = "SELECT " & _
			"*" & _
		" FROM t_FORMA_PAGTO tFP" & _
		" WHERE" & _
			" (id = " & id_default & ")" & _
			" OR " & _
			"(" & _
				"(hab_a_vista=1)" & _
				" AND " & _
				"(id NOT IN " & _
					"(" & _
					"SELECT" & _
						" id_forma_pagto" & _
					" FROM t_ORCAMENTISTA_E_INDICADOR_RESTRICAO_FORMA_PAGTO tOIRFP" & _
					" WHERE" & _ 
						"(" & _
							"(id_orcamentista_e_indicador = '" & Trim(id_orcamentista_e_indicador) & "')" & _
							" OR " & _
							"(id_orcamentista_e_indicador = '" & ID_ORCAMENTISTA_E_INDICADOR_RESTRICAO_FP_TODOS & "')" & _
						")" & _
						" AND (tipo_cliente = '" & Trim(tipo_cliente) & "')" & _
						" AND (st_restricao_ativa <> 0)" & _
					")" & _
				")" & _
			")" & _
		" ORDER BY" & _
			" ordenacao"
	set r = cn.Execute(s)
	strResp = ""
	do while Not r.eof 
		x = Trim("" & r("id"))
		if (id_default<>"") And (id_default=x) then
			strResp = strResp & "<option selected"
			ha_default=True
		else
			strResp = strResp & "<option"
			end if
		strResp = strResp & " value='" & x & "'>"
		strResp = strResp & Trim("" & r("descricao"))
		strResp = strResp & "</option>" & chr(13)
		r.MoveNext
		loop

	if Not ha_default then
		strResp = "<option selected value=''>&nbsp;</option>" & chr(13) & strResp
		end if
	
	forma_pagto_liberada_av_monta_itens_select_incluindo_default = strResp
	r.close
	set r=nothing
end function



' ____________________________________________
' FORMA_PAGTO_DA_ENTRADA_MONTA_ITENS_SELECT
'
function forma_pagto_da_entrada_monta_itens_select(byval id_default)
dim x, r, strResp, ha_default
	id_default = Trim("" & id_default)
	ha_default=False
	set r = cn.Execute("SELECT * FROM t_FORMA_PAGTO WHERE hab_entrada=1 ORDER BY ordenacao")
	strResp = ""
	do while Not r.eof 
		x = Trim("" & r("id"))
		if (id_default<>"") And (id_default=x) then
			strResp = strResp & "<option selected"
			ha_default=True
		else
			strResp = strResp & "<option"
			end if
		strResp = strResp & " value='" & x & "'>"
		strResp = strResp & Trim("" & r("descricao"))
		strResp = strResp & "</option>" & chr(13)
		r.MoveNext
		loop

	if Not ha_default then
		strResp = "<option selected value=''>&nbsp;</option>" & chr(13) & strResp
		end if
		
	forma_pagto_da_entrada_monta_itens_select = strResp
	r.close
	set r=nothing
end function



' _______________________________________________________________
' FORMA_PAGTO_DA_ENTRADA_MONTA_ITENS_SELECT_INCLUINDO_DEFAULT
'
function forma_pagto_da_entrada_monta_itens_select_incluindo_default(byval id_default)
dim x, s, r, strResp, ha_default
	id_default = Trim("" & id_default)
	ha_default=False
	s = "SELECT " & _
			"*" & _
		" FROM t_FORMA_PAGTO" & _
		" WHERE" & _
			" (id = " & id_default & ")" & _
			" OR " & _
			" (hab_entrada = 1)" & _
		" ORDER BY" & _
			" ordenacao"
	set r = cn.Execute(s)
	strResp = ""
	do while Not r.eof 
		x = Trim("" & r("id"))
		if (id_default<>"") And (id_default=x) then
			strResp = strResp & "<option selected"
			ha_default=True
		else
			strResp = strResp & "<option"
			end if
		strResp = strResp & " value='" & x & "'>"
		strResp = strResp & Trim("" & r("descricao"))
		strResp = strResp & "</option>" & chr(13)
		r.MoveNext
		loop

	if Not ha_default then
		strResp = "<option selected value=''>&nbsp;</option>" & chr(13) & strResp
		end if
		
	forma_pagto_da_entrada_monta_itens_select_incluindo_default = strResp
	r.close
	set r=nothing
end function



' ___________________________________________________________________________________________
' FORMA_PAGTO_LIBERADA_DA_ENTRADA_MONTA_ITENS_SELECT
' VERIFICA NA TABELA t_ORCAMENTISTA_E_INDICADOR_RESTRICAO_FORMA_PAGTO SE HÁ
' RESTRIÇÕES P/ O REFERIDO 'INDICADOR' P/ O TIPO DE CLIENTE (PF/PJ) ESPECIFICADO.
' NO CASO DE NÃO SER INFORMADO O INDICADOR OU TIPO DE CLIENTE, TODAS AS FORMAS DE
' PAGAMENTO SÃO RETORNADAS.
function forma_pagto_liberada_da_entrada_monta_itens_select(byval id_default, byval id_orcamentista_e_indicador, byval tipo_cliente)
dim s, x, r, strResp, ha_default
	id_default = Trim("" & id_default)
	ha_default=False
	s = "SELECT " & _
			"*" & _
		" FROM t_FORMA_PAGTO" & _
		" WHERE" & _
			" (hab_entrada=1)" & _
			" AND " & _
			"(" & _
				"id NOT IN " & _
					"(" & _
					"SELECT" & _
						" id_forma_pagto" & _
					" FROM t_ORCAMENTISTA_E_INDICADOR_RESTRICAO_FORMA_PAGTO tOIRFP" & _
					" WHERE" & _ 
						"(" & _
							"(id_orcamentista_e_indicador = '" & Trim(id_orcamentista_e_indicador) & "')" & _
							" OR " & _
							"(id_orcamentista_e_indicador = '" & ID_ORCAMENTISTA_E_INDICADOR_RESTRICAO_FP_TODOS & "')" & _
						")" & _
						" AND (tipo_cliente = '" & Trim(tipo_cliente) & "')" & _
						" AND (st_restricao_ativa <> 0)" & _
					")" & _
			")" & _
		" ORDER BY" & _
			" ordenacao"
	set r = cn.Execute(s)
	strResp = ""
	do while Not r.eof 
		x = Trim("" & r("id"))
		if (id_default<>"") And (id_default=x) then
			strResp = strResp & "<option selected"
			ha_default=True
		else
			strResp = strResp & "<option"
			end if
		strResp = strResp & " value='" & x & "'>"
		strResp = strResp & Trim("" & r("descricao"))
		strResp = strResp & "</option>" & chr(13)
		r.MoveNext
		loop

	if Not ha_default then
		strResp = "<option selected value=''>&nbsp;</option>" & chr(13) & strResp
		end if
		
	forma_pagto_liberada_da_entrada_monta_itens_select = strResp
	r.close
	set r=nothing
end function



' ____________________________________________________________________________________
' FORMA_PAGTO_LIBERADA_DA_ENTRADA_MONTA_ITENS_SELECT_INCLUINDO_DEFAULT
' VERIFICA NA TABELA t_ORCAMENTISTA_E_INDICADOR_RESTRICAO_FORMA_PAGTO SE HÁ
' RESTRIÇÕES P/ O REFERIDO 'INDICADOR' P/ O TIPO DE CLIENTE (PF/PJ) ESPECIFICADO.
' NO CASO DE NÃO SER INFORMADO O INDICADOR OU TIPO DE CLIENTE, TODAS AS FORMAS DE
' PAGAMENTO SÃO RETORNADAS.
' ASSEGURA A INCLUSÃO DA OPÇÃO DEFAULT NA LISTA.
function forma_pagto_liberada_da_entrada_monta_itens_select_incluindo_default(byval id_default, byval id_orcamentista_e_indicador, byval tipo_cliente)
dim s, x, r, strResp, ha_default
	id_default = Trim("" & id_default)
	ha_default=False
	s = "SELECT " & _
			"*" & _
		" FROM t_FORMA_PAGTO" & _
		" WHERE" & _
			" (id = " & id_default & ")" & _
			" OR " & _
			"(" & _
				"(hab_entrada=1)" & _
				" AND " & _
				"(id NOT IN " & _
					"(" & _
					"SELECT" & _
						" id_forma_pagto" & _
					" FROM t_ORCAMENTISTA_E_INDICADOR_RESTRICAO_FORMA_PAGTO tOIRFP" & _
					" WHERE" & _ 
						"(" & _
							"(id_orcamentista_e_indicador = '" & Trim(id_orcamentista_e_indicador) & "')" & _
							" OR " & _
							"(id_orcamentista_e_indicador = '" & ID_ORCAMENTISTA_E_INDICADOR_RESTRICAO_FP_TODOS & "')" & _
						")" & _
						" AND (tipo_cliente = '" & Trim(tipo_cliente) & "')" & _
						" AND (st_restricao_ativa <> 0)" & _
					")" & _
				")" & _
			")" & _
		" ORDER BY" & _
			" ordenacao"
	set r = cn.Execute(s)
	strResp = ""
	do while Not r.eof 
		x = Trim("" & r("id"))
		if (id_default<>"") And (id_default=x) then
			strResp = strResp & "<option selected"
			ha_default=True
		else
			strResp = strResp & "<option"
			end if
		strResp = strResp & " value='" & x & "'>"
		strResp = strResp & Trim("" & r("descricao"))
		strResp = strResp & "</option>" & chr(13)
		r.MoveNext
		loop

	if Not ha_default then
		strResp = "<option selected value=''>&nbsp;</option>" & chr(13) & strResp
		end if
		
	forma_pagto_liberada_da_entrada_monta_itens_select_incluindo_default = strResp
	r.close
	set r=nothing
end function



' ____________________________________________
' FORMA_PAGTO_DA_PRESTACAO_MONTA_ITENS_SELECT
'
function forma_pagto_da_prestacao_monta_itens_select(byval id_default)
dim x, r, strResp, ha_default
	id_default = Trim("" & id_default)
	ha_default=False
	set r = cn.Execute("SELECT * FROM t_FORMA_PAGTO WHERE hab_prestacao=1 ORDER BY ordenacao")
	strResp = ""
	do while Not r.eof 
		x = Trim("" & r("id"))
		if (id_default<>"") And (id_default=x) then
			strResp = strResp & "<option selected"
			ha_default=True
		else
			strResp = strResp & "<option"
			end if
		strResp = strResp & " value='" & x & "'>"
		strResp = strResp & Trim("" & r("descricao"))
		strResp = strResp & "</option>" & chr(13)
		r.MoveNext
		loop

	if Not ha_default then
		strResp = "<option selected value=''>&nbsp;</option>" & chr(13) & strResp
		end if
		
	forma_pagto_da_prestacao_monta_itens_select = strResp
	r.close
	set r=nothing
end function



' _______________________________________________________________
' FORMA_PAGTO_DA_PRESTACAO_MONTA_ITENS_SELECT_INCLUINDO_DEFAULT
'
function forma_pagto_da_prestacao_monta_itens_select_incluindo_default(byval id_default)
dim x, s, r, strResp, ha_default
	id_default = Trim("" & id_default)
	ha_default=False
	s = "SELECT " & _
			"*" & _
		" FROM t_FORMA_PAGTO" & _
		" WHERE" & _
			" (id = " & id_default & ")" & _
			" OR " & _
			" (hab_prestacao = 1)" & _
		" ORDER BY" & _
			" ordenacao"
	set r = cn.Execute(s)
	strResp = ""
	do while Not r.eof 
		x = Trim("" & r("id"))
		if (id_default<>"") And (id_default=x) then
			strResp = strResp & "<option selected"
			ha_default=True
		else
			strResp = strResp & "<option"
			end if
		strResp = strResp & " value='" & x & "'>"
		strResp = strResp & Trim("" & r("descricao"))
		strResp = strResp & "</option>" & chr(13)
		r.MoveNext
		loop

	if Not ha_default then
		strResp = "<option selected value=''>&nbsp;</option>" & chr(13) & strResp
		end if
		
	forma_pagto_da_prestacao_monta_itens_select_incluindo_default = strResp
	r.close
	set r=nothing
end function



' ________________________________________________________________________________________
' FORMA_PAGTO_LIBERADA_DA_PRESTACAO_MONTA_ITENS_SELECT
' VERIFICA NA TABELA t_ORCAMENTISTA_E_INDICADOR_RESTRICAO_FORMA_PAGTO SE HÁ
' RESTRIÇÕES P/ O REFERIDO 'INDICADOR' P/ O TIPO DE CLIENTE (PF/PJ) ESPECIFICADO.
' NO CASO DE NÃO SER INFORMADO O INDICADOR OU TIPO DE CLIENTE, TODAS AS FORMAS DE
' PAGAMENTO SÃO RETORNADAS.
function forma_pagto_liberada_da_prestacao_monta_itens_select(byval id_default, byval id_orcamentista_e_indicador, byval tipo_cliente)
dim s, x, r, strResp, ha_default
	id_default = Trim("" & id_default)
	ha_default=False
	s = "SELECT " & _
			"*" & _
		" FROM t_FORMA_PAGTO" & _
		" WHERE" & _
			" (hab_prestacao=1)" & _
			" AND " & _
			"(" & _
				"id NOT IN " & _
					"(" & _
					"SELECT" & _
						" id_forma_pagto" & _
					" FROM t_ORCAMENTISTA_E_INDICADOR_RESTRICAO_FORMA_PAGTO tOIRFP" & _
					" WHERE" & _ 
						"(" & _
							"(id_orcamentista_e_indicador = '" & Trim(id_orcamentista_e_indicador) & "')" & _
							" OR " & _
							"(id_orcamentista_e_indicador = '" & ID_ORCAMENTISTA_E_INDICADOR_RESTRICAO_FP_TODOS & "')" & _
						")" & _
						" AND (tipo_cliente = '" & Trim(tipo_cliente) & "')" & _
						" AND (st_restricao_ativa <> 0)" & _
					")" & _
			")" & _
		" ORDER BY" & _
			" ordenacao"
	set r = cn.Execute(s)
	strResp = ""
	do while Not r.eof 
		x = Trim("" & r("id"))
		if (id_default<>"") And (id_default=x) then
			strResp = strResp & "<option selected"
			ha_default=True
		else
			strResp = strResp & "<option"
			end if
		strResp = strResp & " value='" & x & "'>"
		strResp = strResp & Trim("" & r("descricao"))
		strResp = strResp & "</option>" & chr(13)
		r.MoveNext
		loop

	if Not ha_default then
		strResp = "<option selected value=''>&nbsp;</option>" & chr(13) & strResp
		end if
		
	forma_pagto_liberada_da_prestacao_monta_itens_select = strResp
	r.close
	set r=nothing
end function



' _________________________________________________________________________________
' FORMA_PAGTO_LIBERADA_DA_PRESTACAO_MONTA_ITENS_SELECT_INCLUINDO_DEFAULT
' VERIFICA NA TABELA t_ORCAMENTISTA_E_INDICADOR_RESTRICAO_FORMA_PAGTO SE HÁ
' RESTRIÇÕES P/ O REFERIDO 'INDICADOR' P/ O TIPO DE CLIENTE (PF/PJ) ESPECIFICADO.
' NO CASO DE NÃO SER INFORMADO O INDICADOR OU TIPO DE CLIENTE, TODAS AS FORMAS DE
' PAGAMENTO SÃO RETORNADAS.
' ASSEGURA A INCLUSÃO DA OPÇÃO DEFAULT NA LISTA.
function forma_pagto_liberada_da_prestacao_monta_itens_select_incluindo_default(byval id_default, byval id_orcamentista_e_indicador, byval tipo_cliente)
dim s, x, r, strResp, ha_default
	id_default = Trim("" & id_default)
	ha_default=False
	s = "SELECT " & _
			"*" & _
		" FROM t_FORMA_PAGTO" & _
		" WHERE" & _
			" (id = " & id_default & ")" & _
			" OR " & _
			"(" & _
				"(hab_prestacao=1)" & _
				" AND " & _
				"(id NOT IN " & _
					"(" & _
					"SELECT" & _
						" id_forma_pagto" & _
					" FROM t_ORCAMENTISTA_E_INDICADOR_RESTRICAO_FORMA_PAGTO tOIRFP" & _
					" WHERE" & _ 
						"(" & _
							"(id_orcamentista_e_indicador = '" & Trim(id_orcamentista_e_indicador) & "')" & _
							" OR " & _
							"(id_orcamentista_e_indicador = '" & ID_ORCAMENTISTA_E_INDICADOR_RESTRICAO_FP_TODOS & "')" & _
						")" & _
						" AND (tipo_cliente = '" & Trim(tipo_cliente) & "')" & _
						" AND (st_restricao_ativa <> 0)" & _
					")" & _
				")" & _
			")" & _
		" ORDER BY" & _
			" ordenacao"
	set r = cn.Execute(s)
	strResp = ""
	do while Not r.eof 
		x = Trim("" & r("id"))
		if (id_default<>"") And (id_default=x) then
			strResp = strResp & "<option selected"
			ha_default=True
		else
			strResp = strResp & "<option"
			end if
		strResp = strResp & " value='" & x & "'>"
		strResp = strResp & Trim("" & r("descricao"))
		strResp = strResp & "</option>" & chr(13)
		r.MoveNext
		loop

	if Not ha_default then
		strResp = "<option selected value=''>&nbsp;</option>" & chr(13) & strResp
		end if
		
	forma_pagto_liberada_da_prestacao_monta_itens_select_incluindo_default = strResp
	r.close
	set r=nothing
end function



' _________________________________________________
' FORMA_PAGTO_DA_PARCELA_UNICA_MONTA_ITENS_SELECT
'
function forma_pagto_da_parcela_unica_monta_itens_select(byval id_default)
dim x, r, strResp, ha_default
	id_default = Trim("" & id_default)
	ha_default=False
	set r = cn.Execute("SELECT * FROM t_FORMA_PAGTO WHERE hab_parcela_unica=1 ORDER BY ordenacao")
	strResp = ""
	do while Not r.eof 
		x = Trim("" & r("id"))
		if (id_default<>"") And (id_default=x) then
			strResp = strResp & "<option selected"
			ha_default=True
		else
			strResp = strResp & "<option"
			end if
		strResp = strResp & " value='" & x & "'>"
		strResp = strResp & Trim("" & r("descricao"))
		strResp = strResp & "</option>" & chr(13)
		r.MoveNext
		loop

	if Not ha_default then
		strResp = "<option selected value=''>&nbsp;</option>" & chr(13) & strResp
		end if
		
	forma_pagto_da_parcela_unica_monta_itens_select = strResp
	r.close
	set r=nothing
end function



' ____________________________________________________________
' FORMA_PAGTO_DA_PARCELA_UNICA_MONTA_ITENS_SELECT_EC
' Rotina específica para ser usada pela loja do e-commerce
function forma_pagto_da_parcela_unica_monta_itens_select_EC(byval id_default)
dim x, r, strResp, ha_default
	id_default = Trim("" & id_default)
	ha_default=False
	set r = cn.Execute("SELECT * FROM t_FORMA_PAGTO WHERE (id IN (" & ID_FORMA_PAGTO_DEPOSITO & ")) ORDER BY ordenacao")
	strResp = ""
	do while Not r.eof 
		x = Trim("" & r("id"))
		if (id_default<>"") And (id_default=x) then
			strResp = strResp & "<option selected"
			ha_default=True
		else
			strResp = strResp & "<option"
			end if
		strResp = strResp & " value='" & x & "'>"
		strResp = strResp & Trim("" & r("descricao"))
		strResp = strResp & "</option>" & chr(13)
		r.MoveNext
		loop

	if Not ha_default then
		strResp = "<option selected value=''>&nbsp;</option>" & chr(13) & strResp
		end if
		
	forma_pagto_da_parcela_unica_monta_itens_select_EC = strResp
	r.close
	set r=nothing
end function



' ___________________________________________________________________
' FORMA_PAGTO_DA_PARCELA_UNICA_MONTA_ITENS_SELECT_INCLUINDO_DEFAULT
'
function forma_pagto_da_parcela_unica_monta_itens_select_incluindo_default(byval id_default)
dim x, s, r, strResp, ha_default
	id_default = Trim("" & id_default)
	ha_default=False
	s = "SELECT " & _
			"*" & _
		" FROM t_FORMA_PAGTO" & _
		" WHERE" & _
			" (id = " & id_default & ")" & _
			" OR " & _
			" (hab_parcela_unica = 1)" & _
		" ORDER BY" & _
			" ordenacao"
	set r = cn.Execute(s)
	strResp = ""
	do while Not r.eof 
		x = Trim("" & r("id"))
		if (id_default<>"") And (id_default=x) then
			strResp = strResp & "<option selected"
			ha_default=True
		else
			strResp = strResp & "<option"
			end if
		strResp = strResp & " value='" & x & "'>"
		strResp = strResp & Trim("" & r("descricao"))
		strResp = strResp & "</option>" & chr(13)
		r.MoveNext
		loop

	if Not ha_default then
		strResp = "<option selected value=''>&nbsp;</option>" & chr(13) & strResp
		end if
		
	forma_pagto_da_parcela_unica_monta_itens_select_incluindo_default = strResp
	r.close
	set r=nothing
end function



' ________________________________________________________________________________________
' FORMA_PAGTO_LIBERADA_DA_PARCELA_UNICA_MONTA_ITENS_SELECT
' VERIFICA NA TABELA t_ORCAMENTISTA_E_INDICADOR_RESTRICAO_FORMA_PAGTO SE HÁ
' RESTRIÇÕES P/ O REFERIDO 'INDICADOR' P/ O TIPO DE CLIENTE (PF/PJ) ESPECIFICADO.
' NO CASO DE NÃO SER INFORMADO O INDICADOR OU TIPO DE CLIENTE, TODAS AS FORMAS DE
' PAGAMENTO SÃO RETORNADAS.
function forma_pagto_liberada_da_parcela_unica_monta_itens_select(byval id_default, byval id_orcamentista_e_indicador, byval tipo_cliente)
dim s, x, r, strResp, ha_default
	id_default = Trim("" & id_default)
	ha_default=False
	s = "SELECT " & _
			"*" & _
		" FROM t_FORMA_PAGTO" & _
		" WHERE" & _
			" (hab_parcela_unica=1)" & _
			" AND " & _
			"(" & _
				"id NOT IN " & _
					"(" & _
					"SELECT" & _
						" id_forma_pagto" & _
					" FROM t_ORCAMENTISTA_E_INDICADOR_RESTRICAO_FORMA_PAGTO tOIRFP" & _
					" WHERE" & _ 
						"(" & _
							"(id_orcamentista_e_indicador = '" & Trim(id_orcamentista_e_indicador) & "')" & _
							" OR " & _
							"(id_orcamentista_e_indicador = '" & ID_ORCAMENTISTA_E_INDICADOR_RESTRICAO_FP_TODOS & "')" & _
						")" & _
						" AND (tipo_cliente = '" & Trim(tipo_cliente) & "')" & _
						" AND (st_restricao_ativa <> 0)" & _
					")" & _
			")" & _
		" ORDER BY" & _
			" ordenacao"
	set r = cn.Execute(s)
	strResp = ""
	do while Not r.eof 
		x = Trim("" & r("id"))
		if (id_default<>"") And (id_default=x) then
			strResp = strResp & "<option selected"
			ha_default=True
		else
			strResp = strResp & "<option"
			end if
		strResp = strResp & " value='" & x & "'>"
		strResp = strResp & Trim("" & r("descricao"))
		strResp = strResp & "</option>" & chr(13)
		r.MoveNext
		loop

	if Not ha_default then
		strResp = "<option selected value=''>&nbsp;</option>" & chr(13) & strResp
		end if
		
	forma_pagto_liberada_da_parcela_unica_monta_itens_select = strResp
	r.close
	set r=nothing
end function



' _________________________________________________________________________________
' FORMA_PAGTO_LIBERADA_DA_PARCELA_UNICA_MONTA_ITENS_SELECT_INCLUINDO_DEFAULT
' VERIFICA NA TABELA t_ORCAMENTISTA_E_INDICADOR_RESTRICAO_FORMA_PAGTO SE HÁ
' RESTRIÇÕES P/ O REFERIDO 'INDICADOR' P/ O TIPO DE CLIENTE (PF/PJ) ESPECIFICADO.
' NO CASO DE NÃO SER INFORMADO O INDICADOR OU TIPO DE CLIENTE, TODAS AS FORMAS DE
' PAGAMENTO SÃO RETORNADAS.
' ASSEGURA A INCLUSÃO DA OPÇÃO DEFAULT NA LISTA.
function forma_pagto_liberada_da_parcela_unica_monta_itens_select_incluindo_default(byval id_default, byval id_orcamentista_e_indicador, byval tipo_cliente)
dim s, x, r, strResp, ha_default
	id_default = Trim("" & id_default)
	ha_default=False
	s = "SELECT " & _
			"*" & _
		" FROM t_FORMA_PAGTO" & _
		" WHERE" & _
			" (id = " & id_default & ")" & _
			" OR " & _
			"(" & _
				"(hab_parcela_unica=1)" & _
				" AND " & _
				"(id NOT IN " & _
					"(" & _
					"SELECT" & _
						" id_forma_pagto" & _
					" FROM t_ORCAMENTISTA_E_INDICADOR_RESTRICAO_FORMA_PAGTO tOIRFP" & _
					" WHERE" & _ 
						"(" & _
							"(id_orcamentista_e_indicador = '" & Trim(id_orcamentista_e_indicador) & "')" & _
							" OR " & _
							"(id_orcamentista_e_indicador = '" & ID_ORCAMENTISTA_E_INDICADOR_RESTRICAO_FP_TODOS & "')" & _
						")" & _
						" AND (tipo_cliente = '" & Trim(tipo_cliente) & "')" & _
						" AND (st_restricao_ativa <> 0)" & _
					")" & _
				")" & _
			")" & _
		" ORDER BY" & _
			" ordenacao"
	set r = cn.Execute(s)
	strResp = ""
	do while Not r.eof 
		x = Trim("" & r("id"))
		if (id_default<>"") And (id_default=x) then
			strResp = strResp & "<option selected"
			ha_default=True
		else
			strResp = strResp & "<option"
			end if
		strResp = strResp & " value='" & x & "'>"
		strResp = strResp & Trim("" & r("descricao"))
		strResp = strResp & "</option>" & chr(13)
		r.MoveNext
		loop

	if Not ha_default then
		strResp = "<option selected value=''>&nbsp;</option>" & chr(13) & strResp
		end if
		
	forma_pagto_liberada_da_parcela_unica_monta_itens_select_incluindo_default = strResp
	r.close
	set r=nothing
end function



' ____________________________________________
' LOJA_DO_ORCAMENTISTA_MONTA_ITENS_SELECT
'
function loja_do_orcamentista_monta_itens_select(byval id_default)
dim x, r, strResp, ha_default
	id_default = Trim("" & id_default)
	ha_default=False
	set r = cn.Execute("SELECT CONVERT(smallint,loja) AS numero_loja, t_LOJA.* FROM t_LOJA ORDER BY numero_loja")
	strResp = ""
	do while Not r.eof 
		x = Trim("" & r("loja"))
		if (id_default<>"") And (converte_numero(id_default)=converte_numero(x)) then
			strResp = strResp & "<option selected"
			ha_default=True
		else
			strResp = strResp & "<option"
			end if
		strResp = strResp & " value='" & x & "'>"
		strResp = strResp & x & " - " & Trim("" & r("nome"))
		strResp = strResp & "</option>" & chr(13)
		r.MoveNext
		loop

	if Not ha_default then
		strResp = "<option selected value=''>&nbsp;</option>" & chr(13) & strResp
		end if
		
	loja_do_orcamentista_monta_itens_select = strResp
	r.close
	set r=nothing
end function



' ____________________________________________
' LOJAS_MONTA_ITENS_SELECT
'
function lojas_monta_itens_select(byval id_default)
dim x, r, strResp, ha_default
	id_default = Trim("" & id_default)
	ha_default=False
	set r = cn.Execute("SELECT CONVERT(smallint,loja) AS numero_loja, t_LOJA.* FROM t_LOJA ORDER BY numero_loja")
	strResp = ""
	do while Not r.eof 
		x = Trim("" & r("loja"))
		if (id_default<>"") And (converte_numero(id_default)=converte_numero(x)) then
			strResp = strResp & "<option selected"
			ha_default=True
		else
			strResp = strResp & "<option"
			end if
		strResp = strResp & " value='" & x & "'>"
		strResp = strResp & x & " - " & Trim("" & r("nome"))
		strResp = strResp & "</option>" & chr(13)
		r.MoveNext
		loop

	if Not ha_default then
		strResp = "<option selected value=''>&nbsp;</option>" & chr(13) & strResp
		end if
		
	lojas_monta_itens_select = strResp
	r.close
	set r=nothing
end function



' ____________________________________________
' VENDEDOR_DO_INDICADOR_MONTA_ITENS_SELECT
'
function vendedor_do_indicador_monta_itens_select(byval id_default)
dim x, r, strResp, ha_default, strSql, strWhere
	id_default = Trim("" & id_default)
	ha_default=False
	strWhere="(vendedor_loja = 1) AND (bloqueado = 0)"
	if id_default <> "" then 
		strWhere= "(" & strWhere & ")" & _
						" OR " & _
				  "(usuario = '" & id_default & "')"
		end if
		
	strSql="SELECT" & _
				" usuario," & _
				" nome," & _
				" nome_iniciais_em_maiusculas" & _
			" FROM t_USUARIO" & _
			" WHERE " & _
				strWhere & _
			" ORDER BY" & _
				" usuario"
	set r = cn.Execute(strSql)
	strResp = ""
	do while Not r.eof 
		x = UCase(Trim("" & r("usuario")))
		if (id_default<>"") And (id_default=x) then
			strResp = strResp & "<option selected"
			ha_default=True
		else
			strResp = strResp & "<option"
			end if
		strResp = strResp & " value='" & x & "'>"
		strResp = strResp & x & " - " & Trim("" & r("nome_iniciais_em_maiusculas"))
		strResp = strResp & "</option>" & chr(13)
		r.MoveNext
		loop

	if Not ha_default then
		strResp = "<option selected value=''>&nbsp;</option>" & chr(13) & strResp
		end if
		
	vendedor_do_indicador_monta_itens_select = strResp
	r.close
	set r=nothing
end function



' _______________________________________________________
' VENDEDOR_DO_INDICADOR_DESTA_LOJA_MONTA_ITENS_SELECT
'
function vendedor_do_indicador_desta_loja_monta_itens_select(byval loja, byval id_default)
dim x, r, strSql, strResp, ha_default
	loja = Trim("" & loja)
	id_default = Trim("" & id_default)
	ha_default=False
	strSql = "SELECT" & _
				" usuario," & _
				" nome," & _
				" nome_iniciais_em_maiusculas" & _
			" FROM t_USUARIO" & _
			" WHERE" & _
				"(usuario = '" & id_default & "')" & _
					" OR " & _
				"(" & _
					" (vendedor_loja = 1)" & _
					" AND (bloqueado = 0)" & _
					" AND" & _
						"(usuario IN " & _
							"(" & _
								"SELECT DISTINCT " & _
									"usuario" & _
								" FROM t_USUARIO_X_LOJA" & _
								" WHERE" & _
									" (CONVERT(smallint, loja) = " & loja & ")" & _
							")" & _
						")" & _
					")" & _
			" ORDER BY" & _
				" usuario"
	set r = cn.Execute(strSql)
	strResp = ""
	do while Not r.eof 
		x = UCase(Trim("" & r("usuario")))
		if (id_default<>"") And (id_default=x) then
			strResp = strResp & "<option selected"
			ha_default=True
		else
			strResp = strResp & "<option"
			end if
		strResp = strResp & " value='" & x & "'>"
		strResp = strResp & x & " - " & Trim("" & r("nome_iniciais_em_maiusculas"))
		strResp = strResp & "</option>" & chr(13)
		r.MoveNext
		loop

	if Not ha_default then
		strResp = "<option selected value=''>&nbsp;</option>" & chr(13) & strResp
		end if
		
	vendedor_do_indicador_desta_loja_monta_itens_select = strResp
	r.close
	set r=nothing
end function



' __________________________________________________
' M I D I A _ M O N T A _ I T E N S _ S E L E C T
'
function midia_monta_itens_select(byval id_default)
dim x,r,s,ha_default
	id_default = Trim("" & id_default)
	ha_default=False
	set r = cn.Execute("SELECT * FROM t_MIDIA WHERE indisponivel=0 ORDER BY apelido")
	s= ""
	do while Not r.eof 
		x = Trim("" & r("id"))
		if (id_default<>"") And (id_default=x) then
			s = s & "<option selected"
			ha_default=True
		else
			s = s & "<option"
			end if
		s = s & " value='" & x & "'>"
		s = s & Trim("" & r("apelido"))
		s = s & "</option>" & chr(13)
		r.MoveNext
		loop

	if Not ha_default then
		s = "<option selected value=''>&nbsp;</option>" & chr(13) & s
		end if
		
	midia_monta_itens_select = s
	r.close
	set r=nothing
end function



' ____________________________________________
' AUTORIZADOR_DESCONTO_MONTA_ITENS_SELECT
'
function autorizador_desconto_monta_itens_select(byval id_default)
dim x, r, strSql, strResp, ha_default
	id_default = Trim("" & id_default)
	ha_default=False
	strSql = "SELECT" & _
				" usuario," & _
				" nome," & _
				" nome_iniciais_em_maiusculas" & _
			" FROM t_USUARIO" & _
			" WHERE" & _
				" (bloqueado = 0)" & _
				" AND (" & _
					"usuario IN " & _
						"(" & _
							"SELECT" & _
								" tPU.usuario" & _
							" FROM t_PERFIL_X_USUARIO tPU" & _
								" INNER JOIN t_PERFIL tP ON (tPU.id_perfil=tP.id)" & _
								" INNER JOIN t_PERFIL_ITEM tPI ON (tP.id=tPI.id_perfil)" & _
							" WHERE" & _
								" (tPI.id_operacao = " & Cstr(OP_CEN_AUTORIZA_SENHA_DESCONTO) & ")" & _
						")" & _
					  ")" & _
			" ORDER BY" & _
				" usuario"
	set r = cn.Execute(strSql)
	strResp = ""
	do while Not r.eof 
		x = UCase(Trim("" & r("usuario")))
		if (id_default<>"") And (id_default=x) then
			strResp = strResp & "<option selected"
			ha_default=True
		else
			strResp = strResp & "<option"
			end if
		strResp = strResp & " value='" & x & "'>"
		strResp = strResp & x & " - " & Trim("" & r("nome_iniciais_em_maiusculas"))
		strResp = strResp & "</option>" & chr(13)
		r.MoveNext
		loop

	if Not ha_default then
		strResp = "<option selected value=''>&nbsp;</option>" & chr(13) & strResp
		end if
		
	autorizador_desconto_monta_itens_select = strResp
	r.close
	set r=nothing
end function



' ___________________________________________________________
' AUTORIZADOR_DESCONTO_CADASTRADO_NA_LOJA_MONTA_ITENS_SELECT
'
function autorizador_desconto_cadastrado_na_loja_monta_itens_select(byval loja, byval id_default)
dim x, r, strSql, strResp, ha_default
	id_default = Trim("" & id_default)
	loja = Trim("" & loja)
	ha_default=False
	strSql = "SELECT" & _
				" usuario," & _
				" nome," & _
				" nome_iniciais_em_maiusculas" & _
			" FROM t_USUARIO" & _
			" WHERE" & _
				" (bloqueado = 0)" & _
				" AND (" & _
					"usuario IN " & _
						"(" & _
							"SELECT" & _
								" tUL.usuario" & _
							" FROM t_USUARIO_X_LOJA tUL" & _
							" WHERE" & _
								" (CONVERT(smallint, tUL.loja) = " & loja & ")" & _
						")" & _
					  ")" & _
			" ORDER BY" & _
				" usuario"
	set r = cn.Execute(strSql)
	strResp = ""
	do while Not r.eof 
		x = UCase(Trim("" & r("usuario")))
		if (id_default<>"") And (id_default=x) then
			strResp = strResp & "<option selected"
			ha_default=True
		else
			strResp = strResp & "<option"
			end if
		strResp = strResp & " value='" & x & "'>"
		strResp = strResp & x & " - " & Trim("" & r("nome_iniciais_em_maiusculas"))
		strResp = strResp & "</option>" & chr(13)
		r.MoveNext
		loop

	if Not ha_default then
		strResp = "<option selected value=''>&nbsp;</option>" & chr(13) & strResp
		end if
		
	autorizador_desconto_cadastrado_na_loja_monta_itens_select = strResp
	r.close
	set r=nothing
end function



' ____________________________________________
' CAPTADORES MONTA ITENS SELECT
'
function captadores_monta_itens_select(byval id_default)
dim x, r, strResp, ha_default
	id_default = Trim("" & id_default)
	ha_default=False
	set r = cn.Execute("SELECT usuario, nome, nome_iniciais_em_maiusculas FROM t_USUARIO ORDER BY usuario")
	strResp = ""
	do while Not r.eof 
		x = UCase(Trim("" & r("usuario")))
		if (id_default<>"") And (id_default=x) then
			strResp = strResp & "<option selected"
			ha_default=True
		else
			strResp = strResp & "<option"
			end if
		strResp = strResp & " value='" & x & "'>"
		strResp = strResp & x & " - " & Trim("" & r("nome_iniciais_em_maiusculas"))
		strResp = strResp & "</option>" & chr(13)
		r.MoveNext
		loop

	if Not ha_default then
		strResp = "<option selected value=''>&nbsp;</option>" & chr(13) & strResp
		end if
		
	captadores_monta_itens_select = strResp
	r.close
	set r=nothing
end function



' ____________________________________________________________________________
' INDICADORES MONTA ITENS SELECT
' LEMBRE-SE: O ORÇAMENTISTA É CONSIDERADO AUTOMATICAMENTE UM INDICADOR!!
function indicadores_monta_itens_select(byval id_default)
dim x, r, strResp, ha_default
	id_default = Trim("" & id_default)
	ha_default=False
	set r = cn.Execute("SELECT apelido, razao_social_nome_iniciais_em_maiusculas FROM t_ORCAMENTISTA_E_INDICADOR ORDER BY apelido")
	strResp = ""
	do while Not r.eof 
		x = UCase(Trim("" & r("apelido")))
		if (id_default<>"") And (id_default=x) then
			strResp = strResp & "<option selected"
			ha_default=True
		else
			strResp = strResp & "<option"
			end if
		strResp = strResp & " value='" & x & "'>"
		strResp = strResp & x & " - " & Trim("" & r("razao_social_nome_iniciais_em_maiusculas"))
		strResp = strResp & "</option>" & chr(13)
		r.MoveNext
		loop

	if Not ha_default then
		strResp = "<option selected value=''>&nbsp;</option>" & chr(13) & strResp
		end if
		
	indicadores_monta_itens_select = strResp
	r.close
	set r=nothing
end function



' ____________________________________________
' VENDEDORES_MONTA_ITENS_SELECT
'
function vendedores_monta_itens_select(byval id_default)
dim x, r, strResp, ha_default, strSql, strWhere
	id_default = Trim("" & id_default)
	ha_default=False
	strWhere="(vendedor_loja = 1)"
	if id_default <> "" then 
		strWhere= "(" & strWhere & ")" & _
						" OR " & _
				  "(usuario = '" & id_default & "')"
		end if
		
	strSql="SELECT" & _
				" usuario," & _
				" nome," & _
				" nome_iniciais_em_maiusculas" & _
			" FROM t_USUARIO" & _
			" WHERE " & _
				strWhere & _
			" ORDER BY" & _
				" usuario"
	set r = cn.Execute(strSql)
	strResp = ""
	do while Not r.eof 
		x = UCase(Trim("" & r("usuario")))
		if (id_default<>"") And (id_default=x) then
			strResp = strResp & "<option selected"
			ha_default=True
		else
			strResp = strResp & "<option"
			end if
		strResp = strResp & " value='" & x & "'>"
		strResp = strResp & x & " - " & Trim("" & r("nome_iniciais_em_maiusculas"))
		strResp = strResp & "</option>" & chr(13)
		r.MoveNext
		loop

	if Not ha_default then
		strResp = "<option selected value=''>&nbsp;</option>" & chr(13) & strResp
		end if
		
	vendedores_monta_itens_select = strResp
	r.close
	set r=nothing
end function



' _____________________________________________________________
' APELIDO EMPRESA NFE EMITENTE MONTA ITENS SELECT
' Usado p/ montar a lista de opções com as empresas que constam
' da tabela t_NFe_EMITENTE exibindo apenas o apelido da empresa.
function apelido_empresa_nfe_emitente_monta_itens_select(byval id_default)
dim x, r, strSql, strResp, ha_default
	id_default = Trim("" & id_default)
	ha_default=False
	strSql = "SELECT" & _
				" id," & _
				" apelido," & _
				" razao_social" & _
			" FROM t_NFe_EMITENTE" & _
			" WHERE" & _
				" (st_ativo <> 0)"
	
	if id_default <> "" then
	'	GARANTE QUE A OPÇÃO DEFAULT CONSTE NA LISTA MESMO QUE NÃO ESTEJA MAIS ATIVA NO CADASTRO, ISSO É IMPORTANTE EM PÁGINAS DE EDIÇÃO
		strSql = strSql & _
				" OR " & _
				"(id = " & id_default & ")"
		end if

	strSql = strSql & _
			" ORDER BY" & _
				" id"
	set r = cn.Execute(strSql)
	strResp = ""
	do while Not r.eof 
		x = Trim("" & r("id"))
		if (id_default<>"") And (id_default<>"0") And (id_default=x) then
			strResp = strResp & "<option selected"
			ha_default=True
		else
			strResp = strResp & "<option"
			end if
		strResp = strResp & " value='" & x & "'>"
		strResp = strResp & Trim("" & r("apelido"))
		strResp = strResp & "</option>" & chr(13)
		r.MoveNext
		loop
	
'	SE NÃO HÁ NENHUM ITEM DEFAULT, INCLUI UM ITEM EM BRANCO P/ SER O DEFAULT
	if Not ha_default then
		strResp = "<option selected value=''>&nbsp;</option>" & chr(13) & strResp		
	else
    	strResp = "<option  value=''>&nbsp;</option>" & chr(13) & strResp
    end if

	apelido_empresa_nfe_emitente_monta_itens_select = strResp
	r.close
	set r=nothing
end function



' _____________________________________________________________
' WMS APELIDO EMPRESA NFE EMITENTE MONTA ITENS SELECT
' Usado p/ montar a lista de opções com as empresas que constam
' da tabela t_NFe_EMITENTE e que estejam habilitadas para o
' o controle do estoque.
' Exibe apenas o apelido da empresa.
function wms_apelido_empresa_nfe_emitente_monta_itens_select(byval id_default)
dim x, r, strSql, strResp, ha_default
	id_default = Trim("" & id_default)
	ha_default=False
	strSql = "SELECT" & _
				" id," & _
				" apelido," & _
				" razao_social" & _
			" FROM t_NFe_EMITENTE" & _
			" WHERE" & _
				" (st_ativo <> 0)" & _
				" AND (st_habilitado_ctrl_estoque <> 0)"
	
	if id_default <> "" then
	'	GARANTE QUE A OPÇÃO DEFAULT CONSTE NA LISTA MESMO QUE NÃO ESTEJA MAIS ATIVA NO CADASTRO, ISSO É IMPORTANTE EM PÁGINAS DE EDIÇÃO
		strSql = strSql & _
				" OR " & _
				"(id = " & id_default & ")"
		end if

	strSql = strSql & _
			" ORDER BY" & _
				" id"
	set r = cn.Execute(strSql)
	strResp = ""
	do while Not r.eof 
		x = Trim("" & r("id"))
		if (id_default<>"") And (id_default<>"0") And (id_default=x) then
			strResp = strResp & "<option selected"
			ha_default=True
		else
			strResp = strResp & "<option"
			end if
		strResp = strResp & " value='" & x & "'>"
		strResp = strResp & Trim("" & r("apelido"))
		strResp = strResp & "</option>" & chr(13)
		r.MoveNext
		loop
	
'	SE NÃO HÁ NENHUM ITEM DEFAULT, INCLUI UM ITEM EM BRANCO P/ SER O DEFAULT
	if Not ha_default then
		strResp = "<option selected value=''>&nbsp;</option>" & chr(13) & strResp		
	else
    	strResp = "<option  value=''>&nbsp;</option>" & chr(13) & strResp
    end if

	wms_apelido_empresa_nfe_emitente_monta_itens_select = strResp
	r.close
	set r=nothing
end function



' _____________________________________________________________
' NFE EMITENTE MONTA ITENS SELECT
' Usado p/ montar a lista de opções do emitente da NFe,
' incluindo a opção 'Cliente' p/ situações em que a NFe de
' retorno é uma nota do cliente.
function nfe_emitente_monta_itens_select(byval id_default)
dim x, r, strSql, strResp, ha_default
	id_default = Trim("" & id_default)
	ha_default=False
	strSql = "SELECT" & _
				" id," & _
				" apelido," & _
				" razao_social" & _
			" FROM t_NFe_EMITENTE" & _
			" ORDER BY" & _
				" id"
	set r = cn.Execute(strSql)
	strResp = ""
	do while Not r.eof 
		x = Trim("" & r("id"))
		if (id_default<>"") And (id_default<>"0") And (id_default=x) then
			strResp = strResp & "<option selected"
			ha_default=True
		else
			strResp = strResp & "<option"
			end if
		strResp = strResp & " value='" & x & "'>"
		strResp = strResp & Trim("" & r("id")) & " - " & "(" & Trim("" & r("apelido")) & ") " & Trim("" & r("razao_social"))
		strResp = strResp & "</option>" & chr(13)
		r.MoveNext
		loop

'	NF EMITIDA PELO PRÓPRIO CLIENTE
	x = Trim("" & COD_NFE_EMITENTE__CLIENTE)
	if (id_default<>"") And (id_default<>"0") And (id_default=x) then
		strResp = strResp & "<option selected"
		ha_default=True
	else
		strResp = strResp & "<option"
		end if

	strResp = strResp & _
			  " value='" & x & "'>" & _
			  "CLIENTE" & _
			  "</option>" & chr(13)
	
'	SE NÃO HÁ NENHUM ITEM DEFAULT, INCLUI UM ITEM EM BRANCO P/ SER O DEFAULT
	if Not ha_default then
		strResp = "<option selected value=''>&nbsp;</option>" & chr(13) & strResp		
	else
    	strResp = "<option  value=''>&nbsp;</option>" & chr(13) & strResp
    end if

	nfe_emitente_monta_itens_select = strResp
	r.close
	set r=nothing
end function



' ____________________________________________
' UF_MONTA_ITENS_SELECT
'
function UF_monta_itens_select(byval id_default)
dim strResp, ha_default, strListaUF, strUF, vUF, intContador
	id_default = UCase(Trim("" & id_default))
	ha_default=False
	strListaUF="AC|AL|AM|AP|BA|CE|DF|ES|GO|MA|MG|MS|MT|PA|PB|PE|PI|PR|RJ|RN|RO|RR|RS|SC|SE|SP|TO"
	vUF=Split(strListaUF,"|")
	for intContador=LBound(vUF) to UBound(vUF)
		strUF = vUF(intContador)
		if (id_default<>"") And (id_default=strUF) then
			strResp = strResp & "<option selected"
			ha_default=True
		else
			strResp = strResp & "<option"
			end if
		strResp = strResp & " value='" & strUF & "'>"
		strResp = strResp & strUF
		strResp = strResp & "</option>" & chr(13)
		next

	if Not ha_default then
		strResp = "<option selected value=''>&nbsp;</option>" & chr(13) & strResp
	else
		strResp = "<option value=''>&nbsp;</option>" & chr(13) & strResp
		end if
		
	UF_monta_itens_select = strResp
end function



' _____________________________________________
' LOJA_TROCA_RAPIDA_MONTA_ITENS_SELECT
'
function loja_troca_rapida_monta_itens_select(byval strUsuario, byval id_default)
dim x, r, strSql, strResp, ha_default
	strUsuario = Trim("" & strUsuario)
	if strUsuario = "" then exit function
	id_default = Trim("" & id_default)
	ha_default=False
	
	if id_default = "" then
		strSql = "SELECT" & _
					" tUL.loja, tL.nome" & _
				" FROM t_USUARIO_X_LOJA tUL" & _
					" INNER JOIN t_LOJA tL ON (tUL.loja=tL.loja)" & _
				" WHERE" & _
					" (usuario = '" & strUsuario & "')" & _
				" ORDER BY" & _
					" tUL.loja"
	else
	'	LEMBRE-SE: O USUÁRIO QUE TEM PERMISSÃO DE ACESSO A TODAS AS LOJAS PODE
	'	ACESSAR UMA LOJA QUE NÃO ESTÁ CADASTRADA EM t_USUARIO_X_LOJA
		strSql = "SELECT DISTINCT" & _
					" loja, nome" & _
				" FROM " & _ 
					"(" & _ 
						"SELECT" & _
							" tUL.loja, tL.nome" & _
						" FROM t_USUARIO_X_LOJA tUL" & _
							" INNER JOIN t_LOJA tL ON (tUL.loja=tL.loja)" & _
						" WHERE" & _
							" (usuario = '" & strUsuario & "')" & _
						" UNION " & _
						"SELECT" & _
							" loja, nome" & _
						" FROM t_LOJA" & _
						" WHERE" & _
							" (CONVERT(smallint, loja) = " & id_default & ")" & _
					") t__AUX" & _
				" ORDER BY" & _
					" loja"
		end if
		
	set r = cn.Execute(strSql)
	strResp = ""
	do while Not r.eof 
		x = Trim("" & r("loja"))
		if (id_default<>"") And (id_default=x) then
			strResp = strResp & "<option selected"
			ha_default=True
		else
			strResp = strResp & "<option"
			end if
		strResp = strResp & " value='" & x & "'>"
		strResp = strResp & "&nbsp;" & Trim("" & r("nome")) & "&nbsp;&nbsp;&nbsp;"
		strResp = strResp & "</option>" & chr(13)
		r.MoveNext
		loop

	if Not ha_default then
		strResp = "<option selected value=''>&nbsp;</option>" & chr(13) & strResp
		end if
		
	loja_troca_rapida_monta_itens_select = strResp
	r.close
	set r=nothing
end function



' _____________________________________________
' BANCO_MONTA_ITENS_SELECT
'
function banco_monta_itens_select(byval id_default)
dim x, r, strResp, ha_default
	id_default = Trim("" & id_default)
	ha_default=False
	set r = cn.Execute("SELECT * FROM t_BANCO ORDER BY Convert(smallint,codigo)")
	strResp = ""
	do while Not r.eof 
		x = Trim("" & r("codigo"))
		if (id_default<>"") And (converte_numero(id_default)=converte_numero(x)) then
			strResp = strResp & "<option selected"
			ha_default=True
		else
			strResp = strResp & "<option"
			end if
		strResp = strResp & " value='" & x & "'>"
		strResp = strResp & Trim("" & r("codigo")) & " - " & Trim("" & r("descricao"))
		strResp = strResp & "</option>" & chr(13)
		r.MoveNext
		loop

	if Not ha_default then
		strResp = "<option selected value=''>&nbsp;</option>" & chr(13) & strResp
	else
		strResp = "<option value=''>&nbsp;</option>" & chr(13) & strResp
		end if
		
	banco_monta_itens_select = strResp
	r.close
	set r=nothing
end function



' ___________________________________________________________________________________
' NIVEL_ACESSO_BLOCO_NOTAS_PEDIDO_MONTA_ITENS_SELECT
'
function nivel_acesso_bloco_notas_pedido_monta_itens_select(byval codigo_default, byval max_cod_nivel_a_listar)
dim strResp, ha_default, strListaItem, strItem, vItem, intContador, blnPular
	codigo_default = Trim("" & codigo_default)
	max_cod_nivel_a_listar = Trim("" & max_cod_nivel_a_listar)
	ha_default=False
	strListaItem = COD_NIVEL_ACESSO_BLOCO_NOTAS_PEDIDO__PUBLICO & "|" & COD_NIVEL_ACESSO_BLOCO_NOTAS_PEDIDO__RESTRITO & "|" & COD_NIVEL_ACESSO_BLOCO_NOTAS_PEDIDO__SIGILOSO
	vItem=Split(strListaItem,"|")

	for intContador=LBound(vItem) to UBound(vItem)
		strItem = vItem(intContador)
		
		blnPular=False
		if converte_numero(max_cod_nivel_a_listar) <> converte_numero(COD_NIVEL_ACESSO_BLOCO_NOTAS_PEDIDO__ILIMITADO) then
			if converte_numero(strItem) > converte_numero(max_cod_nivel_a_listar) then blnPular=True
			end if
		
		if Not blnPular then
			if (codigo_default<>"") And (codigo_default=strItem) then
				strResp = strResp & "<option selected"
				ha_default=True
			else
				strResp = strResp & "<option"
				end if
			strResp = strResp & " value='" & strItem & "'>"
			strResp = strResp & nivel_acesso_bloco_notas_pedido_descricao(strItem)
			strResp = strResp & "</option>" & chr(13)
			end if
		next

	if Not ha_default then
		strResp = "<option selected value=''>&nbsp;</option>" & chr(13) & strResp
	else
		strResp = "<option value=''>&nbsp;</option>" & chr(13) & strResp
		end if
		
	nivel_acesso_bloco_notas_pedido_monta_itens_select = strResp
end function


' ___________________________________________________________________________________
' NIVEL_ACESSO_CHAMADO_PEDIDO_MONTA_ITENS_SELECT
'
function nivel_acesso_chamado_pedido_monta_itens_select(byval codigo_default, byval max_cod_nivel_a_listar, byval exibe_publico_geral)
dim strResp, ha_default, strListaItem, strItem, vItem, intContador, blnPular
	codigo_default = Trim("" & codigo_default)
	max_cod_nivel_a_listar = Trim("" & max_cod_nivel_a_listar)
    exibe_publico_geral = CBool(Trim("" & exibe_publico_geral))
	ha_default=False
	strListaItem = COD_NIVEL_ACESSO_CHAMADO_PEDIDO__PUBLICO & "|" & COD_NIVEL_ACESSO_CHAMADO_PEDIDO__PUBLICO_INTERNO & "|" & COD_NIVEL_ACESSO_CHAMADO_PEDIDO__RESTRITO & "|" & _
                   COD_NIVEL_ACESSO_CHAMADO_PEDIDO__SIGILOSO
	vItem=Split(strListaItem,"|")

	for intContador=LBound(vItem) to UBound(vItem)
		strItem = vItem(intContador)
		
		blnPular=False
		if converte_numero(max_cod_nivel_a_listar) <> converte_numero(COD_NIVEL_ACESSO_CHAMADO_PEDIDO__ILIMITADO) then
			if converte_numero(strItem) > converte_numero(max_cod_nivel_a_listar) then blnPular=True
			end if

        if converte_numero(strItem) = converte_numero(COD_NIVEL_ACESSO_CHAMADO_PEDIDO__PUBLICO) then
            if exibe_publico_geral = False then blnPular = True
            end if
		
		if Not blnPular then
			if (codigo_default<>"") And (codigo_default=strItem) then
				strResp = strResp & "<option selected"
				ha_default=True
			else
				strResp = strResp & "<option"
				end if
			strResp = strResp & " value='" & strItem & "'>"
			strResp = strResp & nivel_acesso_chamado_pedido_descricao(strItem)
			strResp = strResp & "</option>" & chr(13)
			end if
		next

	if Not ha_default then
		strResp = "<option selected value=''>&nbsp;</option>" & chr(13) & strResp
	else
		strResp = "<option value=''>&nbsp;</option>" & chr(13) & strResp
		end if
		
	nivel_acesso_chamado_pedido_monta_itens_select = strResp
end function



' ___________________________________________________________________________________
' FORMATA O Nº SÉRIE DA NFe ADICIONANDO OS ZEROS À ESQUERDA
'
function NFeFormataSerieNF(ByVal numeroSerieNF)
dim s_resp
	s_resp = Trim(Cstr(numeroSerieNF))
	do while Len(s_resp) < 3
		s_resp = "0" & s_resp
		loop
	NFeFormataSerieNF = s_resp
end function



' ___________________________________________________________________________________
' FORMATA O Nº DA NFe ADICIONANDO OS ZEROS À ESQUERDA
'
function NFeFormataNumeroNF(ByVal numeroNF)
dim s_resp
	s_resp = Trim(Cstr(numeroNF))
	do while Len(s_resp) < 9
		s_resp = "0" & s_resp
		loop
	NFeFormataNumeroNF = s_resp
end function

' ___________________________________________________________________________________
' MONTA LINK PARA DANFE
' - traz a nota fiscal de venda de um pedido, se houver
function monta_link_para_DANFE(byval pedido, byval max_dias_emissao, byval html_elemento_interno_anchor)
dim s, strSerieNFe, strNumeroNFe
dim strArqDanfe, strArqDanfeAux, strArqDanfeCompleto, strPathArqDanfe, strLinkDanfe
dim intIdNFeEmitente, lngSerieNFe, lngNumeroNFe, lngQtdeDiasEmissao, lngMaxDiasEmissao
dim blnArqDanfeExiste
dim fso
dim r
dim tNE, tNFE
dim dbcNFe
dim strNfeT1ServidorBd, strNfeT1NomeBd, strNfeT1UsuarioBd, strNfeT1SenhaCriptografadaBd
dim chave
dim senha_decodificada
dim strFlagCancelada
dim blnQtdeDiasOK
dim rP, rNfeEmitente, iCfg
	
	monta_link_para_DANFE = ""
	
	pedido = Trim("" & pedido)
	if pedido = "" then exit function
	
	lngMaxDiasEmissao = 0
	if IsNumeric(max_dias_emissao) then lngMaxDiasEmissao = CLng(max_dias_emissao)
	
'	LOCALIZA A ÚLTIMA NFe EMITIDA P/ O PEDIDO
'   Tipo de NFe: 0-Entrada  1-Saída
	s = "SELECT" & _
			" e.id_nfe_emitente," & _
			" e.NFe_serie_NF," & _
			" e.NFe_numero_NF," & _
			" DateDiff(d, e.dt_emissao, getdate()) AS qtde_dias_emissao" & _
		" FROM t_NFe_EMISSAO e" & _
		" WHERE" & _
			" (e.pedido = '" & pedido & "')" & _
			" AND (e.tipo_NF = '1')" & _
			" AND (e.st_anulado = 0)" & _
			" AND (e.codigo_retorno_NFe_T1 = 1)"
	set rP = get_registro_t_parametro(ID_PARAMETRO_NF_FlagOperacaoTriangular)
	if Trim("" & rP.id) <> "" then
		if (rP.campo_inteiro = 1) then
			s = s & " AND NOT EXISTS" & _
					" (SELECT 1 FROM t_NFe_TRIANGULAR t WHERE t.pedido = '" & pedido & "' AND t.emissao_status = 2 AND e.id_nfe_emitente = t.id_nfe_emitente " & _ 
						" AND e.NFe_serie_NF = t.Nfe_serie_remessa AND e.NFe_numero_NF = t.Nfe_numero_remessa)"
			end if
		end if
	'Para assegurar que não haverá nenhum risco de obter um número de NFe errado (de outro pedido) devido a cancelamento de NFe e
	'reaproveitamento desse número em outro pedido, verifica se o número da NFe consta no pedido
	s = s & _
			" AND (" & _
				"(NFe_numero_NF IN (SELECT num_obs_2 FROM t_PEDIDO WHERE pedido = '" & pedido & "'))" & _
				" OR " & _
				"(NFe_numero_NF IN (SELECT num_obs_3 FROM t_PEDIDO WHERE pedido = '" & pedido & "'))" & _
				" OR " & _
				"(NFe_numero_NF IN (SELECT num_obs_4 FROM t_PEDIDO WHERE pedido = '" & pedido & "'))" & _
			")"

	s = s & " ORDER BY" & _
			" id DESC"
	set r = cn.Execute(s)
	if r.Eof then 
		r.Close
		set r = nothing
		exit function
		end if
	
	set fso = Server.CreateObject("Scripting.FileSystemObject")
	
	blnArqDanfeExiste = false
	do while (Not r.Eof) And (Not blnArqDanfeExiste)
		intIdNFeEmitente = CInt(r("id_nfe_emitente"))
		lngSerieNFe = CLng(r("NFe_serie_NF"))
		strSerieNFe = NFeFormataSerieNF(lngSerieNFe)
		lngNumeroNFe = CLng(r("NFe_numero_NF"))
		strNumeroNFe = NFeFormataNumeroNF(lngNumeroNFe)
		lngQtdeDiasEmissao = CLng(r("qtde_dias_emissao"))
	
	'	EXCEDE PERÍODO MÁXIMO EM QUE O DANFE FICA ACESSÍVEL NO PEDIDO?
		blnQtdeDiasOK = true
		if lngMaxDiasEmissao <> -1 then
			if lngQtdeDiasEmissao > lngMaxDiasEmissao then blnQtdeDiasOK = false 'exit function
			end if
			
		if blnQtdeDiasOK then
			set rNfeEmitente = le_nfe_emitente(intIdNFeEmitente)

			for iCfg=LBound(rNfeEmitente.vCfgDanfe) to UBound(rNfeEmitente.vCfgDanfe)
				if Trim("" & rNfeEmitente.vCfgDanfe(iCfg).convencao_nome_arq_pdf_danfe) <> "" then
				'	MONTA O NOME DO ARQUIVO DA DANFE
					strArqDanfe = ""
					strArqDanfeAux = Trim("" & rNfeEmitente.vCfgDanfe(iCfg).convencao_nome_arq_pdf_danfe)
					if strArqDanfeAux <> "" then
						strArqDanfe = Replace(Replace(strArqDanfeAux, "[NUMERO_NFE]", strNumeroNFe), "[SERIE_NFE]", strSerieNFe)
						strPathArqDanfe = Trim("" & rNfeEmitente.vCfgDanfe(iCfg).diretorio_pdf_danfe)
						strArqDanfeCompleto = strPathArqDanfe & "\" & strArqDanfe
						blnArqDanfeExiste = fso.FileExists(strArqDanfeCompleto)
						end if
					end if
				
				if blnArqDanfeExiste then exit for
				next
			end if 'if blnQtdeDiasOK
		
		r.MoveNext
		loop

	set fso = nothing

	if r.State <> 0 then r.Close
	set r = nothing
	if Not blnArqDanfeExiste then exit function
	
'	VERIFICA SE A NFe FOI CANCELADA
	s = "SELECT" & _
			" NFe_T1_servidor_BD," & _
			" NFe_T1_nome_BD," & _
			" NFe_T1_usuario_BD," & _
			" NFe_T1_senha_BD" & _
		" FROM t_NFe_EMITENTE" & _
		" WHERE" & _
			" (id = " & Cstr(intIdNFeEmitente) & ")"
	set tNE = cn.Execute(s)
	if tNE.Eof then exit function
	
	strNfeT1ServidorBd = Trim("" & tNE("NFe_T1_servidor_BD"))
	strNfeT1NomeBd = Trim("" & tNE("NFe_T1_nome_BD"))
	strNfeT1UsuarioBd = Trim("" & tNE("NFe_T1_usuario_BD"))
	strNfeT1SenhaCriptografadaBd = Trim("" & tNE("NFe_T1_senha_BD"))
	
	tNE.Close
	set tNE = nothing
	
	chave = gera_chave(FATOR_BD)
	decodifica_dado strNfeT1SenhaCriptografadaBd, senha_decodificada, chave
	s = "Provider=SQLOLEDB;" & _
		"Data Source=" & strNfeT1ServidorBd & ";" & _
		"Initial Catalog=" & strNfeT1NomeBd & ";" & _
		"User ID=" & strNfeT1UsuarioBd & ";" & _
		"Password=" & senha_decodificada & ";"
	set dbcNFe = server.CreateObject("ADODB.Connection")
	dbcNFe.ConnectionTimeout = 45
	dbcNFe.CommandTimeout = 900
	dbcNFe.ConnectionString = s
	dbcNFe.Open

	s = "SELECT" & _
			" Nfe," & _
			" Serie," & _
			" Convert(tinyint, Coalesce(CANCELADA,0)) AS CANCELADA" & _
		" FROM NFE" & _
		" WHERE" & _
			" (Serie = '" & NFeFormataSerieNF(lngSerieNFe) & "')" & _
			" AND (Nfe = '" & NFeFormataNumeroNF(lngNumeroNFe) & "')"
	set tNFE = dbcNFe.Execute(s)
	strFlagCancelada = ""
	if tNFE.Eof then
		dbcNFe.Close
		set dbcNFe = nothing
		exit function
		end if
		
	strFlagCancelada = Trim("" & tNFE("CANCELADA"))
	
	tNFE.Close
	set tNFE = nothing
	
	dbcNFe.Close
	set dbcNFe = nothing
	
'	NFE FOI CANCELADA?
	if strFlagCancelada = "1" then exit function

	strLinkDanfe = "<a name='lnkDanfePedido' href='../Global/DownloadDanfe.asp?file=" & strArqDanfe & "&emitente=" & Cstr(intIdNFeEmitente) & "&force=true" & "' title='Clique para consultar o PDF da DANFE deste pedido'>" & html_elemento_interno_anchor & "</a>"
	monta_link_para_DANFE = strLinkDanfe
end function



' ___________________________________________________________________________________
' MONTA LINK PARA DANFE NFE
' Monta o link para a DANFE da NFe especificada, se existir
function monta_link_para_DANFE_NFe(byval pedido, byval numero_NFe, byval max_dias_emissao, byval html_elemento_interno_anchor)
dim s, strSerieNFe, strNumeroNFe
dim strArqDanfe, strArqDanfeAux, strArqDanfeCompleto, strPathArqDanfe, strLinkDanfe
dim intIdNFeEmitente, lngSerieNFe, lngNumeroNFe, lngQtdeDiasEmissao, lngMaxDiasEmissao
dim blnArqDanfeExiste
dim fso
dim r
dim tNE, tNFE
dim dbcNFe
dim strNfeT1ServidorBd, strNfeT1NomeBd, strNfeT1UsuarioBd, strNfeT1SenhaCriptografadaBd
dim chave
dim senha_decodificada
dim strFlagCancelada
dim blnQtdeDiasOK
dim rNfeEmitente, iCfg
	
	monta_link_para_DANFE_NFe = ""
	
	pedido = Trim("" & pedido)
	if pedido = "" then exit function
	
	numero_NFe = Trim("" & numero_NFe)
	if numero_NFe = "" then exit function

	lngMaxDiasEmissao = 0
	if IsNumeric(max_dias_emissao) then lngMaxDiasEmissao = CLng(max_dias_emissao)
	
	s = "SELECT id_nfe_emitente FROM t_PEDIDO WHERE (pedido = '" & pedido & "')"
	set r = cn.Execute(s)
	if r.Eof then exit function

	intIdNFeEmitente = r("id_nfe_emitente")

'	LOCALIZA O REGISTRO DA ÚLTIMA EMISSÃO DESSA NFE
'	NÃO RESTRINGE POR Nº DE PEDIDO P/ NÃO EXCLUIR AS EMISSÕES MANUAIS EM QUE NÃO SE INFORMOU O Nº PEDIDO
'   Tipo de NFe: 0-Entrada  1-Saída
	s = "SELECT" & _
			" e.id_nfe_emitente," & _
			" e.NFe_serie_NF," & _
			" e.NFe_numero_NF," & _
			" DateDiff(d, e.dt_emissao, getdate()) AS qtde_dias_emissao" & _
		" FROM t_NFe_EMISSAO e" & _
		" WHERE" & _
			" (id_nfe_emitente = " & intIdNFeEmitente & ")" & _
			" AND (NFe_numero_NF = " & numero_NFe & ")" & _
			" AND (e.tipo_NF = '1')" & _
			" AND (e.st_anulado = 0)" & _
			" AND (e.codigo_retorno_NFe_T1 = 1)"
	'Para assegurar que não haverá nenhum risco de obter um número de NFe errado (de outro pedido) devido a cancelamento de NFe e
	'reaproveitamento desse número em outro pedido, verifica se o número da NFe consta no pedido
	s = s & _
			" AND (" & _
				"(NFe_numero_NF IN (SELECT num_obs_2 FROM t_PEDIDO WHERE pedido = '" & pedido & "'))" & _
				" OR " & _
				"(NFe_numero_NF IN (SELECT num_obs_3 FROM t_PEDIDO WHERE pedido = '" & pedido & "'))" & _
				" OR " & _
				"(NFe_numero_NF IN (SELECT num_obs_4 FROM t_PEDIDO WHERE pedido = '" & pedido & "'))" & _
			")"

	s = s & " ORDER BY" & _
			" id DESC"
	set r = cn.Execute(s)
	if r.Eof then
		r.Close
		set r = nothing
		exit function
		end if
	
	set fso = Server.CreateObject("Scripting.FileSystemObject")
	
	blnArqDanfeExiste = false
	do while (Not r.Eof) And (Not blnArqDanfeExiste)
		intIdNFeEmitente = CInt(r("id_nfe_emitente"))
		lngSerieNFe = CLng(r("NFe_serie_NF"))
		strSerieNFe = NFeFormataSerieNF(lngSerieNFe)
		lngNumeroNFe = CLng(r("NFe_numero_NF"))
		strNumeroNFe = NFeFormataNumeroNF(lngNumeroNFe)
		lngQtdeDiasEmissao = CLng(r("qtde_dias_emissao"))
	
	'	EXCEDE PERÍODO MÁXIMO EM QUE O DANFE FICA ACESSÍVEL NO PEDIDO?
		blnQtdeDiasOK = true
		if lngMaxDiasEmissao <> -1 then
			if lngQtdeDiasEmissao > lngMaxDiasEmissao then blnQtdeDiasOK = false 'exit function
			end if
			
		if blnQtdeDiasOK then
			set rNfeEmitente = le_nfe_emitente(intIdNFeEmitente)

			for iCfg=LBound(rNfeEmitente.vCfgDanfe) to UBound(rNfeEmitente.vCfgDanfe)
				if Trim("" & rNfeEmitente.vCfgDanfe(iCfg).convencao_nome_arq_pdf_danfe) <> "" then
				'	MONTA O NOME DO ARQUIVO DA DANFE
					strArqDanfe = ""
					strArqDanfeAux = Trim("" & rNfeEmitente.vCfgDanfe(iCfg).convencao_nome_arq_pdf_danfe)
					if strArqDanfeAux <> "" then
						strArqDanfe = Replace(Replace(strArqDanfeAux, "[NUMERO_NFE]", strNumeroNFe), "[SERIE_NFE]", strSerieNFe)
						strPathArqDanfe = Trim("" & rNfeEmitente.vCfgDanfe(iCfg).diretorio_pdf_danfe)
						strArqDanfeCompleto = strPathArqDanfe & "\" & strArqDanfe
						blnArqDanfeExiste = fso.FileExists(strArqDanfeCompleto)
						end if
					end if
				
				if blnArqDanfeExiste then exit for
				next
			end if 'if blnQtdeDiasOK

		r.MoveNext
		loop

	set fso = nothing

	if r.State <> 0 then r.Close
	set r = nothing
	if Not blnArqDanfeExiste then exit function
	
'	VERIFICA SE A NFe FOI CANCELADA
	s = "SELECT" & _
			" NFe_T1_servidor_BD," & _
			" NFe_T1_nome_BD," & _
			" NFe_T1_usuario_BD," & _
			" NFe_T1_senha_BD" & _
		" FROM t_NFe_EMITENTE" & _
		" WHERE" & _
			" (id = " & Cstr(intIdNFeEmitente) & ")"
	set tNE = cn.Execute(s)
	if tNE.Eof then exit function
	
	strNfeT1ServidorBd = Trim("" & tNE("NFe_T1_servidor_BD"))
	strNfeT1NomeBd = Trim("" & tNE("NFe_T1_nome_BD"))
	strNfeT1UsuarioBd = Trim("" & tNE("NFe_T1_usuario_BD"))
	strNfeT1SenhaCriptografadaBd = Trim("" & tNE("NFe_T1_senha_BD"))
	
	tNE.Close
	set tNE = nothing
	
	chave = gera_chave(FATOR_BD)
	decodifica_dado strNfeT1SenhaCriptografadaBd, senha_decodificada, chave
	s = "Provider=SQLOLEDB;" & _
		"Data Source=" & strNfeT1ServidorBd & ";" & _
		"Initial Catalog=" & strNfeT1NomeBd & ";" & _
		"User ID=" & strNfeT1UsuarioBd & ";" & _
		"Password=" & senha_decodificada & ";"
	set dbcNFe = server.CreateObject("ADODB.Connection")
	dbcNFe.ConnectionTimeout = 45
	dbcNFe.CommandTimeout = 900
	dbcNFe.ConnectionString = s
	dbcNFe.Open

	s = "SELECT" & _
			" Nfe," & _
			" Serie," & _
			" Convert(tinyint, Coalesce(CANCELADA,0)) AS CANCELADA" & _
		" FROM NFE" & _
		" WHERE" & _
			" (Serie = '" & NFeFormataSerieNF(lngSerieNFe) & "')" & _
			" AND (Nfe = '" & NFeFormataNumeroNF(lngNumeroNFe) & "')"
	set tNFE = dbcNFe.Execute(s)
	strFlagCancelada = ""
	if tNFE.Eof then
		dbcNFe.Close
		set dbcNFe = nothing
		exit function
		end if
		
	strFlagCancelada = Trim("" & tNFE("CANCELADA"))
	
	tNFE.Close
	set tNFE = nothing
	
	dbcNFe.Close
	set dbcNFe = nothing
	
'	NFE FOI CANCELADA?
	if strFlagCancelada = "1" then exit function
	
	strLinkDanfe = "<a name='lnkDanfePedido' href='../Global/DownloadDanfe.asp?file=" & strArqDanfe & "&emitente=" & Cstr(intIdNFeEmitente) & "&force=true" & "' title='Clique para consultar o PDF da NFe'>" & html_elemento_interno_anchor & "</a>"
	monta_link_para_DANFE_NFe = strLinkDanfe
end function



' ___________________________________________________________________________________
' MONTA LINK PARA ULTIMA DANFE
' - traz a última nota fiscal de um pedido (seja de venda ou de remessa), se houver

function monta_link_para_ultima_DANFE(byval pedido, byval max_dias_emissao, byval html_elemento_interno_anchor)
dim s, strSerieNFe, strNumeroNFe
dim strArqDanfe, strArqDanfeAux, strArqDanfeCompleto, strPathArqDanfe, strLinkDanfe
dim intIdNFeEmitente, lngSerieNFe, lngNumeroNFe, lngQtdeDiasEmissao, lngMaxDiasEmissao
dim blnArqDanfeExiste
dim fso
dim r
dim tNE, tNFE
dim dbcNFe
dim strNfeT1ServidorBd, strNfeT1NomeBd, strNfeT1UsuarioBd, strNfeT1SenhaCriptografadaBd
dim chave
dim senha_decodificada
dim strFlagCancelada
dim blnQtdeDiasOK
dim lngNumVenda
dim lngNumRemessa
dim intIdEmitenteRemessa
dim rP, rNfeEmitente, iCfg
dim blnLocalizouRemessa
	
	monta_link_para_ultima_DANFE = ""
	
	pedido = Trim("" & pedido)
	if pedido = "" then exit function
	
	lngMaxDiasEmissao = 0
	if IsNumeric(max_dias_emissao) then lngMaxDiasEmissao = CLng(max_dias_emissao)

	'VERIFICA SE EXISTE UMA NOTA DE REMESSA VÁLIDA EMITIDA PARA O PEDIDO
	lngNumRemessa = 0
	intIdEmitenteRemessa = 0
	set rP = get_registro_t_parametro(ID_PARAMETRO_NF_FlagOperacaoTriangular)
	if Trim("" & rP.id) <> "" then
		if (rP.campo_inteiro = 1) then
			s = "SELECT * FROM t_NFe_TRIANGULAR t WHERE t.pedido = '" & pedido & "'" & _
				" AND (t.Nfe_venda_emissao_status = 2)" & _
				" AND (t.emissao_status IN (1, 2))"
			'Para assegurar que não haverá nenhum risco de obter um número de NFe errado (de outro pedido) devido a cancelamento de NFe e
			'reaproveitamento desse número em outro pedido, verifica se o número da NFe consta no pedido
			s = s & _
				" AND (" & _
					"(Nfe_numero_remessa IN (SELECT num_obs_2 FROM t_PEDIDO WHERE pedido = '" & pedido & "'))" & _
					" OR " & _
					"(Nfe_numero_remessa IN (SELECT num_obs_3 FROM t_PEDIDO WHERE pedido = '" & pedido & "'))" & _
					" OR " & _
					"(Nfe_numero_remessa IN (SELECT num_obs_4 FROM t_PEDIDO WHERE pedido = '" & pedido & "'))" & _
				")" & _
				" AND (" & _
					"(Nfe_numero_venda IN (SELECT num_obs_2 FROM t_PEDIDO WHERE pedido = '" & pedido & "'))" & _
					" OR " & _
					"(Nfe_numero_venda IN (SELECT num_obs_3 FROM t_PEDIDO WHERE pedido = '" & pedido & "'))" & _
					" OR " & _
					"(Nfe_numero_venda IN (SELECT num_obs_4 FROM t_PEDIDO WHERE pedido = '" & pedido & "'))" & _
				")"
			set r = cn.Execute(s)
			if Not r.Eof then 
				lngNumVenda = CLng(r("Nfe_numero_venda"))
				lngNumRemessa = CLng(r("Nfe_numero_remessa"))
				intIdEmitenteRemessa = CInt(r("id_nfe_emitente"))
				end if
			r.Close
			set r = nothing
			end if
		end if

	
'	TENTA LOCALIZAR A NFe EMITIDA ESPECÍFICA DE REMESSA P/ O PEDIDO
'   Tipo de NFe: 0-Entrada  1-Saída
	blnLocalizouRemessa = False
	if lngNumRemessa > 0 then
		s = "SELECT" & _
				" id_nfe_emitente," & _
				" NFe_serie_NF," & _
				" NFe_numero_NF," & _
				" DateDiff(d, dt_emissao, getdate()) AS qtde_dias_emissao" & _
			" FROM t_NFe_EMISSAO" & _
			" WHERE" & _
				" (pedido = '" & pedido & "')" & _
				" AND (tipo_NF = '1')" & _
				" AND (st_anulado = 0)" & _
				" AND (codigo_retorno_NFe_T1 = 1)" & _
				" AND (NFe_numero_NF = " & CStr(lngNumRemessa) & ")" & _
				" AND (id_nfe_emitente = " & CStr(intIdEmitenteRemessa) & ")"
		'Para assegurar que não haverá nenhum risco de obter um número de NFe errado (de outro pedido) devido a cancelamento de NFe e
		'reaproveitamento desse número em outro pedido, verifica se o número da NFe consta no pedido
		s = s & _
				" AND (" & _
					"(NFe_numero_NF IN (SELECT num_obs_2 FROM t_PEDIDO WHERE pedido = '" & pedido & "'))" & _
					" OR " & _
					"(NFe_numero_NF IN (SELECT num_obs_3 FROM t_PEDIDO WHERE pedido = '" & pedido & "'))" & _
					" OR " & _
					"(NFe_numero_NF IN (SELECT num_obs_4 FROM t_PEDIDO WHERE pedido = '" & pedido & "'))" & _
				")"
		s = s & _
			" ORDER BY" & _
				" id DESC"
		set r = cn.Execute(s)
		if Not r.Eof then 
			blnLocalizouRemessa = True
			end if
		end if

'	SE NÃO LOCALIZOU A NOTA ESPECÍFICA, LOCALIZA A ÚLTIMA NFe EMITIDA P/ O PEDIDO
'   Tipo de NFe: 0-Entrada  1-Saída
	if Not blnLocalizouRemessa then
		s = "SELECT" & _
				" id_nfe_emitente," & _
				" NFe_serie_NF," & _
				" NFe_numero_NF," & _
				" DateDiff(d, dt_emissao, getdate()) AS qtde_dias_emissao" & _
			" FROM t_NFe_EMISSAO" & _
			" WHERE" & _
				" (pedido = '" & pedido & "')" & _
				" AND (tipo_NF = '1')" & _
				" AND (st_anulado = 0)" & _
				" AND (codigo_retorno_NFe_T1 = 1)"
		'Para assegurar que não haverá nenhum risco de obter um número de NFe errado (de outro pedido) devido a cancelamento de NFe e
		'reaproveitamento desse número em outro pedido, verifica se o número da NFe consta no pedido
		s = s & _
				" AND (" & _
					"(NFe_numero_NF IN (SELECT num_obs_2 FROM t_PEDIDO WHERE pedido = '" & pedido & "'))" & _
					" OR " & _
					"(NFe_numero_NF IN (SELECT num_obs_3 FROM t_PEDIDO WHERE pedido = '" & pedido & "'))" & _
					" OR " & _
					"(NFe_numero_NF IN (SELECT num_obs_4 FROM t_PEDIDO WHERE pedido = '" & pedido & "'))" & _
				")"

		if lngNumVenda > 0 then
			s = s & " AND (NFe_numero_NF <> " & CStr(lngNumVenda) & ")"
			end if

		s = s & " ORDER BY" & _
				" id DESC"
		set r = cn.Execute(s)
		if r.Eof then
			r.Close
			set r = nothing
			exit function
			end if
		end if
	
	set fso = Server.CreateObject("Scripting.FileSystemObject")
	
	blnArqDanfeExiste = false
	do while (Not r.Eof) And (Not blnArqDanfeExiste)
		intIdNFeEmitente = CInt(r("id_nfe_emitente"))
		lngSerieNFe = CLng(r("NFe_serie_NF"))
		strSerieNFe = NFeFormataSerieNF(lngSerieNFe)
		lngNumeroNFe = CLng(r("NFe_numero_NF"))
		strNumeroNFe = NFeFormataNumeroNF(lngNumeroNFe)
		lngQtdeDiasEmissao = CLng(r("qtde_dias_emissao"))
	
	'	EXCEDE PERÍODO MÁXIMO EM QUE O DANFE FICA ACESSÍVEL NO PEDIDO?
		blnQtdeDiasOK = true
		if lngMaxDiasEmissao <> -1 then
			if lngQtdeDiasEmissao > lngMaxDiasEmissao then blnQtdeDiasOK = false 'exit function
			end if
			
		if blnQtdeDiasOK then
			set rNfeEmitente = le_nfe_emitente(intIdNFeEmitente)
			for iCfg=LBound(rNfeEmitente.vCfgDanfe) to UBound(rNfeEmitente.vCfgDanfe)
				if Trim("" & rNfeEmitente.vCfgDanfe(iCfg).convencao_nome_arq_pdf_danfe) <> "" then
				'	MONTA O NOME DO ARQUIVO DA DANFE
					strArqDanfe = ""
					strArqDanfeAux = Trim("" & rNfeEmitente.vCfgDanfe(iCfg).convencao_nome_arq_pdf_danfe)
					if strArqDanfeAux <> "" then
						strArqDanfe = Replace(Replace(strArqDanfeAux, "[NUMERO_NFE]", strNumeroNFe), "[SERIE_NFE]", strSerieNFe)
						strPathArqDanfe = Trim("" & rNfeEmitente.vCfgDanfe(iCfg).diretorio_pdf_danfe)
						strArqDanfeCompleto = strPathArqDanfe & "\" & strArqDanfe
						blnArqDanfeExiste = fso.FileExists(strArqDanfeCompleto)
						end if
					end if
				
				if blnArqDanfeExiste then exit for
				next
			end if 'if blnQtdeDiasOK

		r.MoveNext
		loop

	set fso = nothing

	if r.State <> 0 then r.Close
	set r = nothing
	if Not blnArqDanfeExiste then exit function
	
'	VERIFICA SE A NFe FOI CANCELADA
	s = "SELECT" & _
			" NFe_T1_servidor_BD," & _
			" NFe_T1_nome_BD," & _
			" NFe_T1_usuario_BD," & _
			" NFe_T1_senha_BD" & _
		" FROM t_NFe_EMITENTE" & _
		" WHERE" & _
			" (id = " & Cstr(intIdNFeEmitente) & ")"
	set tNE = cn.Execute(s)
	if tNE.Eof then exit function
	
	strNfeT1ServidorBd = Trim("" & tNE("NFe_T1_servidor_BD"))
	strNfeT1NomeBd = Trim("" & tNE("NFe_T1_nome_BD"))
	strNfeT1UsuarioBd = Trim("" & tNE("NFe_T1_usuario_BD"))
	strNfeT1SenhaCriptografadaBd = Trim("" & tNE("NFe_T1_senha_BD"))
	
	tNE.Close
	set tNE = nothing
	
	chave = gera_chave(FATOR_BD)
	decodifica_dado strNfeT1SenhaCriptografadaBd, senha_decodificada, chave
	s = "Provider=SQLOLEDB;" & _
		"Data Source=" & strNfeT1ServidorBd & ";" & _
		"Initial Catalog=" & strNfeT1NomeBd & ";" & _
		"User ID=" & strNfeT1UsuarioBd & ";" & _
		"Password=" & senha_decodificada & ";"
	set dbcNFe = server.CreateObject("ADODB.Connection")
	dbcNFe.ConnectionTimeout = 45
	dbcNFe.CommandTimeout = 900
	dbcNFe.ConnectionString = s
	dbcNFe.Open

	s = "SELECT" & _
			" Nfe," & _
			" Serie," & _
			" Convert(tinyint, Coalesce(CANCELADA,0)) AS CANCELADA" & _
		" FROM NFE" & _
		" WHERE" & _
			" (Serie = '" & NFeFormataSerieNF(lngSerieNFe) & "')" & _
			" AND (Nfe = '" & NFeFormataNumeroNF(lngNumeroNFe) & "')"
	set tNFE = dbcNFe.Execute(s)
	strFlagCancelada = ""
	if tNFE.Eof then
		dbcNFe.Close
		set dbcNFe = nothing
		exit function
		end if
		
	strFlagCancelada = Trim("" & tNFE("CANCELADA"))
	
	tNFE.Close
	set tNFE = nothing
	
	dbcNFe.Close
	set dbcNFe = nothing
	
'	NFE FOI CANCELADA?
	if strFlagCancelada = "1" then exit function

	strLinkDanfe = "<a name='lnkDanfePedido' href='../Global/DownloadDanfe.asp?file=" & strArqDanfe & "&emitente=" & Cstr(intIdNFeEmitente) & "&force=true" & "' title='Clique para consultar o PDF da DANFE deste pedido'>" & html_elemento_interno_anchor & "</a>"
	monta_link_para_ultima_DANFE = strLinkDanfe
end function

' ___________________________________________________________________________________
' MONTA LINK PARA DANFE REMESSA
' - traz a nota fiscal de remessa de um pedido, se houver

function monta_link_para_DANFE_remessa(byval pedido, byval max_dias_emissao, byval html_elemento_interno_anchor)
dim s, strSerieNFe, strNumeroNFe
dim strArqDanfe, strArqDanfeAux, strArqDanfeCompleto, strPathArqDanfe, strLinkDanfe
dim intIdNFeEmitente, lngSerieNFe, lngNumeroNFe, lngQtdeDiasEmissao, lngMaxDiasEmissao
dim blnArqDanfeExiste
dim fso
dim r
dim tNE, tNFE
dim dbcNFe
dim strNfeT1ServidorBd, strNfeT1NomeBd, strNfeT1UsuarioBd, strNfeT1SenhaCriptografadaBd
dim chave
dim senha_decodificada
dim strFlagCancelada
dim blnQtdeDiasOK
dim lngNumVenda
dim lngNumRemessa
dim intIdEmitenteRemessa
dim rP, rNfeEmitente, iCfg
	
	monta_link_para_DANFE_remessa = ""
	
	pedido = Trim("" & pedido)
	if pedido = "" then exit function
	
	lngMaxDiasEmissao = 0
	if IsNumeric(max_dias_emissao) then lngMaxDiasEmissao = CLng(max_dias_emissao)

	'VERIFICA SE EXISTE UMA NOTA DE REMESSA VÁLIDA EMITIDA PARA O PEDIDO
	lngNumRemessa = 0
	intIdEmitenteRemessa = 0
	set rP = get_registro_t_parametro(ID_PARAMETRO_NF_FlagOperacaoTriangular)
	if Trim("" & rP.id) <> "" then
		if (rP.campo_inteiro = 1) then
			s = "SELECT * FROM t_NFe_TRIANGULAR t WHERE t.pedido = '" & pedido & "'" & _
				" AND (t.Nfe_venda_emissao_status = 2)" & _
				" AND (t.emissao_status IN (1, 2))"
			'Para assegurar que não haverá nenhum risco de obter um número de NFe errado (de outro pedido) devido a cancelamento de NFe e
			'reaproveitamento desse número em outro pedido, verifica se o número da NFe consta no pedido
			s = s & _
				" AND (" & _
					"(Nfe_numero_remessa IN (SELECT num_obs_2 FROM t_PEDIDO WHERE pedido = '" & pedido & "'))" & _
					" OR " & _
					"(Nfe_numero_remessa IN (SELECT num_obs_3 FROM t_PEDIDO WHERE pedido = '" & pedido & "'))" & _
					" OR " & _
					"(Nfe_numero_remessa IN (SELECT num_obs_4 FROM t_PEDIDO WHERE pedido = '" & pedido & "'))" & _
				")" & _
				" AND (" & _
					"(Nfe_numero_venda IN (SELECT num_obs_2 FROM t_PEDIDO WHERE pedido = '" & pedido & "'))" & _
					" OR " & _
					"(Nfe_numero_venda IN (SELECT num_obs_3 FROM t_PEDIDO WHERE pedido = '" & pedido & "'))" & _
					" OR " & _
					"(Nfe_numero_venda IN (SELECT num_obs_4 FROM t_PEDIDO WHERE pedido = '" & pedido & "'))" & _
				")"
			set r = cn.Execute(s)
			if Not r.Eof then 
				lngNumVenda = CLng(r("Nfe_numero_venda"))
				lngNumRemessa = CLng(r("Nfe_numero_remessa"))
				intIdEmitenteRemessa = CInt(r("id_nfe_emitente"))
			else
				r.Close
				set r = nothing
				exit function
				end if
			r.Close
			set r = nothing
			end if
		end if

	if lngNumRemessa <= 0 then exit function

'	TENTA LOCALIZAR A NFe EMITIDA ESPECÍFICA DE REMESSA P/ O PEDIDO
'   Tipo de NFe: 0-Entrada  1-Saída
	if lngNumRemessa > 0 then
		s = "SELECT" & _
				" id_nfe_emitente," & _
				" NFe_serie_NF," & _
				" NFe_numero_NF," & _
				" DateDiff(d, dt_emissao, getdate()) AS qtde_dias_emissao" & _
			" FROM t_NFe_EMISSAO" & _
			" WHERE" & _
				" (pedido = '" & pedido & "')" & _
				" AND (tipo_NF = '1')" & _
				" AND (st_anulado = 0)" & _
				" AND (codigo_retorno_NFe_T1 = 1)" & _
				" AND (NFe_numero_NF = " & CStr(lngNumRemessa) & ")" & _
				" AND (id_nfe_emitente = " & CStr(intIdEmitenteRemessa) & ")"
		'Para assegurar que não haverá nenhum risco de obter um número de NFe errado (de outro pedido) devido a cancelamento de NFe e
		'reaproveitamento desse número em outro pedido, verifica se o número da NFe consta no pedido
		s = s & _
				" AND (" & _
					"(NFe_numero_NF IN (SELECT num_obs_2 FROM t_PEDIDO WHERE pedido = '" & pedido & "'))" & _
					" OR " & _
					"(NFe_numero_NF IN (SELECT num_obs_3 FROM t_PEDIDO WHERE pedido = '" & pedido & "'))" & _
					" OR " & _
					"(NFe_numero_NF IN (SELECT num_obs_4 FROM t_PEDIDO WHERE pedido = '" & pedido & "'))" & _
				")"
		s = s & _
			" ORDER BY" & _
				" id DESC"
		set r = cn.Execute(s)
		if r.Eof then
			r.Close
			set r = nothing
			exit function
			end if
		end if
	
	set fso = Server.CreateObject("Scripting.FileSystemObject")
	
	blnArqDanfeExiste = false
	do while (Not r.Eof) And (Not blnArqDanfeExiste)
		intIdNFeEmitente = CInt(r("id_nfe_emitente"))
		lngSerieNFe = CLng(r("NFe_serie_NF"))
		strSerieNFe = NFeFormataSerieNF(lngSerieNFe)
		lngNumeroNFe = CLng(r("NFe_numero_NF"))
		strNumeroNFe = NFeFormataNumeroNF(lngNumeroNFe)
		lngQtdeDiasEmissao = CLng(r("qtde_dias_emissao"))
	
	'	EXCEDE PERÍODO MÁXIMO EM QUE O DANFE FICA ACESSÍVEL NO PEDIDO?
		blnQtdeDiasOK = true
		if lngMaxDiasEmissao <> -1 then
			if lngQtdeDiasEmissao > lngMaxDiasEmissao then blnQtdeDiasOK = false 'exit function
			end if
			
		if blnQtdeDiasOK then
			set rNfeEmitente = le_nfe_emitente(intIdNFeEmitente)

			for iCfg=LBound(rNfeEmitente.vCfgDanfe) to UBound(rNfeEmitente.vCfgDanfe)
				if Trim("" & rNfeEmitente.vCfgDanfe(iCfg).convencao_nome_arq_pdf_danfe) <> "" then
				'	MONTA O NOME DO ARQUIVO DA DANFE
					strArqDanfe = ""
					strArqDanfeAux = Trim("" & rNfeEmitente.vCfgDanfe(iCfg).convencao_nome_arq_pdf_danfe)
					if strArqDanfeAux <> "" then
						strArqDanfe = Replace(Replace(strArqDanfeAux, "[NUMERO_NFE]", strNumeroNFe), "[SERIE_NFE]", strSerieNFe)
						strPathArqDanfe = Trim("" & rNfeEmitente.vCfgDanfe(iCfg).diretorio_pdf_danfe)
						strArqDanfeCompleto = strPathArqDanfe & "\" & strArqDanfe
						blnArqDanfeExiste = fso.FileExists(strArqDanfeCompleto)
						end if
					end if
				
				if blnArqDanfeExiste then exit for
				next
			end if 'if blnQtdeDiasOK
		
		r.MoveNext
		loop

	set fso = nothing

	if r.State <> 0 then r.Close
	set r = nothing
	if Not blnArqDanfeExiste then exit function
	
'	VERIFICA SE A NFe FOI CANCELADA
	s = "SELECT" & _
			" NFe_T1_servidor_BD," & _
			" NFe_T1_nome_BD," & _
			" NFe_T1_usuario_BD," & _
			" NFe_T1_senha_BD" & _
		" FROM t_NFe_EMITENTE" & _
		" WHERE" & _
			" (id = " & Cstr(intIdNFeEmitente) & ")"
	set tNE = cn.Execute(s)
	if tNE.Eof then exit function
	
	strNfeT1ServidorBd = Trim("" & tNE("NFe_T1_servidor_BD"))
	strNfeT1NomeBd = Trim("" & tNE("NFe_T1_nome_BD"))
	strNfeT1UsuarioBd = Trim("" & tNE("NFe_T1_usuario_BD"))
	strNfeT1SenhaCriptografadaBd = Trim("" & tNE("NFe_T1_senha_BD"))
	
	tNE.Close
	set tNE = nothing
	
	chave = gera_chave(FATOR_BD)
	decodifica_dado strNfeT1SenhaCriptografadaBd, senha_decodificada, chave
	s = "Provider=SQLOLEDB;" & _
		"Data Source=" & strNfeT1ServidorBd & ";" & _
		"Initial Catalog=" & strNfeT1NomeBd & ";" & _
		"User ID=" & strNfeT1UsuarioBd & ";" & _
		"Password=" & senha_decodificada & ";"
	set dbcNFe = server.CreateObject("ADODB.Connection")
	dbcNFe.ConnectionTimeout = 45
	dbcNFe.CommandTimeout = 900
	dbcNFe.ConnectionString = s
	dbcNFe.Open

	s = "SELECT" & _
			" Nfe," & _
			" Serie," & _
			" Convert(tinyint, Coalesce(CANCELADA,0)) AS CANCELADA" & _
		" FROM NFE" & _
		" WHERE" & _
			" (Serie = '" & NFeFormataSerieNF(lngSerieNFe) & "')" & _
			" AND (Nfe = '" & NFeFormataNumeroNF(lngNumeroNFe) & "')"
	set tNFE = dbcNFe.Execute(s)
	strFlagCancelada = ""
	if tNFE.Eof then
		dbcNFe.Close
		set dbcNFe = nothing
		exit function
		end if
		
	strFlagCancelada = Trim("" & tNFE("CANCELADA"))
	
	tNFE.Close
	set tNFE = nothing
	
	dbcNFe.Close
	set dbcNFe = nothing
	
'	NFE FOI CANCELADA?
	if strFlagCancelada = "1" then exit function

	strLinkDanfe = "<a name='lnkDanfePedido' href='../Global/DownloadDanfe.asp?file=" & strArqDanfe & "&emitente=" & Cstr(intIdNFeEmitente) & "&force=true" & "' title='Clique para consultar o PDF da DANFE de remessa deste pedido'>" & html_elemento_interno_anchor & "</a>"
	monta_link_para_DANFE_remessa = strLinkDanfe
end function




' ___________________________________________________________________________________
' MONTA LINK PARA DANFE COM ICONE PDF
'
function monta_link_para_DANFE_com_icone_PDF(byval pedido, byval max_dias_emissao)
	monta_link_para_DANFE_com_icone_PDF = monta_link_para_DANFE(pedido, max_dias_emissao, "<img class='notPrint' src='../botao/pdf.png' border='0'>")
end function


' ___________________________________________________________________________________
' MONTA LINK PARA DANFE NFE COM ICONE PDF
'
function monta_link_para_DANFE_NFe_com_icone_PDF(byval pedido, byval numero_NFe, byval max_dias_emissao)
	monta_link_para_DANFE_NFe_com_icone_PDF = monta_link_para_DANFE_NFe(pedido, numero_NFe, max_dias_emissao, "<img class='notPrint' src='../botao/pdf.png' border='0'>")
end function


' ___________________________________________________________________________________
' MONTA LINK PARA DANFE NFE COM ICONE PDF PEQ
'
function monta_link_para_DANFE_NFe_com_icone_PDF_peq(byval pedido, byval numero_NFe, byval max_dias_emissao)
	monta_link_para_DANFE_NFe_com_icone_PDF_peq = monta_link_para_DANFE_NFe(pedido, numero_NFe, max_dias_emissao, "<img class='notPrint' src='../botao/pdf_12.png' border='0'>")
end function


' ___________________________________________________________________________________
' MONTA LINK PARA ULTIMA DANFE COM ICONE PDF
'
function monta_link_para_ultima_DANFE_com_icone_PDF(byval pedido, byval max_dias_emissao)
	monta_link_para_ultima_DANFE_com_icone_PDF = monta_link_para_ultima_DANFE(pedido, max_dias_emissao, "<img class='notPrint' src='../botao/pdf_truck_22.png' border='0'>")
end function

' ___________________________________________________________________________________
' MONTA LINK PARA DANFE REMESSA COM ICONE PDF
'
function monta_link_para_DANFE_remessa_com_icone_PDF(byval pedido, byval max_dias_emissao)
	monta_link_para_DANFE_remessa_com_icone_PDF = monta_link_para_DANFE_remessa(pedido, max_dias_emissao, "<img class='notPrint' src='../botao/pdf_truck_22.png' border='0'>")
end function



' ___________________________________________________________________________________
' IsPedidoNFeEmitida
'
function IsPedidoNFeEmitida(byval pedido, byref strSerieNFe, byref strNumeroNFe, byref msg_erro)
dim s
dim r
dim intIdNFeEmitente, lngSerieNFe, lngNumeroNFe

	IsPedidoNFeEmitida = False
	msg_erro = ""
	
	pedido = Trim("" & pedido)
	if pedido = "" then 
		msg_erro = "Não foi informado o número do pedido para analisar se há NFe emitida!"
		exit function
		end if
	
	s = "SELECT" & _
			" id_nfe_emitente," & _
			" NFe_serie_NF," & _
			" NFe_numero_NF" & _
		" FROM t_NFe_EMISSAO" & _
		" WHERE" & _
			" (pedido = '" & pedido & "')" & _
			" AND (tipo_NF = '1')" & _
			" AND (st_anulado = 0)" & _
			" AND (codigo_retorno_NFe_T1 = 1)" & _
		" ORDER BY" & _
			" id DESC"
	set r = cn.Execute(s)
	if r.Eof then 
		r.Close
		set r = nothing
		exit function
		end if
	
	intIdNFeEmitente = CInt(r("id_nfe_emitente"))
	lngSerieNFe = CLng(r("NFe_serie_NF"))
	strSerieNFe = NFeFormataSerieNF(lngSerieNFe)
	lngNumeroNFe = CLng(r("NFe_numero_NF"))
	strNumeroNFe = NFeFormataNumeroNF(lngNumeroNFe)
	
	IsPedidoNFeEmitida = True
	
	r.Close
	set r = nothing
end function



' ___________________________________________________________________________________
' IsPedidoCancelavelNFeEmitida
'
function IsPedidoCancelavelNFeEmitida(byval pedido, byref strSerieNFe, byref strNumeroNFe, byref msg_erro)
dim s
dim intIdNFeEmitente, lngSerieNFe, lngNumeroNFe
dim r
dim tNE, tNFE
dim dbcNFe
dim strNfeT1ServidorBd, strNfeT1NomeBd, strNfeT1UsuarioBd, strNfeT1SenhaCriptografadaBd
dim chave
dim senha_decodificada
dim strFlagCancelada

	IsPedidoCancelavelNFeEmitida = False

	strSerieNFe = ""
	strNumeroNFe = ""
	msg_erro = ""
	
	pedido = Trim("" & pedido)
	if pedido = "" then 
		msg_erro = "Não foi informado o número do pedido para analisar se há NFe emitida!"
		exit function
		end if
	
	s = "SELECT" & _
			" id_nfe_emitente," & _
			" NFe_serie_NF," & _
			" NFe_numero_NF" & _
		" FROM t_NFe_EMISSAO" & _
		" WHERE" & _
			" (pedido = '" & pedido & "')" & _
			" AND (st_anulado = 0)" & _
			" AND (codigo_retorno_NFe_T1 = 1)" & _
		" ORDER BY" & _
			" id DESC"
	set r = cn.Execute(s)
	if r.Eof then
	'	NÃO HÁ NENHUMA NFe QUE TENHA SIDO ENVIADA C/ SUCESSO P/ O SISTEMA DA TARGET ONE
		IsPedidoCancelavelNFeEmitida = True
		r.Close
		set r = nothing
		exit function
		end if

	intIdNFeEmitente = CInt(r("id_nfe_emitente"))
	lngSerieNFe = CLng(r("NFe_serie_NF"))
	strSerieNFe = NFeFormataSerieNF(lngSerieNFe)
	lngNumeroNFe = CLng(r("NFe_numero_NF"))
	strNumeroNFe = NFeFormataNumeroNF(lngNumeroNFe)

	r.Close
	set r = nothing

'	VERIFICA SE A NFe FOI CANCELADA
	s = "SELECT" & _
			" NFe_T1_servidor_BD," & _
			" NFe_T1_nome_BD," & _
			" NFe_T1_usuario_BD," & _
			" NFe_T1_senha_BD" & _
		" FROM t_NFe_EMITENTE" & _
		" WHERE" & _
			" (id = " & Cstr(intIdNFeEmitente) & ")"
	set tNE = cn.Execute(s)
	if tNE.Eof then 
		msg_erro = "Não foi localizado o registro do emitente (id = " & Cstr(intIdNFeEmitente) & ") ao analisar se há NFe emitida!!"
		exit function
		end if

	strNfeT1ServidorBd = Trim("" & tNE("NFe_T1_servidor_BD"))
	strNfeT1NomeBd = Trim("" & tNE("NFe_T1_nome_BD"))
	strNfeT1UsuarioBd = Trim("" & tNE("NFe_T1_usuario_BD"))
	strNfeT1SenhaCriptografadaBd = Trim("" & tNE("NFe_T1_senha_BD"))
	
	tNE.Close
	set tNE = nothing

	chave = gera_chave(FATOR_BD)
	decodifica_dado strNfeT1SenhaCriptografadaBd, senha_decodificada, chave
	s = "Provider=SQLOLEDB;" & _
		"Data Source=" & strNfeT1ServidorBd & ";" & _
		"Initial Catalog=" & strNfeT1NomeBd & ";" & _
		"User ID=" & strNfeT1UsuarioBd & ";" & _
		"Password=" & senha_decodificada & ";"
	set dbcNFe = server.CreateObject("ADODB.Connection")
	dbcNFe.ConnectionTimeout = 45
	dbcNFe.CommandTimeout = 900
	dbcNFe.ConnectionString = s
	dbcNFe.Open

	s = "SELECT" & _
			" Nfe," & _
			" Serie," & _
			" Convert(tinyint, Coalesce(CANCELADA,0)) AS CANCELADA" & _
		" FROM NFE" & _
		" WHERE" & _
			" (Serie = '" & NFeFormataSerieNF(lngSerieNFe) & "')" & _
			" AND (Nfe = '" & NFeFormataNumeroNF(lngNumeroNFe) & "')"
	set tNFE = dbcNFe.Execute(s)
	strFlagCancelada = ""
	if tNFE.Eof then
	'	NFe FOI ENVIADA AO SISTEMA DA TARGET ONE, MAS PROVAVELMENTE AINDA NÃO FOI PROCESSADA, POR ISSO NÃO CONSTA NESTA TABELA!!
	'	PORTANTO, COMO A NFe AINDA DEVE ESTAR EM PROCESSAMENTO, IMPEDE-SE O CANCELAMENTO DO PEDIDO, CONSIDERANDO-SE COMO SE HOUVESSE UMA NFe EMITIDA!
		dbcNFe.Close
		set dbcNFe = nothing
		exit function
		end if
		
	strFlagCancelada = Trim("" & tNFE("CANCELADA"))
	
	tNFE.Close
	set tNFE = nothing
	
	dbcNFe.Close
	set dbcNFe = nothing

'	NFE FOI CANCELADA?
	if strFlagCancelada = "1" then
		IsPedidoCancelavelNFeEmitida = True
		exit function
		end if
		
end function


' ___________________________________________________________________________________
' IsNFeCompletamenteEmitida
'
function IsNFeCompletamenteEmitida(byval strIdNfeEmitente, byval strSerieNFe, byval strNumeroNFe, ByRef ChaveAcesso, ByRef DataEmissao, ByRef DataAutorizacao)
dim s
dim tNE, tNFE
dim dbcNFe
dim strNfeT1ServidorBd, strNfeT1NomeBd, strNfeT1UsuarioBd, strNfeT1SenhaCriptografadaBd
dim chave
dim senha_decodificada
	
	IsNFeCompletamenteEmitida = False
	DataEmissao = Null
	DataAutorizacao = Null

	if Trim("" & strNumeroNFe) = "" then exit function
	
'	VERIFICA SE A NFe FOI COMPLETAMENTE EMITIDA
	s = "SELECT" & _
			" NFe_T1_servidor_BD," & _
			" NFe_T1_nome_BD," & _
			" NFe_T1_usuario_BD," & _
			" NFe_T1_senha_BD" & _
		" FROM t_NFe_EMITENTE" & _
		" WHERE" & _
			" (id = " & strIdNfeEmitente & ")"
	set tNE = cn.Execute(s)
	if tNE.Eof then exit function
	
	strNfeT1ServidorBd = Trim("" & tNE("NFe_T1_servidor_BD"))
	strNfeT1NomeBd = Trim("" & tNE("NFe_T1_nome_BD"))
	strNfeT1UsuarioBd = Trim("" & tNE("NFe_T1_usuario_BD"))
	strNfeT1SenhaCriptografadaBd = Trim("" & tNE("NFe_T1_senha_BD"))
	
	tNE.Close
	set tNE = nothing
	
	chave = gera_chave(FATOR_BD)
	decodifica_dado strNfeT1SenhaCriptografadaBd, senha_decodificada, chave
	s = "Provider=SQLOLEDB;" & _
		"Data Source=" & strNfeT1ServidorBd & ";" & _
		"Initial Catalog=" & strNfeT1NomeBd & ";" & _
		"User ID=" & strNfeT1UsuarioBd & ";" & _
		"Password=" & senha_decodificada & ";"
	set dbcNFe = server.CreateObject("ADODB.Connection")
	dbcNFe.ConnectionTimeout = 45
	dbcNFe.CommandTimeout = 900
	dbcNFe.ConnectionString = s
	dbcNFe.Open

	s = "SELECT" & _
			" Nfe," & _
			" Serie," & _
            " ChaveAcesso," & _
			" Emissao," & _
			" DataAutorizacao" & _
		" FROM NFE" & _
		" WHERE" & _
			" (Serie = '" & NFeFormataSerieNF(strSerieNFe) & "')" & _
			" AND (Nfe = '" & NFeFormataNumeroNF(strNumeroNFe) & "')" & _
			" AND (CodProcAtual = 100)"
	set tNFE = dbcNFe.Execute(s)
	if tNFE.Eof then
		dbcNFe.Close
		set dbcNFe = nothing
		exit function
    else
        ChaveAcesso = tNFE("ChaveAcesso")
		DataEmissao = tNFE("Emissao")
		DataAutorizacao = tNFE("DataAutorizacao")
		end if
		
	tNFE.Close
	set tNFE = nothing
	
	dbcNFe.Close
	set dbcNFe = nothing
	
	IsNFeCompletamenteEmitida = True
end function



' ___________________________________________________________________________________
' IsNFeCompletamenteEmitidaMontaLinkXmlNFe
'
function IsNFeCompletamenteEmitidaMontaLinkXmlNFe(byval strIdNfeEmitente, byval strSerieNFe, byval strNumeroNFe, byval html_elemento_interno_anchor, ByRef ChaveAcesso, ByRef LinkXml)
dim s
dim tNE, tNFE
dim dbcNFe
dim strNfeT1ServidorBd, strNfeT1NomeBd, strNfeT1UsuarioBd, strNfeT1SenhaCriptografadaBd
dim chave
dim senha_decodificada
dim fso
dim intIdNFeEmitente
dim strArqXml, strArqXmlAux, strArqXmlCompleto
dim strPathArqXml
dim blnArqXmlExiste
dim rNfeEmitente, iCfg
	
	IsNFeCompletamenteEmitidaMontaLinkXmlNFe = False
	
	LinkXml = ""

	if Trim("" & strNumeroNFe) = "" then exit function
	
'	VERIFICA SE A NFe FOI COMPLETAMENTE EMITIDA
	s = "SELECT" & _
			" NFe_T1_servidor_BD," & _
			" NFe_T1_nome_BD," & _
			" NFe_T1_usuario_BD," & _
			" NFe_T1_senha_BD" & _
		" FROM t_NFe_EMITENTE" & _
		" WHERE" & _
			" (id = " & strIdNfeEmitente & ")"
	set tNE = cn.Execute(s)
	if tNE.Eof then exit function
	
	strNfeT1ServidorBd = Trim("" & tNE("NFe_T1_servidor_BD"))
	strNfeT1NomeBd = Trim("" & tNE("NFe_T1_nome_BD"))
	strNfeT1UsuarioBd = Trim("" & tNE("NFe_T1_usuario_BD"))
	strNfeT1SenhaCriptografadaBd = Trim("" & tNE("NFe_T1_senha_BD"))
	
	tNE.Close
	set tNE = nothing
	
	chave = gera_chave(FATOR_BD)
	decodifica_dado strNfeT1SenhaCriptografadaBd, senha_decodificada, chave
	s = "Provider=SQLOLEDB;" & _
		"Data Source=" & strNfeT1ServidorBd & ";" & _
		"Initial Catalog=" & strNfeT1NomeBd & ";" & _
		"User ID=" & strNfeT1UsuarioBd & ";" & _
		"Password=" & senha_decodificada & ";"
	set dbcNFe = server.CreateObject("ADODB.Connection")
	dbcNFe.ConnectionTimeout = 45
	dbcNFe.CommandTimeout = 900
	dbcNFe.ConnectionString = s
	dbcNFe.Open

	s = "SELECT" & _
			" Nfe," & _
			" Serie, " & _
			" ChaveAcesso" & _
		" FROM NFE" & _
		" WHERE" & _
			" (Serie = '" & NFeFormataSerieNF(strSerieNFe) & "')" & _
			" AND (Nfe = '" & NFeFormataNumeroNF(strNumeroNFe) & "')" & _
			" AND (CodProcAtual = 100)"
	set tNFE = dbcNFe.Execute(s)
	if tNFE.Eof then
		dbcNFe.Close
		set dbcNFe = nothing
		exit function
	else
		ChaveAcesso = tNFE("ChaveAcesso")
		end if
		
	tNFE.Close
	set tNFE = nothing
	
	dbcNFe.Close
	set dbcNFe = nothing
	
	intIdNFeEmitente = CLng(strIdNfeEmitente)

	set fso = Server.CreateObject("Scripting.FileSystemObject")

	set rNfeEmitente = le_nfe_emitente(intIdNFeEmitente)

	blnArqXmlExiste = False
	for iCfg=LBound(rNfeEmitente.vCfgDanfe) to UBound(rNfeEmitente.vCfgDanfe)
		if Trim("" & rNfeEmitente.vCfgDanfe(iCfg).convencao_nome_arq_xml_nfe) <> "" then
		'	MONTA O NOME DO ARQUIVO DO XML DA NFE
			strArqXml = ""
			strArqXmlAux = Trim("" & rNfeEmitente.vCfgDanfe(iCfg).convencao_nome_arq_xml_nfe)
			if strArqXmlAux <> "" then
				strArqXml = Replace(Replace(strArqXmlAux, "[NUMERO_NFE]", strNumeroNFe), "[SERIE_NFE]", strSerieNFe)
				strPathArqXml = Trim("" & rNfeEmitente.vCfgDanfe(iCfg).diretorio_xml_nfe)
				strArqXmlCompleto = strPathArqXml & "\" & strArqXml
				blnArqXmlExiste = fso.FileExists(strArqXmlCompleto)
				end if
			end if

		if blnArqXmlExiste then exit for
		next

	set fso = nothing

	if blnArqXmlExiste then
		LinkXml = "<a name='lnkXmlNFePedido' href='../Global/DownloadXmlNFe.asp?file=" & strArqXml & "&emitente=" & Cstr(intIdNFeEmitente) & "&force=true" & "' title='Clique para consultar o XML da NFe deste pedido'>" & html_elemento_interno_anchor & "</a>"
		end if

	IsNFeCompletamenteEmitidaMontaLinkXmlNFe = True
end function



' ____________________________________________
' CODIGO DESCRICAO MONTA ITENS SELECT
'
function codigo_descricao_monta_itens_select(byval grupo, byval id_default)
dim s, x, r, strResp, ha_default
	id_default = Trim("" & id_default)
	ha_default=False
	s = "SELECT " & _
			"*" & _
		" FROM t_CODIGO_DESCRICAO" & _
		" WHERE" & _
			" (grupo='" & grupo & "')" & _
			" AND (" & _
				"(st_inativo = 0)" & _
				" OR " & _
				"(codigo = '" & id_default & "')" & _
				")" & _
		" ORDER BY" & _
			" ordenacao"
	set r = cn.Execute(s)
	strResp = ""
	do while Not r.Eof
		x = UCase(Trim("" & r("codigo")))
		if (id_default<>"") And (id_default=x) then
			strResp = strResp & "<option selected"
			ha_default=True
		else
			strResp = strResp & "<option"
			end if
		strResp = strResp & " value='" & x & "'>"
		strResp = strResp & iniciais_em_maiusculas(Trim("" & r("descricao")))
		strResp = strResp & "</option>" & chr(13)
		r.MoveNext
		loop

	if Not ha_default then
		strResp = "<option selected value=''>&nbsp;</option>" & chr(13) & strResp
		end if
		
	codigo_descricao_monta_itens_select = strResp
	r.close
	set r=nothing
end function


' ____________________________________________
' CODIGO DESCRICAO MONTA ITENS SELECT POR LOJA
'
function codigo_descricao_monta_itens_select_por_loja(byval grupo, byval id_default, byval loja)
dim s, s_where_loja, x, r, strResp, ha_default
	id_default = Trim("" & id_default)
	loja = Trim("" & loja)
	ha_default=False

	s_where_loja = "(Len(LTrim(RTrim(Coalesce(lojas_habilitadas,'')))) = 0)"
	if loja <> "" then
		if s_where_loja <> "" then s_where_loja = s_where_loja & " OR "
		s_where_loja = s_where_loja & _
							"(lojas_habilitadas LIKE '" & BD_CURINGA_TODOS & "|" & loja & "|" & BD_CURINGA_TODOS & "')"
		end if
	
	if s_where_loja <> "" then s_where_loja = " AND (" & s_where_loja & ")"

	s = "SELECT " & _
			"*" & _
		" FROM t_CODIGO_DESCRICAO" & _
		" WHERE" & _
			" (grupo='" & grupo & "')" & _
			s_where_loja & _
			" AND (" & _
				"(st_inativo = 0)" & _
				" OR " & _
				"(codigo = '" & id_default & "')" & _
				")" & _
		" ORDER BY" & _
			" ordenacao"
	set r = cn.Execute(s)
	strResp = ""
	do while Not r.Eof
		x = UCase(Trim("" & r("codigo")))
		if (id_default<>"") And (id_default=x) then
			strResp = strResp & "<option selected"
			ha_default=True
		else
			strResp = strResp & "<option"
			end if
		strResp = strResp & " value='" & x & "'>"
		strResp = strResp & iniciais_em_maiusculas(Trim("" & r("descricao")))
		strResp = strResp & "</option>" & chr(13)
		r.MoveNext
		loop

	if Not ha_default then
		strResp = "<option selected value=''>&nbsp;</option>" & chr(13) & strResp
		end if
		
	codigo_descricao_monta_itens_select_por_loja = strResp
	r.close
	set r=nothing
end function


' ____________________________________________
' CODIGO DESCRICAO MONTA ITENS SELECT ALL
'
function codigo_descricao_monta_itens_select_all(byval grupo, byval id_default)
dim s, x, r, strResp, ha_default
	id_default = Trim("" & id_default)
	ha_default=False
	s = "SELECT " & _
			"*" & _
		" FROM t_CODIGO_DESCRICAO" & _
		" WHERE" & _
			" (grupo='" & grupo & "')" & _						
		" ORDER BY" & _
			" ordenacao"
	set r = cn.Execute(s)
	strResp = ""
	do while Not r.Eof
		x = UCase(Trim("" & r("codigo")))
		if (id_default<>"") And (id_default=x) then
			strResp = strResp & "<option selected"
			ha_default=True
		else
			strResp = strResp & "<option"
			end if
		strResp = strResp & " value='" & x & "'>"
		strResp = strResp & iniciais_em_maiusculas(Trim("" & r("descricao")))
		strResp = strResp & "</option>" & chr(13)
		r.MoveNext
		loop

	if Not ha_default then
		strResp = "<option selected value=''>&nbsp;</option>" & chr(13) & strResp
		end if
		
	codigo_descricao_monta_itens_select_all = strResp
	r.close
	set r=nothing
end function


' _____________________________________________
' ZONA_DEPOSITO_MONTA_ITENS_SELECT
'
function zona_deposito_monta_itens_select(byval id_default)
dim x, r, strResp, ha_default
	id_default = Trim("" & id_default)
	ha_default=False
	set r = cn.Execute("SELECT * FROM t_WMS_DEPOSITO_MAP_ZONA WHERE (st_ativo <> 0) ORDER BY ordenacao")
	strResp = ""
	do while Not r.eof 
		x = Trim("" & r("id"))
		if (id_default<>"") And (id_default=x) then
			strResp = strResp & "<option selected"
			ha_default=True
		else
			strResp = strResp & "<option"
			end if
		strResp = strResp & " value='" & x & "'>"
		strResp = strResp & "&nbsp;&nbsp;" & Trim("" & r("zona_codigo")) & "&nbsp;&nbsp;&nbsp;"
		strResp = strResp & "</option>" & chr(13)
		r.MoveNext
		loop

	if Not ha_default then
		strResp = "<option selected value=''>&nbsp;</option>" & chr(13) & strResp
		end if
		
	zona_deposito_monta_itens_select = strResp
	r.close
	set r=nothing
end function



' _____________________________________________
' DESCRICAO COD REL TRANSACOES CIELO
'
function descricao_cod_rel_transacoes_cielo(byval codigo)
dim strResp

	codigo = Trim("" & codigo)
	
	select case codigo
		case COD_REL_TRANSACOES_CIELO__TRANSACAO_AUTORIZADA
			strResp = "Transações Autorizadas"
		case COD_REL_TRANSACOES_CIELO__TRANSACAO_NAO_AUTORIZADA
			strResp = "Transações Não Autorizadas"
		case COD_REL_TRANSACOES_CIELO__TRANSACAO_EM_SITUACAO_DESCONHECIDA
			strResp = "Transações em Situação Desconhecida"
		case COD_REL_TRANSACOES_CIELO__TRANSACAO_CANCELADA_PELO_USUARIO
			strResp = "Transações Canceladas pelo Usuário"
		case COD_REL_TRANSACOES_CIELO__TRANSACAO_EM_ANDAMENTO
			strResp = "Transações em Andamento"
		case COD_REL_TRANSACOES_CIELO__TRANSACAO_EM_AUTENTICACAO
			strResp = "Transações em Autenticação"
		case else
			strResp = "Código Desconhecido: " & codigo
	end select
	
	descricao_cod_rel_transacoes_cielo = strResp
end function



' _____________________________________________
' DESCRICAO COD REL BRASPAG AF REVIEW
'
function descricao_cod_rel_braspag_af_review(byval codigo)
dim strResp

	codigo = Trim("" & codigo)
	
	select case codigo
		case COD_REL_BRASPAG_AF_REVIEW__REVISAO_MANUAL_PENDENTE
			strResp = "Revisão Manual Pendente"
		case COD_REL_BRASPAG_AF_REVIEW__REVISAO_MANUAL_TRATADA_ACCEPT
			strResp = "Revisado para 'Aprovado'"
		case COD_REL_BRASPAG_AF_REVIEW__REVISAO_MANUAL_TRATADA_REJECT
			strResp = "Revisado para 'Rejeitado'"
		case COD_REL_BRASPAG_AF_REVIEW__APROVADO_AUTOMATICAMENTE
			strResp = "Aprovado Automaticamente"
		case COD_REL_BRASPAG_AF_REVIEW__REJEITADO_AUTOMATICAMENTE
			strResp = "Rejeitado Automaticamente"
		case else
			strResp = "Código Desconhecido: " & codigo
	end select
	
	descricao_cod_rel_braspag_af_review = strResp
end function



' _____________________________________________
' DESCRICAO COD REL TRANSACOES BRASPAG
'
function descricao_cod_rel_transacoes_braspag(byval codigo)
dim strResp

	codigo = Trim("" & codigo)
	
	select case codigo
		case COD_REL_TRANSACOES_BRASPAG__TRANSACAO_AUTORIZADA
			strResp = "Transações Autorizadas"
		case COD_REL_TRANSACOES_BRASPAG__TRANSACAO_NAO_AUTORIZADA
			strResp = "Transações Não Autorizadas"
		case COD_REL_TRANSACOES_BRASPAG__TRANSACAO_CAPTURADA
			strResp = "Transações Capturadas"
		case COD_REL_TRANSACOES_BRASPAG__TRANSACAO_CAPTURA_CANCELADA
			strResp = "Transações com Capturas Canceladas"
		case COD_REL_TRANSACOES_BRASPAG__TRANSACAO_ESTORNADA
			strResp = "Transações Estornadas"
		case COD_REL_TRANSACOES_BRASPAG__TRANSACAO_ESTORNO_PENDENTE
			strResp = "Transações com Estorno Pendente"
		case COD_REL_TRANSACOES_BRASPAG__TRANSACAO_COM_ERRO_DESQUALIFICANTE
			strResp = "Transações com Erro Desqualificante"
		case COD_REL_TRANSACOES_BRASPAG__TRANSACAO_AGUARDANDO_RESPOSTA
			strResp = "Transações Aguardando Resposta"
		case else
			strResp = "Código Desconhecido: " & codigo
	end select
	
	descricao_cod_rel_transacoes_braspag = strResp
end function



' _________________________________________________
' BRASPAG CARTAO VALIDADE MES MONTA ITENS SELECT
'
function braspag_cartao_validade_mes_monta_itens_select(byval id_default)
dim strResp, ha_default, strListaMes, strMes, vMes, intContador
	id_default = Trim("" & id_default)
	ha_default=False
	strListaMes="01|02|03|04|05|06|07|08|09|10|11|12"
	vMes=Split(strListaMes,"|")
	for intContador=LBound(vMes) to UBound(vMes)
		strMes = vMes(intContador)
		if (id_default<>"") And (converte_numero(id_default)=converte_numero(strMes)) then
			strResp = strResp & "<option selected"
			ha_default=True
		else
			strResp = strResp & "<option"
			end if
		strResp = strResp & " value=""" & strMes & """>"
		strResp = strResp & strMes
		strResp = strResp & "</option>" & chr(13)
		next

	if Not ha_default then
		strResp = "<option selected value="""">Mês</option>" & chr(13) & strResp
	else
		strResp = "<option value="""">&nbsp;Mês&nbsp;</option>" & chr(13) & strResp
		end if
	
	braspag_cartao_validade_mes_monta_itens_select = strResp
end function



' _________________________________________________
' BRASPAG CARTAO VALIDADE ANO MONTA ITENS SELECT
'
function braspag_cartao_validade_ano_monta_itens_select(byval id_default)
dim strResp, ha_default, intAnoBase, intAno, strAno, intContador
	intAnoBase = Year(Date)
	for intContador=1 to 16
		intAno = intAnoBase + (intContador - 1)
		strAno = CStr(intAno)
		if (id_default<>"") And (converte_numero(id_default)=intAno) then
			strResp = strResp & "<option selected"
			ha_default=True
		else
			strResp = strResp & "<option"
			end if
		strResp = strResp & " value=""" & strAno & """>"
		strResp = strResp & strAno
		strResp = strResp & "</option>" & chr(13)
		next
	
	if Not ha_default then
		strResp = "<option selected value="""">Ano</option>" & chr(13) & strResp
	else
		strResp = "<option value="""">&nbsp;Ano&nbsp;</option>" & chr(13) & strResp
		end if
	
	braspag_cartao_validade_ano_monta_itens_select = strResp
end function



' _________________________________________________
' OBTÉM TRANSPORTADORA PELO CEP
'
function obtem_transportadora_pelo_cep(Byval cep)
dim r, strTransp, strSql, strCep

	strCep = retorna_so_digitos(cep)
	
	strSql = "SELECT " & _
				" transportadora_id" & _
			" FROM t_TRANSPORTADORA_CEP " & _
			" WHERE"

	strSql = strSql & _
		" (" & _
			" ((tipo_range = 1) AND (cep_unico = '" & strCep & "'))" & _
			" OR" & _
			" ((tipo_range = 2) AND ('" & strCep & "' BETWEEN cep_faixa_inicial AND cep_faixa_final))" & _
		") "

'	EXECUTA A CONSULTA
	strTransp = ""
	
	set r = cn.Execute(strSql)
'	ENCONTROU DADOS NA TABELA CEPs DE ENTREGA
	if Not r.Eof then
		strTransp = Trim("" & r("transportadora_id"))
		end if

	r.Close
	Set r = Nothing

	obtem_transportadora_pelo_cep = strTransp

end function



' __________________________________________________
' WMS_NFE_EMITENTE_MONTA_ITENS_SELECT
'
function wms_nfe_emitente_monta_itens_select(byval id_default)
dim x,r,s,ha_default
	id_default = Trim("" & id_default)
	ha_default=False
	set r = cn.Execute("SELECT * FROM t_NFe_EMITENTE WHERE (st_ativo <> 0) ORDER BY id")
	s= ""
	do while Not r.eof 
		x = Trim("" & r("id"))
		if (id_default<>"") And (id_default=x) then
			s = s & "<option selected"
			ha_default=True
		else
			s = s & "<option"
			end if
		s = s & " value='" & x & "'>"
		s = s & cnpj_cpf_formata("" & r("cnpj")) & " (" & Ucase(Trim("" & r("uf"))) & ") - " & Ucase(Trim("" & r("razao_social")))
		s = s & "</option>" & chr(13)
		r.MoveNext
		loop

	if Not ha_default then
		s = "<option selected value=''>&nbsp;</option>" & chr(13) & s
		end if
	
	wms_nfe_emitente_monta_itens_select = s
	r.close
	set r=nothing
end function



' __________________________________________________
' WMS_USUARIO_X_NFE_EMITENTE_MONTA_ITENS_SELECT
'
function wms_usuario_x_nfe_emitente_monta_itens_select(byval usuario, byval id_default)
dim x,r,s,s_sql,ha_default
	usuario = Trim("" & usuario)
	id_default = Trim("" & id_default)
	ha_default=False
	s_sql = "SELECT" & _
				" tUXNE.id_nfe_emitente," & _
				" tNE.apelido" & _
			" FROM t_NFe_EMITENTE tNE" & _
				" INNER JOIN t_USUARIO_X_NFe_EMITENTE tUXNE ON (tNE.id=tUXNE.id_nfe_emitente)" & _
			" WHERE" & _
				" (usuario = '" & usuario & "')" & _
				" AND (tNE.st_ativo = 1)" & _
				" AND (tNE.st_habilitado_ctrl_estoque = 1)" & _
			" ORDER BY" & _
				" ordem"
	set r = cn.Execute(s_sql)
	s= ""
	do while Not r.Eof
		x = Trim("" & r("id_nfe_emitente"))
		if (id_default<>"") And (id_default=x) then
			s = s & "<option selected"
			ha_default=True
		else
			s = s & "<option"
			end if
		s = s & " value='" & x & "'>"
		s = s & Trim("" & r("apelido"))
		s = s & "</option>" & chr(13)
		r.MoveNext
		loop

	if Not ha_default then
		s = "<option selected value=''>&nbsp;</option>" & chr(13) & s
		end if
		
	wms_usuario_x_nfe_emitente_monta_itens_select = s
	r.close
	set r=nothing
end function

' __________________________________________________
' origem_pedido_monta_itens_select
'
function origem_pedido_monta_itens_select(byval id_default)
dim x, r, strResp
	id_default = Trim("" & id_default)

	set r = cn.Execute("SELECT * FROM t_CODIGO_DESCRICAO WHERE grupo='PedidoECommerce_Origem' AND st_inativo=0")
	strResp = ""
	do while Not r.eof 
		x = Trim("" & r("codigo"))
		if (id_default=x) then
			strResp = strResp & "<option selected"
		else
			strResp = strResp & "<option"
			end if
		strResp = strResp & " value='" & x & "'>"
		strResp = strResp & Trim("" & r("descricao"))
		strResp = strResp & "</option>" & chr(13)
		r.MoveNext
		loop
	
    strResp = "<option value=''>&nbsp;</option>" & strResp

	origem_pedido_monta_itens_select = strResp
	r.close
	set r=nothing
end function

' __________________________________________________
' monta_url_rastreio
'
function monta_url_rastreio(byval numero_pedido, byval numero_NF, byval transportadora_id, byval loja)
dim s_url, s_cnpj
dim lista_transportadora
dim rPSSW

	monta_url_rastreio = ""

	if Trim("" & numero_NF) = "" then exit function

	if Trim("" & numero_NF) = "0" then exit function

	transportadora_id = Ucase(Trim("" & transportadora_id))
	if transportadora_id = "" then exit function

	s_cnpj = obtemCnpjNFeEmitentePeloPedido(numero_pedido)
	if s_cnpj = "" then exit function

	set rPSSW = get_registro_t_parametro(ID_PARAMETRO_SSW_Rastreamento_Lista_Transportadoras)
	lista_transportadora = Trim("" & rPSSW.campo_texto)

	s_url = ""
	if InStr(lista_transportadora, "|" & transportadora_id & "|") <> 0 then
		s_url = URL_SSW_COMPLETA & s_cnpj & "/" & numero_NF
		end if

	monta_url_rastreio = s_url
end function



' __________________________________________________
' monta_link_rastreio
'
function monta_link_rastreio(byval numero_pedido, byval numero_NF, byval transportadora_id, byval loja)
dim s_url, s_link
dim rPSSW, blnConsultaViaWebAPI, sJsMethodName

	monta_link_rastreio = ""
	if Trim("" & numero_NF) = "" then exit function

	blnConsultaViaWebAPI = False
	set rPSSW = get_registro_t_parametro(ID_PARAMETRO_SSW_Rastreamento_via_WebAPI_FlagHabilitacao)
	if (Trim("" & rPSSW.id) <> "") And (rPSSW.campo_inteiro = 1) then blnConsultaViaWebAPI = True

	if blnConsultaViaWebAPI then
		sJsMethodName = "fRastreioConsultaViaWebApiView"
	else
		sJsMethodName = "fRastreioConsultaView"
		end if

	s_link = ""
	s_url = monta_url_rastreio(numero_pedido, numero_NF, transportadora_id, loja)
	if s_url <> "" then
		s_link = "<a href='javascript:" & sJsMethodName & "(" & _
					chr(34) & s_url & chr(34) & _
				");' style='cursor:default;' title='clique para consultar dados de rastreamento do pedido'>" & _
				"<img id='imgRastreioConsultaView' src='../imagem/truck_16.png' class='notPrint' />" & _
				"</a>"
		end if
	monta_link_rastreio = s_link
end function



' __________________________________________________
' monta_url_rastreio_do_emitente
'
function monta_url_rastreio_do_emitente(byval cnpj_emitente, byval numero_NF, byval transportadora_id, byval ssw_lista_transportadoras, byval loja)
dim s_url, s_cnpj
dim lista_transportadora
dim rPSSW

	monta_url_rastreio_do_emitente = ""

	if Trim("" & numero_NF) = "" then exit function

	if Trim("" & numero_NF) = "0" then exit function

	transportadora_id = Ucase(Trim("" & transportadora_id))
	if transportadora_id = "" then exit function

	s_cnpj = retorna_so_digitos(Trim("" & cnpj_emitente))
	if s_cnpj = "" then exit function

	lista_transportadora = Trim("" & ssw_lista_transportadoras)

	if lista_transportadora = "" then
		set rPSSW = get_registro_t_parametro(ID_PARAMETRO_SSW_Rastreamento_Lista_Transportadoras)
		lista_transportadora = Trim("" & rPSSW.campo_texto)
		end if

	s_url = ""
	if InStr(lista_transportadora, "|" & transportadora_id & "|") <> 0 then
		s_url = URL_SSW_COMPLETA & s_cnpj & "/" & numero_NF
		end if

	monta_url_rastreio_do_emitente = s_url
end function



' __________________________________________________
' monta_link_rastreio_do_emitente
'
function monta_link_rastreio_do_emitente(byval cnpj_emitente, byval numero_NF, byval transportadora_id, byval ssw_lista_transportadoras, byval loja)
dim s_url, s_link
dim rPSSW, blnConsultaViaWebAPI, sJsMethodName

	monta_link_rastreio_do_emitente = ""
	if Trim("" & numero_NF) = "" then exit function

	blnConsultaViaWebAPI = False
	set rPSSW = get_registro_t_parametro(ID_PARAMETRO_SSW_Rastreamento_via_WebAPI_FlagHabilitacao)
	if (Trim("" & rPSSW.id) <> "") And (rPSSW.campo_inteiro = 1) then blnConsultaViaWebAPI = True

	if blnConsultaViaWebAPI then
		sJsMethodName = "fRastreioConsultaViaWebApiView"
	else
		sJsMethodName = "fRastreioConsultaView"
		end if

	s_link = ""
	s_url = monta_url_rastreio_do_emitente(cnpj_emitente, numero_NF, transportadora_id, ssw_lista_transportadoras, loja)
	if s_url <> "" then
		s_link = "<a href='javascript:" & sJsMethodName & "(" & _
					chr(34) & s_url & chr(34) & _
				");' style='cursor:default;' title='clique para consultar dados de rastreamento do pedido'>" & _
				"<img id='imgRastreioConsultaView' src='../imagem/truck_16.png' class='notPrint' />" & _
				"</a>"
		end if
	monta_link_rastreio_do_emitente = s_link
end function



' --------------------------------------------------------------------------------
'   obtemCnpjNFeEmitentePeloPedido
'   A partir do nº do pedido, identifica e retorna o CNPJ da empresa responsável pela emissão da NF
function obtemCnpjNFeEmitentePeloPedido(ByVal pedido)
dim r, s
	obtemCnpjNFeEmitentePeloPedido = ""
	pedido = Trim("" & pedido)
	s = "SELECT " & _
			"tNE.cnpj" & _
		" FROM t_PEDIDO tP INNER JOIN t_NFe_EMITENTE tNE ON (tP.id_nfe_emitente = tNE.id)" & _
		" WHERE" & _
			" (pedido = '" & pedido & "')"
	set r = cn.Execute(s)
	if Not r.Eof then obtemCnpjNFeEmitentePeloPedido = retorna_so_digitos(Trim("" & r("cnpj")))

	if r.State <> 0 then r.Close
	set r = Nothing
end function



' __________________________________________________
' monta_funcao_js_normaliza_numero_pedido_e_sufixo
'
function monta_funcao_js_normaliza_numero_pedido_e_sufixo
dim strJs, s_letra_ano, msg_erro
	call le_ano_letra_seq_tabela_controle(NSU_PEDIDO, s_letra_ano, msg_erro)
	strJs = "function normaliza_numero_pedido_e_sufixo(pedido) {" & vbCrLf & _
			"	var s_resp;" & vbCrLf & _
			"	s_resp = ucase(trim(" & chr(34) & chr(34) & " + pedido));" & vbCrLf & _
			"	if (retorna_so_digitos(s_resp).length > " & Cstr(TAM_MIN_NUM_PEDIDO) & ") return '';" & vbCrLf & _
			"	if (s_resp.length > 0) {" & vbCrLf & _
			"		if (isDigit(s_resp.charAt(s_resp.length - 1))) {" & vbCrLf & _
			"			s_resp = s_resp + " & chr(34) & s_letra_ano & chr(34) & ";" & vbCrLf & _
			"		}" & vbCrLf & _
			"	}" & vbCrLf & _
			"	s_resp = normaliza_num_pedido(s_resp);" & vbCrLf & _
			"	return s_resp;" & vbCrLf & _
			"}" & vbCrLf
	monta_funcao_js_normaliza_numero_pedido_e_sufixo = strJs
end function



' __________________________________________________
' is_nome_e_sobrenome_iguais
'
function is_nome_e_sobrenome_iguais(byval nome_completo_1, byval nome_completo_2)
dim v1, v2
dim nome1, nome2, sobrenome1, sobrenome2
	nome_completo_1 = retira_acentuacao(Ucase(Trim("" & nome_completo_1)))
	nome_completo_2 = retira_acentuacao(Ucase(Trim("" & nome_completo_2)))
	if Instr(nome_completo_1, " ") <> 0 then
		v1 = Split(nome_completo_1, " ")
		nome1 = v1(LBound(v1))
		sobrenome1 = v1(UBound(v1))
	else
		nome1 = nome_completo_1
		sobrenome1 = ""
		end if

	if Instr(nome_completo_2, " ") <> 0 then
		v2 = Split(nome_completo_2, " ")
		nome2 = v2(LBound(v2))
		sobrenome2 = v2(UBound(v2))
	else
		nome2 = nome_completo_2
		sobrenome2 = ""
		end if

	if (nome1 = nome2) And (sobrenome1 = sobrenome2) then
		is_nome_e_sobrenome_iguais = True
	else
		is_nome_e_sobrenome_iguais = False
		end if
end function


' __________________________________________________
' isEmailOk
'
function isEmailOk(byval emailAddress)
dim objRegExpr
	Set objRegExpr = New regexp
	objRegExpr.Pattern = "^([0-9a-zA-Z]([-.\w]*[0-9a-zA-Z][_]*)*@([0-9a-zA-Z][-\w]*\.)+[a-zA-Z]{2,9})$"
	isEmailOk = objRegExpr.Test(emailAddress)
end function

' __________________________________________________
' obtem_email_usuario
'
function obtem_email_usuario(Byval usuario)
dim r, strEmail, strSql, strUsuario

	strUsuario = Trim(usuario)
	
	strSql = "SELECT email FROM t_USUARIO WHERE usuario = '" & strUsuario & "'"

'	EXECUTA A CONSULTA
	strEmail = ""
	
	set r = cn.Execute(strSql)
'	ENCONTROU DADOS NA TABELA 
	if Not r.Eof then
		strEmail = Trim("" & r("email"))
		end if

	r.Close
	Set r = Nothing

	obtem_email_usuario = strEmail

end function


' __________________________________________________
' descricao_multi_CD_regra_tipo_pessoa
'
function descricao_multi_CD_regra_tipo_pessoa(byval codigo_tipo_pessoa)
dim s_descricao
	codigo_tipo_pessoa = Ucase(Trim("" & codigo_tipo_pessoa))
	if codigo_tipo_pessoa = COD_WMS_MULTI_CD_REGRA__TIPO_PESSOA__PESSOA_FISICA then
		s_descricao = "Pessoa Física"
	elseif codigo_tipo_pessoa = COD_WMS_MULTI_CD_REGRA__TIPO_PESSOA__PRODUTOR_RURAL then
		s_descricao = "Produtor Rural"
	elseif codigo_tipo_pessoa = COD_WMS_MULTI_CD_REGRA__TIPO_PESSOA__PESSOA_JURIDICA_CONTRIBUINTE then
		s_descricao = "PJ Contribuinte"
	elseif codigo_tipo_pessoa = COD_WMS_MULTI_CD_REGRA__TIPO_PESSOA__PESSOA_JURIDICA_NAO_CONTRIBUINTE then
		s_descricao = "PJ Não Contribuinte"
	elseif codigo_tipo_pessoa = COD_WMS_MULTI_CD_REGRA__TIPO_PESSOA__PESSOA_JURIDICA_ISENTO then
		s_descricao = "PJ Isento"
	else
		s_descricao = ""
		end if

	descricao_multi_CD_regra_tipo_pessoa = s_descricao
end function


' ------------------------------------------------------------------------
'   TipoPessoa_get_array
' 
function TipoPessoa_get_array
dim v
dim sigla
	sigla = COD_WMS_MULTI_CD_REGRA__TIPO_PESSOA__PESSOA_FISICA & _
			" " & COD_WMS_MULTI_CD_REGRA__TIPO_PESSOA__PRODUTOR_RURAL & _
			" " & COD_WMS_MULTI_CD_REGRA__TIPO_PESSOA__PESSOA_JURIDICA_CONTRIBUINTE & _
			" " & COD_WMS_MULTI_CD_REGRA__TIPO_PESSOA__PESSOA_JURIDICA_NAO_CONTRIBUINTE & _
			" " & COD_WMS_MULTI_CD_REGRA__TIPO_PESSOA__PESSOA_JURIDICA_ISENTO
	v = Split(sigla, " ")
	TipoPessoa_get_array = v
end function


' ------------------------------------------------------------------------
'   converte_cst_nfe_fabricante_para_entrada_estoque
' 
function converte_cst_nfe_fabricante_para_entrada_estoque(byval cst)
dim s_resp

	cst = Trim("" & cst)

	if cst = "000" then
		s_resp = "000"
	elseif cst = "010" then
		s_resp = "060"
	elseif cst = "100" then
		s_resp = "200"
	elseif cst = "110" then
		s_resp = "260"
	elseif cst = "200" then
		s_resp = "200"
	elseif cst = "300" then
		s_resp = "200"
	elseif cst = "400" then
		s_resp = "000"
	elseif cst = "441" then
		s_resp = "000"
	elseif cst = "500" then
		s_resp = "000"
	elseif cst = "600" then
		s_resp = "200"
	elseif cst = "700" then
		s_resp = "200"
	elseif cst = "800" then
		s_resp = "000"
	elseif cst = "160" then
		s_resp = "260"
	elseif cst = "141" then
		s_resp = "241"
	else
		s_resp = ""
		end if

	converte_cst_nfe_fabricante_para_entrada_estoque = s_resp
end function


' ________________________________________________________
' ordenacao_lista_indicadores_monta_itens_select
'
function ordenacao_lista_indicadores_monta_itens_select(byval ordenacao_default)
dim i, v, x, strResp, ha_default, codigo, descricao
dim ListaOpcaoOrdenacao
	ListaOpcaoOrdenacao = Array("ID|Identificação", "UF|UF")
	ordenacao_default = Trim("" & ordenacao_default)
	ha_default=False
	strResp = ""
	for i = Lbound(ListaOpcaoOrdenacao) to Ubound(ListaOpcaoOrdenacao)
		x = Trim("" & ListaOpcaoOrdenacao(i))
		v = Split(x, "|")
		codigo = Trim("" & v(LBound(v)))
		descricao = Trim("" & v(LBound(v)+1))
		if (ordenacao_default<>"") And (ordenacao_default=codigo) then
			strResp = strResp & "<option selected"
			ha_default=True
		else
			strResp = strResp & "<option"
			end if
		strResp = strResp & " value='" & codigo & "'>"
		strResp = strResp & descricao
		strResp = strResp & "</option>" & chr(13)
		next

	if Not ha_default then
		strResp = "<option selected value=''>&nbsp;</option>" & chr(13) & strResp
		end if
		
	ordenacao_lista_indicadores_monta_itens_select = strResp
end function


' ________________________________________________________
' exibe_alerta_nf_cancelada
'
function exibe_alerta_nf_cancelada(byval pedido, byval obs_1)
dim x, s_obs_1
	pedido = Trim("" & pedido)
	s_obs_1 = Ucase(Trim("" & obs_1))

	if (Instr(s_obs_1, "#NFCANCEL") = 0) And (Instr(s_obs_1, "#NF CANCEL") = 0) then exit function

	x = "<table width='649' class='Q' cellspacing='0' style='3pt solid red;'>" & chr(13) & _
		"	<tr style='background-color:red;'>" & chr(13) & _
		"		<td align='center'>" & _
					"<span style='background-color:red;color:yellow;font-size:12pt;font-weight:bold;'>ATENÇÃO!</span>" & _
					"<br />" & _
					"<span style='background-color:red;color:yellow;font-size:12pt;font-weight:bold;'>Pedido com NF CANCELADA!</span>" & _
				"</td>" & chr(13) & _
		"	</tr>" & chr(13) & _
		"</table>" & chr(13) & _
		"<br />" & chr(13) & _
		chr(13)

	Response.Write x

end function
%>
