<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =========================================================
'	  P E S Q U I S A D E I N D I C A D O R E S E X E C . A S P
'     =========================================================
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

	Const COD_PESQUISAR_POR_UF_LOCALIDADE = "POR_UF_LOCALIDADE"
	Const COD_PESQUISAR_POR_BAIRRO = "POR_BAIRRO"
	Const COD_PESQUISAR_POR_CEP = "POR_CEP"
	Const COD_PESQUISAR_POR_NOME = "POR_NOME"
	Const COD_PESQUISAR_POR_CPF_CNPJ = "POR_CPF_CNPJ"
	Const COD_PESQUISAR_ASSOCIADOS_AO_VENDEDOR = "POR_ASSOCIADOS_AO_VENDEDOR"
	
	dim usuario, loja
	usuario = Trim(Session("usuario_atual"))
	loja = Trim(Session("loja_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 
	If (loja = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	CONECTA COM O BANCO DE DADOS
	dim cn, rs, msg_erro
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	if Not operacao_permitida(OP_LJA_PESQUISA_INDICADORES, s_lista_operacoes_permitidas) then 
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	dim alerta
	dim s, s_filtro, s_lista_loja_bs, s_lista_loja_aux
	dim rb_pesquisar_por, c_loja, c_uf_pesq, c_localidade_pesq, c_cep_pesq, c_indicador, c_vendedor, c_cpfcnpj
	dim c_bairro, c_uf_bairro, c_cidade_bairro
	

	alerta = ""

	rb_pesquisar_por = Trim(Request("rb_pesquisar_por"))
	c_uf_pesq = Trim(Request("c_uf_pesq"))
	c_localidade_pesq = Trim(Request("c_localidade_pesq"))
	c_cep_pesq = retorna_so_digitos(Trim(Request("c_cep_pesq")))
	c_uf_bairro = Trim(Request("uf_bairro"))
	c_cidade_bairro = Trim(Request("cidade_bairro"))
	c_bairro = Trim(Request("bairro_pesq"))
	c_indicador = Trim(Request("c_indicador"))
	c_cpfcnpj = retorna_so_digitos(Trim(Request("c_cpfcnpj_pesq")))
	c_vendedor = Trim(Request("c_vendedor"))
	
	if (rb_pesquisar_por = "") then
		alerta = "Nenhum parâmetro de pesquisa foi fornecido."
	elseif (rb_pesquisar_por <> COD_PESQUISAR_POR_UF_LOCALIDADE) And _
	       (rb_pesquisar_por <> COD_PESQUISAR_POR_BAIRRO) And _
		   (rb_pesquisar_por <> COD_PESQUISAR_POR_CEP) And _
		   (rb_pesquisar_por <> COD_PESQUISAR_POR_NOME) And _
		   (rb_pesquisar_por <> COD_PESQUISAR_POR_CPF_CNPJ) And _
		   (rb_pesquisar_por <> COD_PESQUISAR_ASSOCIADOS_AO_VENDEDOR) then
		alerta = "Parâmetro de pesquisa selecionado é inválido."
		end if
	
	if alerta = "" then
		if rb_pesquisar_por = COD_PESQUISAR_POR_UF_LOCALIDADE then
			if c_uf_pesq = "" then
				alerta = "Informe o UF a ser pesquisado."
				end if
			end if
		end if
		
	if alerta = "" then
	    if rb_pesquisar_por = COD_PESQUISAR_POR_BAIRRO then
	        if c_uf_bairro = "" then
	            alerta = "Selecione a UF do bairro."
	        end if
	        if c_cidade_bairro = "" then
	            alerta = "Selecione a cidade do bairro."
	        end if
	        if c_bairro = "" then
	            alerta = "Escoha ao menos 1 bairro para a pesquisa."
	        end if
	    end if
	end if

	if alerta = "" then
		if rb_pesquisar_por = COD_PESQUISAR_POR_CEP then
			if c_cep_pesq = "" then
				alerta = "Informe o CEP a ser pesquisado."
			elseif (Len(c_cep_pesq) <> 5)  And (Len(c_cep_pesq) <> 8) then
				alerta = "CEP informado possui tamanho inválido."
				end if
			end if
		end if
		
	if alerta = "" then
		if rb_pesquisar_por = COD_PESQUISAR_POR_NOME then
			if c_indicador = "" then
				alerta = "Selecione um indicador da lista."
				end if
			end if
		end if
		
    if alerta = "" then
        if rb_pesquisar_por = COD_PESQUISAR_POR_CPF_CNPJ then
            if c_cpfcnpj = "" then
                alerta = "Informe o CPF ou CNPJ a ser pesquisado."
            end if
        end if
    end if

	if alerta = "" then
		if rb_pesquisar_por = COD_PESQUISAR_ASSOCIADOS_AO_VENDEDOR then
			if c_vendedor = "" then
				alerta = "Selecione um vendedor da lista."
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
dim s_sql, s_where, s_cep_sql
dim n_reg, n_reg_total
dim intLargApelido, intLargNome, intLargTelefone, intLargLoja, intLargCidade, intLargStatus
dim strCidade, strUF, strLoja, strStatus, strVendedor
dim strDdd, strTelefone, strDddCel, strTelCel, strListaTelefones
dim v_cidades, s_where_temp, v_bairros, bairros_temp, j, i

	s_lista_loja_bs = ""
	s_sql = "SELECT loja FROM t_LOJA WHERE (unidade_negocio = '" & COD_UNIDADE_NEGOCIO_LOJA__BS & "') ORDER BY loja"
	if rs.State <> 0 then rs.Close
	rs.open s_sql, cn
	do while Not rs.Eof
		if s_lista_loja_bs <> "" then s_lista_loja_bs = s_lista_loja_bs & ", "
		s_lista_loja_bs = s_lista_loja_bs & Trim("" & rs("loja"))
		rs.MoveNext
		loop

	j = ""
	v_bairros = ""
'	CRITÉRIOS COMUNS
	s_where = ""
	
'   POR UF/LOCALIDADE

    if rb_pesquisar_por = COD_PESQUISAR_POR_UF_LOCALIDADE then
        if c_uf_pesq <> "" then
            if s_where <> "" then s_where = s_where & " AND"
            s_where = s_where & " (t_ORCAMENTISTA_E_INDICADOR.uf = '" & c_uf_pesq & "')"
        end if
        if Not operacao_permitida(OP_LJA_CONSULTA_UNIVERSAL_PEDIDO_ORCAMENTO, s_lista_operacoes_permitidas) then
                if s_where <> "" then s_where = s_where & " AND"
                s_where = s_where & " (t_ORCAMENTISTA_E_INDICADOR.vendedor='" & usuario & "')"
            end if

        s_where_temp = ""
        if c_localidade_pesq <> "" then
            v_cidades = split(c_localidade_pesq, ", ")
            for i = LBound(v_cidades) to Ubound(v_cidades)
                if s_where_temp <> "" then s_where_temp = s_where_temp & " OR"
                s_where_temp = s_where_temp & " (t_ORCAMENTISTA_E_INDICADOR.cidade = '" & trim(replace(v_cidades(i), "'", "''")) & "' COLLATE Latin1_General_CI_AI)"
            next
            if s_where_temp <> "" then
                s_where_temp = " AND (" & s_where_temp & ") "
                s_where = s_where & s_where_temp
            end if
        end if
    end if

'	SE A CONSULTA FOR POR CEP, TENTA LOCALIZAR O MAIS PRÓXIMO. SE NÃO ENCONTRAR, TORNA A REGIÃO
'	CADA VEZ MAIS ABRANGENTE
	if rb_pesquisar_por = COD_PESQUISAR_POR_CEP then
		s_cep_sql = mid(c_cep_pesq, 1, 5)
		do while Len(s_cep_sql) >= 2
			s_sql = "SELECT " & _
						" Coalesce(Count(*), 0) AS qtde" & _
					" FROM t_ORCAMENTISTA_E_INDICADOR" & _
					" WHERE" & _
						" (cep LIKE '" & s_cep_sql & BD_CURINGA_TODOS & "')" & _
						" AND"
			if CStr(loja) = CStr(NUMERO_LOJA_ECOMMERCE_AR_CLUBE) then
				s_lista_loja_aux = loja
				if s_lista_loja_bs <> "" then s_lista_loja_aux = s_lista_loja_aux & ", " & s_lista_loja_bs
				s_sql = s_sql & _
						" (CONVERT(smallint, loja) IN (" & s_lista_loja_aux & "))"
			else
				s_sql = s_sql & _
						" (CONVERT(smallint, loja) = " & loja & ")"
				end if
			
			if rs.State <> 0 then rs.Close
			rs.open s_sql, cn
			
			if Not rs.Eof then
				if CLng(rs("qtde")) > 0 then exit do
				end if
			
			s_cep_sql = Mid(s_cep_sql, 1, Len(s_cep_sql)-1)
			Loop
				
		if s_cep_sql <> "" then
			if s_where <> "" then s_where = s_where & " AND"
			s_where = s_where & " (t_ORCAMENTISTA_E_INDICADOR.cep LIKE '" & s_cep_sql & BD_CURINGA_TODOS & "')"
        if Not operacao_permitida(OP_LJA_CONSULTA_UNIVERSAL_PEDIDO_ORCAMENTO, s_lista_operacoes_permitidas) then
                if s_where <> "" then s_where = s_where & " AND"
                s_where = s_where & " (t_ORCAMENTISTA_E_INDICADOR.vendedor='" & usuario & "')"
            end if
			end if
		end if
		
'   POR BAIRRO
    if rb_pesquisar_por = COD_PESQUISAR_POR_BAIRRO then
        if c_uf_bairro <> "" then
            if s_where <> "" then s_where = s_where & " AND"
            s_where = s_where & " (t_ORCAMENTISTA_E_INDICADOR.uf = '" & c_uf_bairro & "')"
        end if
       if Not operacao_permitida(OP_LJA_CONSULTA_UNIVERSAL_PEDIDO_ORCAMENTO, s_lista_operacoes_permitidas) then
                if s_where <> "" then s_where = s_where & " AND"
                s_where = s_where & " (t_ORCAMENTISTA_E_INDICADOR.vendedor='" & usuario & "')"
            end if
        if c_cidade_bairro <> "" then
            if s_where <> "" then s_where = s_where & " AND"
            s_where = s_where & " (t_ORCAMENTISTA_E_INDICADOR.cidade = '" & c_cidade_bairro & "' COLLATE Latin1_General_CI_AI)"
        end if
        bairros_temp = ""
        if c_bairro <> "" then
            v_bairros = split(c_bairro, ", ")
            for j = LBound(v_bairros) to UBound(v_bairros)
                if bairros_temp <> "" then bairros_temp = bairros_temp & " OR"
                bairros_temp = bairros_temp & " (t_ORCAMENTISTA_E_INDICADOR.bairro = '" & Trim(replace(v_bairros(j), "'", "''")) & "' COLLATE Latin1_General_CI_AI)"
            next
            if bairros_temp <> "" then
                bairros_temp = " AND (" & bairros_temp & ") "
                s_where = s_where & bairros_temp
            end if
        end if
    end if
    
'	POR NOME
	if rb_pesquisar_por = COD_PESQUISAR_POR_NOME then
		if c_indicador <> "" then
			if s_where <> "" then s_where = s_where & " AND"
			s_where = s_where & " (t_ORCAMENTISTA_E_INDICADOR.apelido = '" & c_indicador & "')"
			end if
        if Not operacao_permitida(OP_LJA_CONSULTA_UNIVERSAL_PEDIDO_ORCAMENTO, s_lista_operacoes_permitidas) then
                if s_where <> "" then s_where = s_where & " AND"
                s_where = s_where & " (t_ORCAMENTISTA_E_INDICADOR.vendedor='" & usuario & "')"
            end if
		end if

'	POR CPF/CNPJ
	if rb_pesquisar_por = COD_PESQUISAR_POR_CPF_CNPJ then
		if c_cpfcnpj <> "" then
			if s_where <> "" then s_where = s_where & " AND"
			s_where = s_where & " (t_ORCAMENTISTA_E_INDICADOR.cnpj_cpf = '" & c_cpfcnpj & "')"
			end if
        if Not operacao_permitida(OP_LJA_CONSULTA_UNIVERSAL_PEDIDO_ORCAMENTO, s_lista_operacoes_permitidas) then
                if s_where <> "" then s_where = s_where & " AND"
                s_where = s_where & " (t_ORCAMENTISTA_E_INDICADOR.vendedor='" & usuario & "')"
            end if
		end if
		
'	ASSOCIADOS AO VENDEDOR
	if rb_pesquisar_por = COD_PESQUISAR_ASSOCIADOS_AO_VENDEDOR then
		if c_vendedor <> "" then
			if s_where <> "" then s_where = s_where & " AND"
			s_where = s_where & " (t_ORCAMENTISTA_E_INDICADOR.vendedor = '" & c_vendedor & "')"
			end if
        if Not operacao_permitida(OP_LJA_CONSULTA_UNIVERSAL_PEDIDO_ORCAMENTO, s_lista_operacoes_permitidas) then
                if s_where <> "" then s_where = s_where & " AND"
                s_where = s_where & " (t_ORCAMENTISTA_E_INDICADOR.vendedor='" & usuario & "')"
            end if
		end if
		
'	APENAS INDICADORES DESTA LOJA
	s_lista_loja_aux = loja
	if s_lista_loja_bs <> "" then s_lista_loja_aux = s_lista_loja_aux & ", " & s_lista_loja_bs

	if s_where <> "" then s_where = s_where & " AND"
	if CStr(loja) = CStr(NUMERO_LOJA_ECOMMERCE_AR_CLUBE) then
		s_where = s_where & _
						" (CONVERT(smallint, loja) IN (" & s_lista_loja_aux & "))"
	else
		s_where = s_where & _
						" (CONVERT(smallint, loja) = " & loja & ")"
		end if
	
'	MONTA SQL DE CONSULTA
	s_sql = "SELECT " & _
				"apelido, " & _
				"razao_social_nome_iniciais_em_maiusculas, " & _
				"ddd, telefone, " & _
				"ddd_cel, tel_cel, " & _
				"nextel, " & _
				"loja, " & _
				"cidade, uf, " & _
				"status, " & _
				"vendedor, " & _
				"(" & _
					"SELECT " & _
						" Coalesce(Count(*), 0)" & _
					" FROM t_PEDIDO" & _
					" WHERE" & _
						" indicador = t_ORCAMENTISTA_E_INDICADOR.apelido" & _
				") AS QtdePedidos" & _
			" FROM t_ORCAMENTISTA_E_INDICADOR" & _
			" WHERE" & _
				s_where & _
			" ORDER BY" & _
				" QtdePedidos DESC, " & _
				" apelido"
	
	
  ' CABEÇALHO
	intLargApelido = 80
	intLargNome = 210
	intLargTelefone = 95
	intLargLoja = 30
	intLargCidade = 157
	intLargStatus = 40
	
	cab_table = "<TABLE cellSpacing=0>" & chr(13)
	
	cab = _
		"	<TR style='background:azure' NOWRAP>" & chr(13) & _
		"		<TD class='MDBE MC' valign='bottom' NOWRAP><P style='width:" & CStr(intLargApelido) & "px' class='R'>Indicador</P></TD>" &  chr(13) & _
		"		<TD class='MD MB MC' valign='bottom'><P style='width:" & CStr(intLargNome) & "px' class='R'>Nome</P></TD>" & chr(13) & _
		"		<TD class='MD MB MC' valign='bottom'><P style='width:" & CStr(intLargTelefone) & "px' class='R'>Telefone</P></TD>" & chr(13) & _
		"		<TD class='MD MB MC' valign='bottom'><P style='width:" & CStr(intLargLoja) & "px' class='R'>Loja</P></TD>" & chr(13) & _
		"		<TD class='MD MB MC' valign='bottom'><P style='width:" & CStr(intLargApelido) & "px' class='R'>Vendedor</P></TD>" & chr(13) & _
		"		<TD class='MD MB MC' valign='bottom'><P style='width:" & CStr(intLargCidade) & "px' class='R'>Cidade</P></TD>" & chr(13) & _
		"		<TD class='MD MB MC' valign='bottom'><P style='width:" & CStr(intLargStatus) & "px' class='R'>Status</P></TD>" & chr(13) & _
        "       <TD style='background-color:#fff'>&nbsp;</TD>" & _
		"	</TR>" & chr(13)
	
'	LAÇO P/ LEITURA DO RECORDSET
	x = cab_table & _
		cab
	n_reg = 0
	n_reg_total = 0
	set r = cn.execute(s_sql)
	do while Not r.Eof
	
		n_reg = n_reg + 1
		n_reg_total = n_reg_total + 1

		x = x & "	<TR>" & chr(13)
		
	'>  INDICADOR (APELIDO)
		x = x & "		<TD class='MDBE' valign='top' style='width:" & CStr(intLargApelido) & "px'>" & _
							"<P class='C'>" & _
								"<a href='javascript:fOPConsultar(" & chr(34) & r("apelido") & chr(34) & ")' title='clique para consultar o cadastro'>" & _
								Trim("" & r("apelido")) & _
							"</P>" & _
						"</TD>" & chr(13)

	'>  INDICADOR (NOME)
		x = x & "		<TD class='MD MB' valign='top' style='width:" & CStr(intLargNome) & "px'>" & _
							"<P class='Cn'>" & _
								"<a href='javascript:fOPConsultar(" & chr(34) & r("apelido") & chr(34) & ")' title='clique para consultar o cadastro'>" & _
								Trim("" & r("razao_social_nome_iniciais_em_maiusculas")) & _
							"</P>" & _
						"</TD>" & chr(13)

	'>  TELEFONE
		strDdd = Trim("" & r("ddd"))
		strTelefone = Trim("" & r("telefone"))
		strDddCel = Trim("" & r("ddd_cel"))
		strTelCel = Trim("" & r("tel_cel"))
		if strTelefone <> "" then strTelefone = formata_ddd_telefone_ramal(strDdd, strTelefone, "")
		if strTelCel <> "" then strTelCel = formata_ddd_telefone_ramal(strDddCel, strTelCel, "")
		strListaTelefones = strTelefone
		if (strListaTelefones <> "") And (strTelCel <> "") then strListaTelefones = strListaTelefones & "<br>"
		strListaTelefones = strListaTelefones & strTelCel
		if Trim("" & r("nextel")) <> "" then
			if strListaTelefones <> "" then strListaTelefones = strListaTelefones & "<br>"
			strListaTelefones = strListaTelefones & Trim("" & r("nextel"))
			end if
		
		if strListaTelefones = "" then strListaTelefones = "&nbsp;"
		x = x & "		<TD class='MD MB' valign='top' style='width:" & CStr(intLargTelefone) & "px'>" & _
							"<P class='Cn'>" & strListaTelefones & "</P>" & _
						"</TD>" & chr(13)

	'>  LOJA
		strLoja = Trim("" & r("loja"))
		if strLoja = "" then strLoja = "&nbsp;"
		x = x & "		<TD class='MD MB' valign='top' style='width:" & CStr(intLargLoja) & "px'>" & _
							"<P class='Cn'>" & strLoja & "</P>" & _
						"</TD>" & chr(13)
						
	'>  VENDEDOR
	    strVendedor = Trim("" & r("vendedor"))
	        if strVendedor = "" then 
	        strVendedor = "&nbsp;"
	        end if
	        x = x & "   <TD class='MD MB' valign='top' style='width:" & CStr(intLargApelido) & "px'>" & _
	                        "<P class='Cn'>" & strVendedor & "</P>" & _
	                    "</TD>" & chr(13)
	    

	'>  CIDADE
		strCidade = iniciais_em_maiusculas(Trim("" & r("cidade")))
		strUF = Trim("" & r("uf"))
		if (strCidade <> "") And (strUF <> "") then 
			strCidade = strCidade & " / " & strUF 
		else 
			strCidade = strCidade & strUF
			end if
		if strCidade = "" then strCidade = "&nbsp;"
		x = x & "		<TD class='MD MB' valign='top' style='width:" & CStr(intLargCidade) & "px'>" & _
							"<P class='Cn'>" & strCidade & "</P>" & _
						"</TD>" & chr(13)

	'>  STATUS
 		if Trim("" & r("status"))="A" then 
 			strStatus = "<span style='color:#006600'>Ativo</span>"
 		else 
 			strStatus = "<span style='color:#ff0000'>Inativo</span>"
 			end if
		
		x = x & "		<TD class='MD MB' valign='top' style='width:" & CStr(intLargStatus) & "px'>" & _
							"<P class='Cn'>" & strStatus & "</P>" & _
						"</TD>" & chr(13)
    
    '>  CONSULTA/EDITA
		x=x & " <TD valign='middle' style='border:0'><a href='javascript:fOPConsultar(""" & r("apelido") & """)'><img src='../imagem/lupa_20x20.png' style='border:0;width:18px;height:18px' title='Consultar cadastro'></a>"
		x=x & " <a href='javascript:fOPEditar(""" & r("apelido") & """)'><img src='../imagem/edita_20x20.gif' style='border:0;width:20px;height:20px' title='Editar cadastro'></a></TD>"
		
        x=x & "</TR>" & chr(13)
		
		if (n_reg mod 100) = 0 then
			Response.Write x
			x = ""
			end if
		
		r.MoveNext
		loop


  ' MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	if n_reg_total = 0 then
		x = cab_table & _
			cab & _
			"	<TR NOWRAP>" & chr(13) & _
			"		<TD class='MT' colspan=7><P class='ALERTA'>&nbsp;NENHUM INDICADOR ENCONTRADO&nbsp;</P></TD>" & chr(13) & _
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

function fOPConsultar(s_id) {
    window.status = "Aguarde ...";
    fREL.id_selecionado.value = s_id;
    fREL.action = "OrcamentistaEIndicadorConsulta.asp";
    fREL.submit();
}
function fOPEditar(s_id) {
    window.status = "Aguarde ...";
    fREL.id_selecionado.value = s_id;
    fREL.action = "OrcamentistaEIndicadorEdita.asp";
    fREL.submit();
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

<form id="fREL" name="fREL" method="post" action="OrcamentistaEIndicadorConsulta.asp" >
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="c_uf_pesq" id="c_uf_pesq" value="<%=c_uf_pesq%>">
<input type="hidden" name="c_localidade_pesq" id="c_localidade_pesq" value="<%=c_localidade_pesq%>">
<input type="hidden" name="c_cep_pesq" id="c_cep_pesq" value="<%=c_cep_pesq%>">
<input type="hidden" name="id_selecionado" id="id_selecionado" value=''>
<input type="hidden" name="operacao_selecionada" id="operacao_selecionada" value='<%=OP_CONSULTA%>'>
<input type="hidden" name="url_origem" id="url_origem" value="PesquisaDeIndicadoresFiltro.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" />
<input type="hidden" name="url_partida_pesq_ind" id="url_partida_pesq_ind" value="X" />

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="740" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Pesquisa de Indicadores</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>

<!-- FILTROS -->
<% 
	s_filtro = "<table width='740' cellPadding='0' CellSpacing='0' style='border-bottom:1px solid black' border='0'>"

	s = loja
	if s = "" then s = "todas"
	s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
			   "<p class='N'>Loja:&nbsp;</p></td><td valign='top'>" & _
			   "<p class='N'>" & s & "</p></td></tr>"

	if rb_pesquisar_por = COD_PESQUISAR_POR_UF_LOCALIDADE then
		s = "UF / Localidade"
	elseif rb_pesquisar_por = COD_PESQUISAR_POR_CEP then
		s = "CEP"
	elseif rb_pesquisar_por = COD_PESQUISAR_POR_NOME then
		s = "Nome"
	elseif rb_pesquisar_por = COD_PESQUISAR_POR_CPF_CNPJ then
	    s = "CPF/CNPJ"
	elseif rb_pesquisar_por = COD_PESQUISAR_POR_BAIRRO then
	    s = "Bairro"
	elseif rb_pesquisar_por = COD_PESQUISAR_ASSOCIADOS_AO_VENDEDOR then
		s = "Associados ao Vendedor"
	else
		s = ""
		end if
	s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
			   "<p class='N'>Pesquisar por:&nbsp;</p></td><td valign='top'>" & _
			   "<p class='N'>" & s & "</p></td></tr>"
	
	if rb_pesquisar_por = COD_PESQUISAR_POR_UF_LOCALIDADE then
		s = c_uf_pesq
		if s = "" then s = "todos"
		s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
				   "<p class='N'>UF:&nbsp;</p></td><td valign='top'>" & _
				   "<p class='N'>" & s & "</p></td></tr>"
	
		s = c_localidade_pesq
		if s = "" then s = "todos"
		s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
				   "<p class='N'>Localidade:&nbsp;</p></td><td valign='top'>" & _
				   "<p class='N'>" & s & "</p></td></tr>"
    end if
    
    if rb_pesquisar_por = COD_PESQUISAR_POR_BAIRRO then
        s = c_uf_bairro
        if s = "" then s = "todos"
        s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
                   "<p class='N'>UF:&nbsp;</p></td><td valign='top'>" & _
                   "<p class='N'>" & s & "</p></td></tr>"
                   
        s = c_cidade_bairro
        if s = "" then s = "todos"
        s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
                   "<p class='N'>Cidade:&nbsp;</p></td><td valign='top'>" & _
                   "<p class='N'>" & s & "</p></td></tr>"
           
        s = c_bairro
        if s = "" then s = "todos"
        s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
                   "<p class='N'>Bairros:&nbsp;</p></td><td valign='top'>" & _
                   "<p class='N'>" & s & "</p></td></tr>"
    end if

	if rb_pesquisar_por = COD_PESQUISAR_POR_CEP then	
		if c_cep_pesq <> "" then s = cep_formata(c_cep_pesq) else s = ""
		if s = "" then s = "todos"
		s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
				   "<p class='N'>CEP:&nbsp;</p></td><td valign='top'>" & _
				   "<p class='N'>" & s & "</p></td></tr>"
		end if

	if rb_pesquisar_por = COD_PESQUISAR_POR_NOME then	
		if c_indicador <> "" then s = c_indicador else s = ""
		if s = "" then s = "todos"
		s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
				   "<p class='N'>Indicador:&nbsp;</p></td><td valign='top'>" & _
				   "<p class='N'>" & s & "</p></td></tr>"
		end if
		
	if rb_pesquisar_por = COD_PESQUISAR_POR_CPF_CNPJ then	
		if c_cpfcnpj <> "" then s = c_cpfcnpj else s = ""
		if s = "" then s = "todos"
		s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
				   "<p class='N'>CPF/CNPJ:&nbsp;</p></td><td valign='top'>" & _
				   "<p class='N'>" & s & "</p></td></tr>"
		end if

	if rb_pesquisar_por = COD_PESQUISAR_ASSOCIADOS_AO_VENDEDOR then	
		if c_vendedor <> "" then s = c_vendedor else s = ""
		if s = "" then s = "todos"
		s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
				   "<p class='N'>Vendedor:&nbsp;</p></td><td valign='top'>" & _
				   "<p class='N'>" & s & "</p></td></tr>"
		end if
	
	s_filtro = s_filtro & "<tr><td align='right' valign='top' NOWRAP>" & _
			   "<p class='N'>Emissão:&nbsp;</p></td><td valign='top' width='99%'>" & _
			   "<p class='N'>" & formata_data_hora(Now) & "</p></td></tr>"

	s_filtro = s_filtro & "</table>"
	Response.Write s_filtro
%>

<!--  RELATÓRIO  -->
<br>
<% consulta_executa %>

<!-- ************   SEPARADOR   ************ -->
<table width="740" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>
<br>


<table class="notPrint" width="740" cellSpacing="0">
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
