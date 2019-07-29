<%

'	==================================================================================================================================================================
'	Descrição:
'	~~~~~~~~~~
'	qtde_parcelas_loja: oferece opção de pagamento parcelado Lojista com cartão de crédito. 0=não, número de parcelas (com juros à cargo da loja)
'	qtde_parcelas_emissor: oferece opção de pagamento parcelado Emissor com cartão de crédito. 0=não, número de parcelas (com juros à cargo do cliente)
'	==================================================================================================================================================================
Function CieloGeraOpcoesPagto(Byval valor, Byval bandeira, Byval qtde_parcelas_loja, Byval vl_min_parcela_loja, Byval qtde_parcelas_emissor, Byval vl_min_parcela_emissor, texto_obs_juros)

Dim x, n, qtde_parcelas_loja_ok
dim credito_habilitado, debito_habilitado

	if bandeira = CIELO_BANDEIRA__VISA then
		credito_habilitado = True
		debito_habilitado = False
	elseif bandeira = CIELO_BANDEIRA__MASTERCARD then
		credito_habilitado = True
		debito_habilitado = False
	elseif bandeira = CIELO_BANDEIRA__AMEX then
		credito_habilitado = True
		debito_habilitado = False
	elseif bandeira = CIELO_BANDEIRA__ELO then
		credito_habilitado = True
		debito_habilitado = False
	elseif bandeira = CIELO_BANDEIRA__DINERS then
		credito_habilitado = True
		debito_habilitado = False
	elseif bandeira = CIELO_BANDEIRA__DISCOVER then
		credito_habilitado = True
		debito_habilitado = False
	elseif bandeira = CIELO_BANDEIRA__AURA then
		credito_habilitado = True
		debito_habilitado = False
	elseif bandeira = CIELO_BANDEIRA__JCB then
		credito_habilitado = True
		debito_habilitado = False
	elseif bandeira = CIELO_BANDEIRA__CELULAR then
		credito_habilitado = True
		debito_habilitado = True
	else
		credito_habilitado = False
		debito_habilitado = False
		end if

	texto_obs_juros = ""
	
	qtde_parcelas_loja_ok = 0
%>


	<% if debito_habilitado then %>
			<input type="radio" name="rb_forma_pagamento" value="A"><%=SIMBOLO_MONETARIO%>&nbsp;<%=formata_moeda(valor)%> à Vista (no Débito)<br>
	<% end if %>
	
	<% if credito_habilitado then %>
			<input type="radio" name="rb_forma_pagamento" value="1"><%=SIMBOLO_MONETARIO%>&nbsp;<%=formata_moeda(valor)%> à Vista (no Crédito)<br>
	<% end if %>


<%	if qtde_parcelas_loja > 1 Then
		for x = 2 to qtde_parcelas_loja 
			if (valor / x) < vl_min_parcela_loja then exit for 
			qtde_parcelas_loja_ok = x
%>
			<input type="radio" name="rb_forma_pagamento" value="PL|<%=Cstr(x)%>"> <%=x%> x de <%=SIMBOLO_MONETARIO%>&nbsp;<%=formata_moeda(valor/x)%> iguais<br>
<%		next
	end if
%>


<%	if qtde_parcelas_emissor > 1 Then
		n = 0
		for x = 2 to qtde_parcelas_emissor 
			if x > qtde_parcelas_loja_ok then
				if (valor / x) < vl_min_parcela_emissor then exit for
				n = n + 1
				if n = 1 then 
					if bandeira = CIELO_BANDEIRA__VISA Then
						texto_obs_juros = "* Verificar a taxa de juros praticada junto ao emissor do seu cartão Visa."
					elseif bandeira = CIELO_BANDEIRA__MASTERCARD Then
						texto_obs_juros = "* Verificar a taxa de juros praticada junto ao emissor do seu cartão Mastercard."
					elseif bandeira = CIELO_BANDEIRA__AMEX Then
						texto_obs_juros = "* Verificar a taxa de juros praticada junto ao emissor do seu cartão Amex."
					elseif bandeira = CIELO_BANDEIRA__ELO Then
						texto_obs_juros = "* Verificar a taxa de juros praticada junto ao emissor do seu cartão ELO."
					elseif bandeira = CIELO_BANDEIRA__DINERS Then
						texto_obs_juros = "* Verificar a taxa de juros praticada junto ao emissor do seu cartão Diners."
					elseif bandeira = CIELO_BANDEIRA__DISCOVER Then
						texto_obs_juros = "* Verificar a taxa de juros praticada junto ao emissor do seu cartão Discover."
					elseif bandeira = CIELO_BANDEIRA__AURA Then
						texto_obs_juros = "* Verificar a taxa de juros praticada junto ao emissor do seu cartão Aura."
					elseif bandeira = CIELO_BANDEIRA__JCB Then
						texto_obs_juros = "* Verificar a taxa de juros praticada junto ao emissor do seu cartão JCB."
						end if
					end if
%>
			<input type="radio" name="rb_forma_pagamento" value="PC|<%=Cstr(x)%>"><font color="red"> <%=x%> x de <%=SIMBOLO_MONETARIO%>&nbsp;<%=formata_moeda(valor/x)%> mais juros*</font><br>
<%			end if
		next
	end if
%>

<%
End Function
%>