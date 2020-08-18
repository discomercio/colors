<%
' =========================================
'          F  U  N  �  �  E  S
' =========================================


' ----------------------------
'   DESCRI��O DO ERRO
' 
Function erro_descricao(id_erro)
dim s

	select case id
	    case ERR_CONEXAO
	        s = "N�O FOI POSS�VEL REALIZAR A CONEX�O COM O BANCO DE DADOS."
		case ERR_IDENTIFICACAO
		    s = "OS DADOS INFORMADOS NA IDENTIFICA��O EST�O INCORRETOS."
		case ERR_IDENTIFICACAO_LOJA
			s = "OS DADOS INFORMADOS EST�O INCORRETOS.<BR>(Loja inv�lida)"
		case ERR_SESSAO
			s = "SESS�O EXPIRADA"
			s = s + "<BR>"
			s = s + "<i>A sess�o expirou porque passaram-se mais de " & Cstr(SESSION_CTRL_TIMEOUT_SESSAO_MIN) & " minutos sem enviar dados para o servidor."
			s = s + "<BR>"
			s = s + "� necess�rio realizar nova identifica��o do usu�rio.</i>"
		case ERR_SENHA_INVALIDA
			s = "OS DADOS INFORMADOS PARA AUTENTICA��O EST�O INCORRETOS."
		case ERR_SENHA_NAO_INFORMADA
			s = "SENHA N�O FOI INFORMADA."
		case ERR_ACESSO_INSUFICIENTE
			s = "N�VEL DE ACESSO INSUFICIENTE."
		case ERR_USUARIO_NAO_ESPECIFICADO
			s = "N�O FOI SELECIONADO NENHUM USU�RIO PARA EDI��O."
		case ERR_LOJA_NAO_ESPECIFICADA
			s = "N�O FOI SELECIONADA NENHUMA LOJA PARA EDI��O."        
		case ERR_OPERACAO_NAO_ESPECIFICADA
			s = "N�O FOI ESPECIFICADA A OPERA��O A SER REALIZADA."
		case ERR_USUARIO_JA_CADASTRADO
			s = "USU�RIO J� CADASTRADO."
		case ERR_USUARIO_NAO_CADASTRADO
			s = "USU�RIO N�O CADASTRADO."
		case ERR_LOJA_JA_CADASTRADA
			s = "LOJA J� CADASTRADA."
		case ERR_LOJA_NAO_CADASTRADA
			s = "LOJA N�O CADASTRADA."
		case ERR_USUARIO_BLOQUEADO
			s = "ACESSO NEGADO."
		case ERR_FALHA_OPERACAO_BD
			s = "FALHA NA OPERA��O COM O BANCO DE DADOS."
		case ERR_FABRICANTE_NAO_ESPECIFICADO
			s = "N�O FOI SELECIONADO NENHUM FABRICANTE PARA EDI��O."
		case ERR_FABRICANTE_JA_CADASTRADO
			s = "FABRICANTE J� CADASTRADO."
		case ERR_FABRICANTE_NAO_CADASTRADO
			s = "FABRICANTE N�O CADASTRADO."
		case ERR_MIDIA_NAO_ESPECIFICADA
			s = "N�O FOI SELECIONADO NENHUM VE�CULO DE M�DIA PARA EDI��O."
		case ERR_MIDIA_JA_CADASTRADA
			s = "VE�CULO DE M�DIA J� CADASTRADO."
		case ERR_MIDIA_NAO_CADASTRADA
			s = "VE�CULO DE M�DIA N�O CADASTRADO."
		case ERR_CLIENTE_NAO_ESPECIFICADO
			s = "N�O FOI SELECIONADO NENHUM CLIENTE PARA EDI��O."
		case ERR_CLIENTE_JA_CADASTRADO
			s = "CLIENTE J� CADASTRADO."
		case ERR_CLIENTE_NAO_CADASTRADO
			s = "CLIENTE N�O CADASTRADO."
		case ERR_CLIENTE_FALHA_RECUPERAR_DADOS
			s = "FALHA AO TENTAR RECUPERAR OS DADOS DO CLIENTE."
		case ERR_FALHA_OPERACAO_GERAR_NSU
			s = "FALHA NA OPERA��O COM O BANCO DE DADOS AO TENTAR GERAR NSU."
		case ERR_NSU_NAO_LOCALIZADO
			s = "N�O FOI ENCONTRADO NO BANCO DE DADOS O REGISTRO COM O NSU ESPECIFICADO."
		case ERR_NSU_JA_EM_USO
			s = "O NSU J� EST� SENDO USADO POR OUTRO REGISTRO NO BANCO DE DADOS."
		case ERR_ID_INVALIDO
			s = "O N�MERO IDENTIFICADOR DO REGISTRO NO BANCO DE DADOS � INV�LIDO."
		case ERR_CONSULTAR_ESTOQUE
			s = "FALHA AO CONSULTAR DADOS DO ESTOQUE."
		case ERR_FALHA_OPERACAO_MOVIMENTO_ESTOQUE
			s = "FALHA NA MOVIMENTA��O DE PRODUTOS NO ESTOQUE."
		case ERR_FALHA_OPERACAO_GRAVAR_LOG_ESTOQUE
			s = "FALHA AO GRAVAR O LOG DA MOVIMENTA��O NO ESTOQUE."
		case ERR_PEDIDO_NAO_ESPECIFICADO
			s = "N�O FOI ESPECIFICADO NENHUM N�MERO DE PEDIDO."
		case ERR_PEDIDO_INVALIDO
			s = "N�MERO DE PEDIDO INV�LIDO."
		case ERR_PEDIDO_ACESSO_NEGADO
			s = "ACESSO NEGADO PARA CONSULTAR ESTE PEDIDO."
		case ERR_FALHA_GERAR_ID_PEDIDO_FILHOTE
			s = "FALHA AO TENTAR GERAR UM NOVO N�MERO IDENTIFICADOR PARA O FILHOTE DE PEDIDO."
		case ERR_FALHA_OPERACAO_CRIAR_ADO
			s = "FALHA AO TENTAR CRIAR UM OBJETO ADO."
		case ERR_ESTOQUE_NAO_ESPECIFICADO
			s = "LOTE DO ESTOQUE N�O ESPECIFICADO."
		case ERR_PAG_DEST_INDEFINIDA
			s = "P�GINA DE DESTINO INDEFINIDA."
		case ERR_TIT_REL_INDEFINIDO
			s = "T�TULO DO RELAT�RIO INDEFINIDO."
		case ERR_GRUPO_LOJAS_NAO_ESPECIFICADO
			s = "N�O FOI SELECIONADO NENHUM GRUPO DE LOJAS PARA EDI��O."
		case ERR_GRUPO_LOJAS_JA_CADASTRADO
			s = "GRUPO DE LOJAS J� CADASTRADO."
		case ERR_GRUPO_LOJAS_NAO_CADASTRADO
			s = "GRUPO DE LOJAS N�O CADASTRADO."
		case ERR_ORCAMENTO_NAO_ESPECIFICADO
			s = "N�O FOI ESPECIFICADO NENHUM N�MERO DE OR�AMENTO."
		case ERR_ORCAMENTO_INVALIDO
			s = "N�MERO DE OR�AMENTO INV�LIDO."
		case ERR_CNPJ_CPF_INVALIDO
			s = "CNPJ/CPF FORNECIDO � INV�LIDO."
		case ERR_TIPO_CARTAO_INVALIDO
			s = "TIPO DE CART�O INV�LIDO."
		case ERR_OPCAO_PAGTO_INVALIDA
			s = "OP��O DE PAGAMENTO INV�LIDA."
		case ERR_VISANET_TID_NAO_ESPECIFICADO
			s = "N� DA TRANSA��O (TID) N�O ESPECIFICADO."
		case ERR_TRANSPORTADORA_NAO_ESPECIFICADA
			s = "N�O FOI SELECIONADA NENHUMA TRANSPORTADORA PARA EDI��O."
		case ERR_TRANSPORTADORA_JA_CADASTRADA
			s = "TRANSPORTADORA J� CADASTRADA."
		case ERR_TRANSPORTADORA_NAO_CADASTRADA
			s = "TRANSPORTADORA N�O CADASTRADA."
		case ERR_PERFIL_NAO_ESPECIFICADO
			s = "N�O FOI SELECIONADO NENHUM PERFIL PARA EDI��O."
		case ERR_PERFIL_JA_CADASTRADO
			s = "PERFIL J� CADASTRADO."
		case ERR_PERFIL_NAO_CADASTRADO
			s = "PERFIL N�O CADASTRADO."
		case ERR_ORCAMENTISTA_INDICADOR_NAO_ESPECIFICADO
			s = "N�O FOI SELECIONADO NENHUM OR�AMENTISTA / INDICADOR PARA EDI��O."
		case ERR_ORCAMENTISTA_INDICADOR_JA_CADASTRADO
			s = "OR�AMENTISTA / INDICADOR J� CADASTRADO."
		case ERR_ORCAMENTISTA_INDICADOR_NAO_CADASTRADO
			s = "OR�AMENTISTA / INDICADOR N�O CADASTRADO."
		case ERR_ID_JA_EM_USO_POR_USUARIO
			s = "IDENTIFICA��O J� EST� EM USO POR UM USU�RIO CONVENCIONAL DO SISTEMA."
		case ERR_ID_JA_EM_USO_POR_ORCAMENTISTA
			s = "IDENTIFICA��O J� EST� EM USO POR UM OR�AMENTISTA / INDICADOR."
		case ERR_ORDEM_SERVICO_NAO_CADASTRADA
			s = "ORDEM DE SERVI�O N�O CADASTRADA."
		case ERR_CEP_NAO_ESPECIFICADO
			s = "CEP N�O FOI INFORMADO"
		case ERR_CEP_INVALIDO
			s = "CEP COM TAMANHO INV�LIDO"
		case ERR_IDENTIFICADOR_NAO_FORNECIDO
			s = "IDENTIFICADOR N�O FOI FORNECIDO."
		case ERR_REGISTRO_NAO_CADASTRADO
			s = "REGISTRO N�O CADASTRADO."
		case ERR_IDENTIFICADOR_JA_CADASTRADO
			s = "IDENTIFICADOR J� CADASTRADO."
		case ERR_IDENTIFICADOR_NAO_CADASTRADO
			s = "IDENTIFICADOR N�O CADASTRADO"
		case ERR_ID_NAO_INFORMADO
			s = "N� DE IDENTIFICA��O N�O FOI INFORMADO."
		case ERR_ID_JA_CADASTRADO
			s = "N� DE IDENTIFICA��O J� EST� CADASTRADO."
		case ERR_ID_NAO_CADASTRADO
			s = "N� DE IDENTIFICA��O N�O EST� CADASTRADO."
		case ERR_FIN_NATUREZA_OPERACAO_NAO_ESPECIFICADO
			s = "NATUREZA DA OPERA��O N�O ESPECIFICADA."
		case ERR_CAD_CLIENTE_ENDERECO_NUMERO_NAO_PREENCHIDO
			s = "N�MERO DO ENDERE�O N�O EST� PREENCHIDO NO CADASTRO DO CLIENTE."
		case ERR_CAD_CLIENTE_ENDERECO_EXCEDE_TAMANHO_MAXIMO
			s = "ENDERE�O NO CADASTRO DO CLIENTE EXCEDE O TAMANHO M�XIMO."
		case ERR_NIVEL_ACESSO_INSUFICIENTE
			s = "N�VEL DE ACESSO INSUFICIENTE."
		case ERR_PARAMETRO_OBRIGATORIO_NAO_ESPECIFICADO
			s = "PAR�METRO OBRIGAT�RIO N�O FOI ESPECIFICADO."
		case ERR_HORARIO_MANUTENCAO_SISTEMA
			s = "Sistema em manuten��o no per�odo das " & HORARIO_INICIO_MANUTENCAO_SISTEMA & " at� " & HORARIO_TERMINO_MANUTENCAO_SISTEMA & "<br /><br />Por favor, acesse novamente mais tarde."
		case ERR_HORARIO_REBOOT_SERVIDOR
			s = "Sistema indispon�vel no per�odo das " & HORARIO_INICIO_REBOOT_SERVIDOR & " at� " & HORARIO_TERMINO_REBOOT_SERVIDOR & "<br /><br />Por favor, acesse novamente mais tarde."
		case ERR_CIELO_VALOR_PAGTO_NAO_INFORMADO
			s = "O valor do pagamento n�o foi informado."
		case ERR_CIELO_FORMA_PAGTO_NAO_INFORMADO
			s = "A forma de pagamento n�o foi informada."
		case ERR_CIELO_FORMA_PAGTO_INVALIDO
			s = "A forma de pagamento informada � inv�lida."
		case ERR_CIELO_QTDE_PARCELAS_INVALIDA
			s = "A quantidade de parcelas informada � inv�lida."
		case ERR_BRASPAG_VALOR_PAGTO_NAO_INFORMADO
			s = "O valor do pagamento n�o foi informado."
		case ERR_BRASPAG_FORMA_PAGTO_NAO_INFORMADO
			s = "A forma de pagamento n�o foi informada."
		case ERR_BRASPAG_FORMA_PAGTO_INVALIDO
			s = "A forma de pagamento informada � inv�lida."
		case ERR_BRASPAG_QTDE_PARCELAS_INVALIDA
			s = "A quantidade de parcelas informada � inv�lida."
		case ERR_BRASPAG_NOME_TITULAR_NAO_INFORMADO
			s = "O nome do titular do cart�o n�o foi informado."
		case ERR_BRASPAG_NOME_TITULAR_INVALIDO
			s = "O nome do titular do cart�o � inv�lido."
		case ERR_BRASPAG_NUMERO_CARTAO_NAO_INFORMADO
			s = "O n�mero do cart�o n�o foi informado."
		case ERR_BRASPAG_NUMERO_CARTAO_COM_TAMANHO_INVALIDO
			s = "O n�mero do cart�o possui tamanho inv�lido."
		case ERR_BRASPAG_VALIDADE_MES_NAO_INFORMADO
			s = "O m�s da validade do cart�o n�o foi informado."
		case ERR_BRASPAG_VALIDADE_ANO_NAO_INFORMADO
			s = "O ano da validade do cart�o n�o foi informado."
		case ERR_BRASPAG_CODIGO_SEGURANCA_NAO_INFORMADO
			s = "O c�digo de seguran�a do cart�o n�o foi informado."
		case ERR_BRASPAG_PERGUNTA_CARTAO_PROPRIO_NAO_RESPONDIDA
			s = "N�o foi informado se o cart�o pertence ao comprador do pedido."
		case ERR_BRASPAG_FATURA_ENDERECO_LOGRADOURO_NAO_INFORMADO
			s = "O endere�o da fatura n�o foi informado."
		case ERR_BRASPAG_FATURA_ENDERECO_NUMERO_NAO_INFORMADO
			s = "O n�mero do endere�o da fatura n�o foi informado."
		case ERR_BRASPAG_FATURA_ENDERECO_UF_NAO_INFORMADA
			s = "A UF do endere�o da fatura n�o foi informada."
		case ERR_BRASPAG_FATURA_CEP_NAO_INFORMADO
			s = "O CEP do endere�o da fatura n�o foi informado."
		case ERR_BRASPAG_FATURA_CEP_COM_TAMANHO_INVALIDO
			s = "O CEP do endere�o da fatura possui tamanho inv�lido."
		case ERR_BRASPAG_FATURA_ENDERECO_CIDADE_NAO_INFORMADO
			s = "A cidade do endere�o da fatura n�o foi informada."
		case ERR_BRASPAG_FATURA_TEL_PAIS_NAO_INFORMADO
			s = "O c�digo do pa�s n�o foi informado no n�mero de telefone."
		case ERR_BRASPAG_FATURA_TEL_PAIS_INVALIDO
			s = "O c�digo do pa�s do n�mero de telefone � inv�lido."
		case ERR_BRASPAG_FATURA_TEL_DDD_NAO_INFORMADO
			s = "O DDD n�o foi informado."
		case ERR_BRASPAG_FATURA_TEL_DDD_INVALIDO
			s = "O DDD informado � inv�lido."
		case ERR_BRASPAG_FATURA_TEL_NUMERO_NAO_INFORMADO
			s = "O n�mero de telefone n�o foi informado."
		case ERR_BRASPAG_FATURA_TEL_NUMERO_INVALIDO
			s = "O n�mero de telefone informado � inv�lido."
		case ERR_EC_PRODUTO_COMPOSTO_NAO_ESPECIFICADO
			s = "PRODUTO COMPOSTO N�O FOI ESPECIFICADO."
		case ERR_EC_PRODUTO_COMPOSTO_JA_CADASTRADO
			s = "PRODUTO COMPOSTO J� CADASTRADO."
		case ERR_EC_PRODUTO_COMPOSTO_NAO_CADASTRADO
			s = "PRODUTO COMPOSTO N�O CADASTRADO."
        case ERR_EC_PRODUTO_PRE_LISTA_CADASTRADO
            s = "PRODUTO J� CADASTRADO NA PR� LISTA"
        case ERR_EC_PRODUTO_COMPOSTO_ITEM_JA_CADASTRADO
            s = "N�MERO DO PRODUTO COMPOSTO N�O PODE SER O MESMO QUE DE UM ITEM DE OUTRO PRODUTO COMPOSTO"
		case ERR_PRODUTO_NAO_ESPECIFICADO
			s = "N�O FOI SELECIONADO NENHUM PRODUTO PARA EDI��O."
		case ERR_PRODUTO_COMPOSTO_JA_CADASTRADO
			s = "PRODUTO COMPOSTO J� CADASTRADO."
		case ERR_PRODUTO_COMPOSTO_NAO_CADASTRADO
			s = "PRODUTO COMPOSTO N�O CADASTRADO."
		case ERR_PRODUTO_COMPOSTO_NAO_ESPECIFICADO
			s = "N�O FOI SELECIONADO NENHUM PRODUTO COMPOSTO PARA EDI��O."
		case ERR_NENHUM_CD_HABILITADO_PARA_USUARIO
			s = "NENHUM CD HABILITADO PARA O USU�RIO"
        case ERR_INDICADORES_VENDEDOR_INFORMADO_JA_PROCESSADO
            s = "O VENDEDOR J� FOI PROCESSADO NO M�S DE COMPET�NCIA INFORMADO."
		case ERR_QTDE_CARTOES_INVALIDA
			s = "QUANTIDADE DE CART�ES � INV�LIDA"
		case ERR_MULTI_CD_REGRA_NAO_ESPECIFICADA
			s = "REGRA PARA CONSUMO DO ESTOQUE N�O FOI ESPECIFICADA"
		case ERR_MULTI_CD_REGRA_APELIDO_NAO_INFORMADO
			s = "APELIDO PARA A REGRA DO CONSUMO DO ESTOQUE N�O FOI INFORMADO"
		case ERR_MULTI_CD_REGRA_JA_CADASTRADA
			s = "REGRA PARA CONSUMO DO ESTOQUE J� CADASTRADA"
		case ERR_MULTI_CD_REGRA_NAO_CADASTRADA
			s = "REGRA PARA CONSUMO DO ESTOQUE N�O CADASTRADA"
		case else
			s = "ERRO"
			if (id_erro <> "") AND (id_erro <> "0") then s = s + " (c�digo: " + id_erro + ")"
			s = s + "!!"
		end select
	
	erro_descricao = s
	
End Function 



' ---------------------------
'   GERA CHAVE (CRIPTOGRAFIA)
' 
Function gera_chave(fator)

    Const COD_MINIMO = 35
    Const COD_MAXIMO = 96
    Const TAMANHO_CHAVE = 128

    Dim i 
    Dim k 
    Dim s 

    s = ""
    For i = 1 To TAMANHO_CHAVE
        k = (COD_MAXIMO - COD_MINIMO) + 1
        k = (k * fator)
        k = (k * i) + COD_MINIMO
        k = k Mod 128
        s = s & Chr(k)
        Next 

    gera_chave = s

End Function




' ---------------------------
'   CODIFICA DADO PELA CHAVE
' 
Sub codifica_dado(byval origem, destino, byval chave)

    Dim i 
    Dim i_chave 
    Dim i_dado 
    Dim k 
    Dim s
    Dim s_origem 
    Dim s_destino 


    destino = ""
    s_destino = ""
    s_origem = Trim(origem)
    If len(s_origem) > 15 then s_origem = Left(s_origem,15)
     
    For i = 1 To Len(s_origem)
        i_chave = (Asc(Mid(chave, i, 1)) * 2) + 1
        i_dado = Asc(Mid(s_origem, i, 1)) * 2
        k = i_chave Xor i_dado
        s_destino = s_destino & Chr(k)
        Next

    s_origem = s_destino
    s_destino = ""
    For i = 1 To Len(s_origem)
        k = Asc(Mid(s_origem, i, 1))
        s = Hex(k)
        While Len(s) < 2: s = "0" & s: Wend
        s_destino = s_destino & s
        Next
        
    While Len(s_destino) < 30: s_destino = "0" & s_destino: Wend 
    s_destino = "0x" & LCase(s_destino)
    
    destino = s_destino

End Sub



' -----------------------------
'   DECODIFICA DADO PELA CHAVE
' 
Sub decodifica_dado(byval origem, destino, byval chave)

    Dim i 
    Dim i_chave 
    Dim i_dado 
    Dim k 
    Dim s
    Dim s_origem 
    Dim s_destino 

     
    destino = ""
    s_destino = ""
    s_origem = Trim("" & origem)

    i = Len(s_origem) - 2
    if i < 0 then i = 0
    s_origem = Right(s_origem, i) 
    s_origem = UCase(s_origem)
    For i = 1 To Len(s_origem) Step 2
        s = Mid(s_origem, i, 2)
        If s <> "00" Then
            s_origem = Right(s_origem, Len(s_origem) - (i - 1))
            Exit For
            End If
        Next
        
    For i = 1 To Len(s_origem) Step 2
        s = Mid(s_origem, i, 2)
        s = "&H" & s
        s_destino = s_destino & Chr(s)
        Next
    
    s_origem = s_destino
    s_destino = ""
    For i = 1 To Len(s_origem)
        i_chave = (Asc(Mid(chave, i, 1)) * 2) + 1
        i_dado = Asc(Mid(s_origem, i, 1))
        k = i_chave Xor i_dado
        k = k \ 2
        s_destino = s_destino & Chr(k)
        Next

    destino = s_destino

End Sub



' -----------------------------
'   CRIPTOGRAFA
' 
function criptografa(byval texto)
dim strChave
	texto = "" & texto
	strChave = gera_chave(FATOR_CRIPTO_SESSION_CTRL)
	criptografa = CriptografaTexto(texto, strChave)
end function



' -----------------------------
'   DECRIPTOGRAFA
' 
function decriptografa(byval texto)
dim strChave
	texto = "" & texto
	strChave = gera_chave(FATOR_CRIPTO_SESSION_CTRL)
	decriptografa = DecriptografaTexto(texto, strChave)
end function



' -----------------------------
'   CRIPTOGRAFA TEXTO
' 
Function CriptografaTexto(byval origem, byval chave)

    Dim i 
    Dim i_chave 
    Dim i_dado 
    Dim k 
    Dim s
    Dim s_origem 
    Dim s_destino 

    s_destino = ""
    s_origem = Trim(origem)
     
    While Len(chave) < Len(s_origem): chave = chave & chave : Wend
    
    For i = 1 To Len(s_origem)
        i_chave = (Asc(Mid(chave, i, 1)) * 2) + 1
        i_dado = Asc(Mid(s_origem, i, 1)) * 2
        k = i_chave Xor i_dado
        s_destino = s_destino & Chr(k)
        Next

    s_origem = s_destino
    s_destino = ""
    For i = 1 To Len(s_origem)
        k = Asc(Mid(s_origem, i, 1))
        s = Hex(k)
        While Len(s) < 2: s = "0" & s: Wend
        s_destino = s_destino & s
        Next
        
    While Len(s_destino) < 30: s_destino = "0" & s_destino: Wend 
    s_destino = "0x" & LCase(s_destino)
    
    CriptografaTexto = s_destino

End Function



' -----------------------------
'   DECRIPTOGRAFA TEXTO
' 
Function DecriptografaTexto(byval origem, byval chave)

    Dim i 
    Dim i_chave 
    Dim i_dado 
    Dim k 
    Dim s
    Dim s_origem 
    Dim s_destino 
    
    s_destino = ""
    s_origem = Trim("" & origem)

    i = Len(s_origem) - 2
    if i < 0 then i = 0
    s_origem = Right(s_origem, i) 
    s_origem = UCase(s_origem)
    For i = 1 To Len(s_origem) Step 2
        s = Mid(s_origem, i, 2)
        If s <> "00" Then
            s_origem = Right(s_origem, Len(s_origem) - (i - 1))
            Exit For
            End If
        Next
        
    For i = 1 To Len(s_origem) Step 2
        s = Mid(s_origem, i, 2)
        s = "&H" & s
        s_destino = s_destino & Chr(s)
        Next
    
    s_origem = s_destino
    
    While Len(chave) < Len(s_origem): chave = chave & chave : Wend
    
    s_destino = ""
    For i = 1 To Len(s_origem)
        i_chave = (Asc(Mid(chave, i, 1)) * 2) + 1
        i_dado = Asc(Mid(s_origem, i, 1))
        k = i_chave Xor i_dado
        k = k \ 2
        s_destino = s_destino & Chr(k)
        Next

    DecriptografaTexto = s_destino

End Function



' _______________________
' X _ S A U D A C A O
'
function x_saudacao
dim h
	
	h = Cdbl(Time)
	
	if h > (18 * 1/24) then
		x_saudacao = "Boa noite"
	elseif h > (12 * 1/24) then
		x_saudacao = "Boa tarde"
	else
		x_saudacao = "Bom dia"
		end if
		
end function



' _______________________________________
' IS HORARIO MANUTENCAO SISTEMA
'
function isHorarioManutencaoSistema
dim dtInicio, dtTermino
	isHorarioManutencaoSistema = False
	if Not TRATAR_HORARIO_MANUTENCAO_SISTEMA then exit function
	dtInicio = StrToDateTime(formata_data(Date) & " " & HORARIO_INICIO_MANUTENCAO_SISTEMA)
	dtTermino = StrToDateTime(formata_data(Date) & " " & HORARIO_TERMINO_MANUTENCAO_SISTEMA)
'	VERIFICA SE O HOR�RIO DE T�RMINO EST� NO MESMO DIA QUE O IN�CIO (VIROU O DIA?)
	if (dtTermino < dtInicio) then
	'	HOR�RIO DE T�RMINO AINDA EST� VIGENTE?
		if Now <= dtTermino then
		'	ENT�O A DATA DE IN�CIO FOI ONTEM
			dtInicio = DateAdd("d", -1, dtInicio)
		else
		'	O HOR�RIO DE T�RMINO DE HOJE J� PASSOU, ENT�O O PR�XIMO � AMANH�
			dtTermino = DateAdd("d", 1, dtTermino)
			end if
		end if
		
	if (Now > dtInicio) And (Now < dtTermino) then isHorarioManutencaoSistema = True
end function



' _______________________________________
' IS HORARIO REBOOT SERVIDOR
'
function isHorarioRebootServidor
dim dtInicio, dtTermino
	isHorarioRebootServidor = False
	if Not TRATAR_HORARIO_REBOOT_SERVIDOR then exit function
	dtInicio = StrToDateTime(formata_data(Date) & " " & HORARIO_INICIO_REBOOT_SERVIDOR)
	dtTermino = StrToDateTime(formata_data(Date) & " " & HORARIO_TERMINO_REBOOT_SERVIDOR)
'	VERIFICA SE O HOR�RIO DE T�RMINO EST� NO MESMO DIA QUE O IN�CIO (VIROU O DIA?)
	if (dtTermino < dtInicio) then
	'	HOR�RIO DE T�RMINO AINDA EST� VIGENTE?
		if Now <= dtTermino then
		'	ENT�O A DATA DE IN�CIO FOI ONTEM
			dtInicio = DateAdd("d", -1, dtInicio)
		else
		'	O HOR�RIO DE T�RMINO DE HOJE J� PASSOU, ENT�O O PR�XIMO � AMANH�
			dtTermino = DateAdd("d", 1, dtTermino)
			end if
		end if
		
	if (Now > dtInicio) And (Now < dtTermino) then isHorarioRebootServidor = True
end function



' _______________________________________
' G E R A _ S E N H A _ A L E A T O R I A
'
function gera_senha_aleatoria
dim c, i, s, s_pwd
	s_pwd = ""
	do while len(s_pwd) < 10
		randomize
		i = Int((122 - 48 + 1) * Rnd + 48)
		if (i >= 48) And (i <= 57) then
			c = chr(i)
		elseif (i >= 65) And (i <= 90) then
			c = chr(i)
		elseif (i >= 97) And (i <= 122) then
			c = chr(i)
		else
			c = ""
			end if
			
		s_pwd = s_pwd & c
		loop
		
	gera_senha_aleatoria = s_pwd
end function



' _______________________________________
' IsDigit
'
function IsDigit(byval c)
	c = Trim("" & c)
	IsDigit = ((c>="0") AND (c<="9"))
end function



' _______________________________________
' IsLetra
'
function IsLetra(byval c)
	c = Ucase(Trim("" & c))
	IsLetra = ((c>="A") AND (c<="Z"))
end function



' _______________________________________
' RETORNA_SO_DIGITOS
'
Function retorna_so_digitos(ByVal s_numero)
Dim s
Dim c
Dim i
    s = ""
    For i = 1 To Len(s_numero)
        c = Mid(s_numero, i, 1)
        If IsNumeric(c) Then s = s & c
        Next
    retorna_so_digitos = s
End Function



' _______________________________________
' E X T R A C T F I L E N A M E
'
function ExtractFileName(nome_arquivo, remover_extensao)
dim s_resp
dim s
dim i
    
    s_resp = ""
    for i = Len(nome_arquivo) To 1 Step -1
        s = Mid(nome_arquivo, i, 1)
        if s <> "/" then
            s_resp = s & s_resp
        else
            exit for
            end if
        next
     
  '�REMOVE A EXTENS�O TAMB�M?
    if remover_extensao And (InStr(s_resp, ".") <> 0) then
        do while Len(s_resp) > 0
            s = right(s_resp, 1)
            s_resp = left(s_resp, Len(s_resp) - 1)
            if s = "." then exit do
            loop
        end if
       
    ExtractFileName = s_resp
    
end function



' _______________________________________
' RETORNA_DESCRICAO_NIVEL
'
function retorna_descricao_nivel(codigo_nivel)
dim s_resp
	
	select case codigo_nivel
		case ID_VENDEDOR
			s_resp = "Vendedor"
		case ID_SEPARADOR
			s_resp = "Separador"
		case ID_ADMINISTRADOR 
			s_resp = "Administrador"
		case ID_GERENCIAL 
			s_resp = "Gerencial"
		case else
			s_resp=""
		end select
		
	retorna_descricao_nivel=s_resp
end function



' ------------------------------------------------------------
'   SUBSTITUI O CARACTER ESPECIFICADO PELO NOVO
' 
Function substitui_caracteres(byval Texto, byval antigo, byval novo)
Dim i
Dim s

    substitui_caracteres = ""
    
    s = ""
    For i = 1 To Len(Texto)
        If Mid(Texto, i, 1) = antigo Then
           If novo <> "" Then If Asc(novo) <> 0 Then s = s & novo
        Else
           s = s & Mid(Texto, i, 1)
           End If
        Next
    
    substitui_caracteres = s

End Function



' ------------------------------------------------------------------------
'   CONCATENA O TEXTO-BASE A QUANTIDADE DE VEZES ESPECIFICADA
' 
Function repete_texto(ByVal texto_base, ByVal n_vezes)
dim s
dim i
    repete_texto = ""
    s = ""
    For i = 1 To n_vezes
        s = s & texto_base
        Next
    repete_texto = s
End Function



' ------------------------------------------------------------------------
'   PAD LEFT
' 
Function PadLeft(byval texto_base, byval tamanho)
dim s_resp
	s_resp = "" & texto_base
	if len(s_resp) < tamanho then s_resp = Space(tamanho-len(s_resp)) & texto_base
	PadLeft = s_resp
End Function



' ------------------------------------------------------------------------
'   PAD RIGHT
' 
Function PadRight(byval texto_base, byval tamanho)
dim s_resp
	s_resp = "" & texto_base
	if len(s_resp) < tamanho then s_resp = texto_base & Space(tamanho-len(s_resp))
	PadRight = s_resp
End Function



' ------------------------------------------------------------------------
'   FILTRA CARACTERES PROIBIDOS PARA CAMPOS USADOS COMO IDENTIFICADORES
' 
Function filtra_nome_identificador(byval nome)
Dim s
	s = nome
	s = substitui_caracteres(s, chr(34), "")
	s = substitui_caracteres(s, chr(39), "")
	s = substitui_caracteres(s, "|", "")
	filtra_nome_identificador = s
End Function



' ------------------------------------------------------------------------
'   CEP OK?
' 
Function cep_ok(byval cep)
dim s_cep
	cep_ok = False
	s_cep = "" & cep
	s_cep = retorna_so_digitos(s_cep)
	if ((len(s_cep)=0) Or (len(s_cep)=5) Or (len(s_cep)=8)) then cep_ok = True
End Function



' ------------------------------------------------------------------------
'   DDD OK?
' 
Function ddd_ok(byval ddd)
dim s_ddd
	ddd_ok = False
	s_ddd = "" & ddd
	s_ddd = retorna_so_digitos(s_ddd)
	if ((len(s_ddd)=0) Or (len(s_ddd)=2)) then ddd_ok = True
End Function



' ------------------------------------------------------------------------
'   TELEFONE OK?
' 
Function telefone_ok(byval telefone)
dim s_tel
	telefone_ok = False
	s_tel = "" & telefone
	s_tel = retorna_so_digitos(s_tel)
	if ((len(s_tel)=0) Or (len(s_tel)>=6)) then telefone_ok = True
End Function



' ------------------------------------------------------------------------
'   UF_get_array
' 
function UF_get_array
const sigla = "AC AL AM AP BA CE DF ES GO MA MG MS MT PA PB PE PI PR RJ RN RO RR RS SC SE SP TO"
dim v
	v = Split(sigla, " ")
	UF_get_array = v
end function



' ------------------------------------------------------------------------
'   UF OK?
' 
Function uf_ok(ByVal uf)
const sigla = "AC AL AM AP BA CE DF ES GO MA MG MS MT PA PB PE PI PR RJ RN RO RR RS SC SE SP TO  "

	uf_ok = False

    uf = Trim(uf)
    If uf = "" Then
        UF_ok = True
        Exit Function
        End If

    If Len(uf) <> 2 Then
        UF_ok = False
        Exit Function
        End If

    UF_ok = (InStr(sigla, uf) <> 0)

End Function



' ------------------------------------------------------------------------
'   UF DESCRICAO
'
function UF_descricao(byval uf)
dim s_resp

	UF_descricao = ""
	uf = UCase(Trim("" & uf))

	select case uf
		case "AC": s_resp = "Acre"
		case "AL": s_resp = "Alagoas"
		case "AM": s_resp = "Amazonas"
		case "AP": s_resp = "Amap�"
		case "BA": s_resp = "Bahia"
		case "CE": s_resp = "Cear�"
		case "DF": s_resp = "Distrito Federal"
		case "ES": s_resp = "Esp�rito Santo"
		case "GO": s_resp = "Goi�s"
		case "MA": s_resp = "Maranh�o"
		case "MG": s_resp = "Minas Gerais"
		case "MS": s_resp = "Mato Grosso do Sul"
		case "MT": s_resp = "Mato Grosso"
		case "PA": s_resp = "Par�"
		case "PB": s_resp = "Para�ba"
		case "PE": s_resp = "Pernambuco"
		case "PI": s_resp = "Piau�"
		case "PR": s_resp = "Paran�"
		case "RJ": s_resp = "Rio de Janeiro"
		case "RN": s_resp = "Rio Grande do Norte"
		case "RO": s_resp = "Rond�nia"
		case "RR": s_resp = "Roraima"
		case "RS": s_resp = "Rio Grande do Sul"
		case "SC": s_resp = "Santa Catarina"
		case "SE": s_resp = "Sergipe"
		case "SP": s_resp = "S�o Paulo"
		case "TO": s_resp = "Tocantins"
		case else: s_resp = ""
	end select
	
	UF_descricao = s_resp
end function



' ------------------------------------------------------------------------
'   TELEFONE_FORMATA
' 
Function telefone_formata(byval telefone)
dim i
dim s_tel
	s_tel = "" & telefone
	s_tel = retorna_so_digitos(s_tel)
	
	telefone_formata = s_tel
	
	if ((s_tel="") Or (len(s_tel)>9) Or (Not telefone_ok(s_tel))) then exit function
		 
	i=len(s_tel)-4
	s_tel = mid(s_tel, 1, i) & "-" & mid(s_tel, i+1, len(s_tel))
	
	telefone_formata = s_tel
	
End Function



' ------------------------------------------------------------------------
'   CEP_FORMATA
' 
Function cep_formata(byval cep)
dim s_cep
	s_cep = "" & cep
	s_cep = retorna_so_digitos(s_cep)
	
	cep_formata = s_cep
	
	if ((s_cep="") Or (len(s_cep)=5) Or (Not cep_ok(s_cep))) then exit function
	
	s_cep = mid(s_cep,1,5) & "-" & mid(s_cep,6,3)
	
	cep_formata = s_cep
	
End Function



' ------------------------------------------------------------------------
'   CNPJ OK?
' 
Function cnpj_ok(byval cnpj)
Const p1 = "543298765432"
Const p2 = "6543298765432"
Dim d
Dim i
Dim tudo_igual
	
	cnpj_ok = False

	cnpj = "" & cnpj	
    cnpj = retorna_so_digitos(cnpj)
    
    If Trim(cnpj) = "" Then 
		cnpj_ok = True
		Exit Function
		end if
		
    If Len(cnpj) <> 14 Then Exit Function

	tudo_igual=True
	for i = 1 to (len(cnpj)-1)
		if mid(cnpj,i,1) <> mid(cnpj,i+1,1) then
			tudo_igual=False
			exit for
			end if
		next

	if tudo_igual then Exit Function

'�  VERIFICA O PRIMEIRO CHECK DIGIT
    d = 0
    For i = 1 To 12
        d = d + Clng(Mid(p1, i, 1)) * Clng(Mid(cnpj, i, 1))
        Next

    d = 11 - (d Mod 11)
    If d > 9 Then d = 0
    If d <> Clng(Mid(cnpj, 13, 1)) Then Exit Function

'�  VERIFICA O SEGUNDO CHECK DIGIT
    d = 0
    For i = 1 To 13
        d = d + Clng(Mid(p2, i, 1)) * Clng(Mid(cnpj, i, 1))
        Next

    d = 11 - (d Mod 11)
    If d > 9 Then d = 0
    If d <> Clng(Mid(cnpj, 14, 1)) Then Exit Function

    cnpj_ok = True
         
End Function



' ------------------------------------------------------------------------
'   CPF OK?
' 
Function cpf_ok(byval cpf)
Dim d
Dim i
Dim tudo_igual

	cpf_ok = False
	
    cpf = "" & cpf
    cpf = retorna_so_digitos(cpf)
    
    If Trim(cpf) = "" Then 
		cpf_ok = True
		Exit Function
		end if

    If Len(cpf) <> 11 Then Exit Function

	tudo_igual=True
	for i = 1 to (len(cpf)-1)
		if mid(cpf,i,1) <> mid(cpf,i+1,1) then
			tudo_igual=False
			exit for
			end if
		next

	if tudo_igual then Exit Function
	
	
'�  VERIFICA O PRIMEIRO CHECK DIGIT
    d = 0
    For i = 1 To 9
        d = d + (11 - i) * Clng(Mid(cpf, i, 1))
        Next

    d = 11 - (d Mod 11)
    If d > 9 Then d = 0
    If d <> Clng(Mid(cpf, 10, 1)) Then Exit Function

'�  VERIFICA O SEGUNDO CHECK DIGIT
    d = 0
    For i = 2 To 10
        d = d + (12 - i) * Clng(Mid(cpf, i, 1))
        Next
    
    d = 11 - (d Mod 11)
    If d > 9 Then d = 0
    If d <> Clng(Mid(cpf, 11, 1)) Then Exit Function

    cpf_ok = True


End Function



' ------------------------------------------------------------------------
'   CNPJ_CPF_FORMATA
' 
Function cnpj_cpf_formata(ByVal cnpj_cpf)
Dim s
Dim s_resp

    cnpj_cpf = "" & cnpj_cpf
    s = retorna_so_digitos(cnpj_cpf)
    
  '�CPF
    If Len(s) = 11 Then
        s_resp = Mid(s, 1, 3) & "." & Mid(s, 4, 3) & "." & Mid(s, 7, 3) & "/" & Mid(s, 10, 2)
  '�CNPJ
    ElseIf Len(s) = 14 Then
        s_resp = Mid(s, 1, 2) & "." & Mid(s, 3, 3) & "." & Mid(s, 6, 3) & "/" & Mid(s, 9, 4) & "-" & Mid(s, 13, 2)
  '�DESCONHECIDO
    Else
        s_resp = cnpj_cpf
        End If
    
    cnpj_cpf_formata = s_resp
    
End Function



' ------------------------------------------------------------------------
'   CNPJ_FORMATA
' 
function cnpj_formata(byval cnpj)
	
	cnpj_formata = cnpj_cpf_formata(cnpj)
	
end function



' ------------------------------------------------------------------------
'   CPF_FORMATA
' 
function cpf_formata(byval cpf)
	
	cpf_formata = cnpj_cpf_formata(cpf)
	
end function



' ------------------------------------------------------------------------
'   SEXO OK?
' 
Function sexo_ok(byval sexo)
dim s
	sexo_ok = False
	s = Ucase(Trim("" & sexo))
	if (s="M") Or (s="F") then sexo_ok = True
end function



' ------------------------------------------------------------------------
'   CNPJ / CPF OK?
' 
Function cnpj_cpf_ok(byval cnpj_cpf)
dim s_cnpj_cpf
	
	cnpj_cpf_ok = False
	
	s_cnpj_cpf = retorna_so_digitos(cnpj_cpf)
	
	if s_cnpj_cpf = "" then
		cnpj_cpf_ok = True
		exit function
		end if
	
	if len(s_cnpj_cpf)=11 then
		if Not cpf_ok(s_cnpj_cpf) then exit function
	elseif len(s_cnpj_cpf)=14 then
		if Not cnpj_ok(s_cnpj_cpf) then exit function
	else
		exit function
		end if
	
	cnpj_cpf_ok = True
		
end function



' ------------------------------------------------------------------------
'   DECODIFICA_DATA
'   Desmembra a data e retorna os respectivos valores para dia, m�s e ano.
function decodifica_data(byval dt, byref dia, byref mes, byref ano)
	
	decodifica_data = False
	
	dia=""
	mes=""
	ano=""
	
	if IsNull(dt) then Exit Function
	if Not IsDate(dt) then Exit Function
	
	dt=CDate(dt)
		
'   DIA
	dia = Cstr(Day(dt))
	If Len(dia) = 1 Then dia = "0" & dia
	
'   M�S
	mes = Cstr(Month(dt))
	If Len(mes) = 1 Then mes = "0" & mes
	
'   ANO
	ano = Cstr(Year(dt))
	If Len(ano) = 2 Then 
	    If CInt(ano) > 90 then 
			ano = "19" & ano
		else
			ano = "20" & ano
			End if
		End if
	
	decodifica_data = True
	
end function



' --------------------------------------------------------------------------
'   DECODIFICA_HORA
'   Desmembra a data e retorna os respectivos valores para hora, min e seg.
function decodifica_hora(byval dt, byref hora, byref min, byref seg) 

	decodifica_hora = False
	
	hora=""
	min=""
	seg=""
	
	if IsNull(dt) then Exit Function
	if Not IsDate(dt) then Exit Function
	
	dt=CDate(dt)
		
'   HORA
	hora = Cstr(Hour(dt))
	If Len(hora) = 1 Then hora = "0" & hora
	
'   MINUTO
	min = Cstr(Minute(dt))
	If Len(min) = 1 Then min = "0" & min
	
'   SEGUNDO
	seg = Cstr(Second(dt))
	If Len(seg) = 1 Then seg = "0" & seg
	
	decodifica_hora = True
			
End Function



' ------------------------------------------------------------------------
'   FORMATA_DATA
'   Formata somente a data: DD/MM/YYYY
function formata_data(byval dt)
dim dia
dim mes 
dim ano
	formata_data = ""
	if Not decodifica_data(dt, dia, mes, ano) then exit function
	formata_data = dia & "/" & mes & "/" & ano 
end function



' ------------------------------------------------------------------------
'   FORMATA_DATA_MMDDYYYY
'   Formata somente a data: MM/DD/YYYY
function formata_data_mmddyyyy(byval dt)
dim dia
dim mes 
dim ano
	formata_data_mmddyyyy = ""
	if Not decodifica_data(dt, dia, mes, ano) then exit function
	formata_data_mmddyyyy = mes & "/" & dia & "/" & ano
end function



' ------------------------------------------------------------------------
'   FORMATA_HORA
'   Formata somente a hora: HH:NN:SS
function formata_hora(byval dt)
dim hora
dim min
dim seg
	formata_hora=""
	if Not decodifica_hora(dt, hora, min, seg) then exit function
	formata_hora=hora & ":" & min & ":" & seg
end function



' ------------------------------------------------------------------------
'   FORMATA_HORA_HHMM
'   Formata somente a hora: HH:NN
function formata_hora_hhmm(byval dt)
dim hora
dim min
dim seg
	formata_hora_hhmm=""
	if Not decodifica_hora(dt, hora, min, seg) then exit function
	formata_hora_hhmm=hora & ":" & min
end function



' ------------------------------------------------------------------------
'   FORMATA_DATA_YYYYMMDD
'   Formata somente a data: YYYYMMDD
function formata_data_yyyymmdd(byval dt)
dim s_data
	s_data = formata_data(dt)
	s_data = mid(s_data,7,4) & mid(s_data,4,2) & mid(s_data,1,2)
	formata_data_yyyymmdd = s_data
end function



' ------------------------------------------------------------------------
'   FORMATA_DATA_COM_SEPARADOR_YYYYMMDD
'   Formata a data usando o separador especificado: YYYY?MM?DD
function formata_data_com_separador_yyyymmdd(byval dt, byval separador)
dim s_data
	s_data = formata_data(dt)
	s_data = mid(s_data,7,4) & separador & mid(s_data,4,2) & separador & mid(s_data,1,2)
	formata_data_com_separador_yyyymmdd = s_data
end function



' ------------------------------------------------------------------------
'   FORMATA_HORA_HHNNSS
'   Formata somente a hora: HHNNSS, sem separadores
function formata_hora_hhnnss(byval dt)
	formata_hora_hhnnss=retorna_so_digitos(formata_hora(dt))
end function



' ------------------------------------------------------------------------
'   FORMATA HHNNSS PARA HH:NN:SS
'   Formata a sequ�ncia hhnnss para hh:nn:ss
function formata_hhnnss_para_hh_nn_ss(byval hhnnss)
dim s
dim s_hora
dim s_min
dim s_seg
	formata_hhnnss_para_hh_nn_ss  = ""
	hhnnss = Trim("" & hhnnss)
	s = retorna_so_digitos(hhnnss)
	if s <> "" then
		s_hora = mid(hhnnss,1,2)
		s_min = mid(hhnnss,3,2)
		s_seg = mid(hhnnss,5,2)
		s = s_hora & ":" & s_min
		if s_seg <> "" then s = s & ":" & s_seg
		formata_hhnnss_para_hh_nn_ss = s
		end if
end function



' ------------------------------------------------------------------------
'   FORMATA HHNNSS PARA HH:NN
'   Formata a sequ�ncia hhnnss para hh:nn
function formata_hhnnss_para_hh_nn(byval hhnnss)
dim s
dim s_hora
dim s_min
dim s_seg
	formata_hhnnss_para_hh_nn  = ""
	hhnnss = Trim("" & hhnnss)
	s = retorna_so_digitos(hhnnss)
	if s <> "" then
		s_hora = mid(hhnnss,1,2)
		s_min = mid(hhnnss,3,2)
		s_seg = mid(hhnnss,5,2)
		s = s_hora & ":" & s_min
		formata_hhnnss_para_hh_nn = s
		end if
end function



' ------------------------------------------------------------------------
'   FORMATA_DATA_HORA
'   Formata a data e hora: DD/MM/YYYY HH:NN:SS
function formata_data_hora(byval dt)
dim s
dim dia
dim mes 
dim ano
dim hora
dim min
dim seg
	formata_data_hora=""
	if Not decodifica_data(dt, dia, mes, ano) then exit function
	s=dia & "/" & mes & "/" & ano 
	if decodifica_hora(dt, hora, min, seg) then s=s & " " & hora & ":" & min & ":" & seg
	formata_data_hora=s
end function



' ------------------------------------------------------------------------
'   FORMATA_DATA_HORA_SEM_SEG
'   Formata a data e hora: DD/MM/YYYY HH:NN
function formata_data_hora_sem_seg(byval dt)
dim s
dim dia
dim mes 
dim ano
dim hora
dim min
dim seg
	formata_data_hora_sem_seg=""
	if Not decodifica_data(dt, dia, mes, ano) then exit function
	s=dia & "/" & mes & "/" & ano 
	if decodifica_hora(dt, hora, min, seg) then s=s & " " & hora & ":" & min
	formata_data_hora_sem_seg=s
end function



' ------------------------------------------------------------------------
'   FORMATA_DATA_E_TALVEZ_HORA
'   Formata a data e hora (se houver hora): DD/MM/YYYY HH:NN:SS
'	Sen�o ser� apenas a data: DD/MM/YYYY
function formata_data_e_talvez_hora(byval dt)
dim s
dim dia
dim mes 
dim ano
dim hora
dim min
dim seg
	formata_data_e_talvez_hora=""
	if Not decodifica_data(dt, dia, mes, ano) then exit function
	s=dia & "/" & mes & "/" & ano 
	if decodifica_hora(dt, hora, min, seg) then 
		if (hora <> "00") Or (min <> "00") Or (seg <> "00") then
			s=s & " " & hora & ":" & min & ":" & seg
			end if
		end if
	formata_data_e_talvez_hora=s
end function



' ------------------------------------------------------------------------
'	FORMATA_DATA_E_TALVEZ_HORA_HHMM
'	Formata a data e hora (se houver hora): DD/MM/YYYY HH:NN
'	Sen�o ser� apenas a data: DD/MM/YYYY
'	Lembrando que mesmo que a informa��o referente aos segundos n�o
'	seja exibida, o fato desse campo ser diferente de zero significa
'	que h� informa��o sobre o hor�rio armazenado.
function formata_data_e_talvez_hora_hhmm(byval dt)
dim s
dim dia
dim mes 
dim ano
dim hora
dim min
dim seg
	formata_data_e_talvez_hora_hhmm=""
	if Not decodifica_data(dt, dia, mes, ano) then exit function
	s=dia & "/" & mes & "/" & ano 
	if decodifica_hora(dt, hora, min, seg) then 
		if (hora <> "00") Or (min <> "00") Or (seg <> "00") then
			s=s & " " & hora & ":" & min
			end if
		end if
	formata_data_e_talvez_hora_hhmm=s
end function



' ================================================================================
'      FORMATA DURACAO HMS
' --------------------------------------------------------------------------------
'      Formata o tempo decorrido no formato '0h0m0s'.
'      Se n�o exceder 60m, ent�o o formato ser� '0m0s'.
'      Se n�o exceder 60s, ent�o o formato ser� '0s'.
' --------------------------------------------------------------------------------
function formata_duracao_hms(dt)
dim hh, mm, ss
dim strResp
	hh = Hour(dt)
	mm = Minute(dt)
	ss = Second(dt)
	strResp = ""
	
	if strResp = "" then
		if hh > 0 then strResp = Cstr(hh) & "h"
	else
		strResp = strResp & normaliza_a_esq(Cstr(hh), 2) & "h"
		end if
		
	if strResp = "" then
		if mm > 0 then strResp = Cstr(mm) & "m"
	else
		strResp = strResp & normaliza_a_esq(Cstr(mm), 2) & "m"
		end if
		
	if strResp = "" then
		strResp = Cstr(ss) & "s"
	else
		strResp = strResp & normaliza_a_esq(Cstr(ss), 2) & "s"
		end if
		
	formata_duracao_hms = strResp
end function



' ------------------------------------------------------------------------
'   BD_OBTEM_MES
'   Retorna o m�s em ingl�s por extenso ou apenas c/ os 3 primeiros caracteres.
function bd_obtem_mes(ByVal i, byval por_extenso)
dim s
dim j
    If IsNumeric(i) Then j = CInt(i) Else j = 0

    if is_sgbd_access then
		select case j
		    case 1: s = "JANEIRO"
		    case 2: s = "FEVEREIRO"
		    case 3: s = "MAR�O"
		    case 4: s = "ABRIL"
		    case 5: s = "MAIO"
		    case 6: s = "JUNHO"
		    case 7: s = "JULHO"
		    case 8: s = "AGOSTO"
		    case 9: s = "SETEMBRO"
		    case 10: s = "OUTUBRO"
		    case 11: s = "NOVEMBRO"
		    case 12: s = "DEZEMBRO"
		    case else: s = ""
		    end select
	else    
		select case j
		    case 1: s = "JANUARY"
		    case 2: s = "FEBRUARY"
		    case 3: s = "MARCH"
		    case 4: s = "APRIL"
		    case 5: s = "MAY"
		    case 6: s = "JUNE"
		    case 7: s = "JULY"
		    case 8: s = "AUGUST"
		    case 9: s = "SEPTEMBER"
		    case 10: s = "OCTOBER"
		    case 11: s = "NOVEMBER"
		    case 12: s = "DECEMBER"
		    case else: s = ""
		    end select
	end if

    If Not por_extenso Then s = Mid(s, 1, 3)
    bd_obtem_mes = s
end function



' ------------------------------------------------------------------------
'   BD_FORMATA_DATA
'   Monta a express�o para criar um tipo datetime do SQL Server com a data (sem hora) especificada
function bd_formata_data(byval dt)
dim dia, mes, ano
dim strDia, strMes, strAno
dim strData
	if is_sgbd_access then
		bd_formata_data = "NULL"
		if Not decodifica_data(dt, dia, mes, ano) then exit function
		bd_formata_data = "CDate('" & bd_obtem_mes(mes, False) & " " & dia & " " & ano & "')"
	else
		bd_formata_data = "NULL"
		if Not decodifica_data(dt, dia, mes, ano) then exit function
		strDia = Cstr(dia)
		if Len(strDia) = 1 then strDia = "0" & strDia
		strMes = Cstr(mes)
		if Len(strMes) = 1 then strMes = "0" & strMes
		strAno = Cstr(ano)
		strData = strAno & "-" & strMes & "-" & strDia
		bd_formata_data = "Convert(datetime, '" & strData & "', 120)"
		end if
end function



' ------------------------------------------------------------------------
'   BD_MONTA_DATA
'   Monta a express�o para criar um tipo datetime do SQL Server com a data (sem hora) especificada
function bd_monta_data(byval dt)
	bd_monta_data = bd_formata_data(dt)
end function



' ------------------------------------------------------------------------
'   BD_FORMATA_DATA_HORA
'   Monta a express�o para criar um tipo datetime do SQL Server com a data/hora especificada
function bd_formata_data_hora(byval dt)
dim s
dim dia, mes, ano
dim hora, min, seg
dim strDia, strMes, strAno
dim strHora, strMin, strSeg
dim strDataHora

	if is_sgbd_access then
		bd_formata_data_hora = "NULL"
		if Not decodifica_data(dt, dia, mes, ano) then exit function
		s = bd_obtem_mes(mes, False) & " " & dia & " " & ano
		if decodifica_hora(dt, hora, min, seg) then s=s & " " & hora & ":" & min & ":" & seg
		bd_formata_data_hora = "CDate('" & s & "')"
	else
		bd_formata_data_hora = "NULL"
		if Not decodifica_data(dt, dia, mes, ano) then exit function
		if Not decodifica_hora(dt, hora, min, seg) then
			hora = 0
			min = 0
			seg = 0
			end if
		strDia = Cstr(dia)
		if Len(strDia) = 1 then strDia = "0" & strDia
		strMes = Cstr(mes)
		if Len(strMes) = 1 then strMes = "0" & strMes
		strAno = Cstr(ano)
		strHora = Cstr(hora)
		if Len(strHora) = 1 then strHora = "0" & strHora
		strMin = Cstr(min)
		if Len(strMin) = 1 then strMin = "0" & strMin
		strSeg = Cstr(seg)
		if Len(strSeg) = 1 then strSeg = "0" & strSeg
		strDataHora = strAno & "-" & strMes & "-" & strDia & " " & strHora & ":" & strMin & ":" & strSeg
		bd_formata_data_hora = "Convert(datetime, '" & strDataHora & "', 120)"
		end if
end function




' ------------------------------------------------------------------------
'   BD_MONTA_DATA_HORA
'   Monta a express�o para criar um tipo datetime do SQL Server com a data/hora especificada
function bd_monta_data_hora(byval dt)
	bd_monta_data_hora = bd_formata_data_hora(dt)
end function



' ------------------------------------------------------------------------
'   STR TO DATE
'   Converte a data do tipo texto no formato DD/MM/YYYY para o tipo date.
function StrToDate(byval data) 
const DATA_PORTUGUES = "12/01/2000"
const DATA_INGLES = "01/12/2000"
dim strSeparador
dim vAux
dim dia
dim mes
dim ano
dim eh_portugues
dim i
dim c
dim strDataOriginal
dim dtResp
		
	StrToDate = 0
	
	if IsNull(data) then exit function
	if Instr(1, data, " ") then data = Trim(data)

'   NORMALIZA TAMANHO
	strSeparador = ""
	for i=1 to Len(data)
		c = Mid(data,i,1)
		if Not IsNumeric(c) then
			strSeparador = c
			exit for
			end if
		next

	strDataOriginal = data
	if strSeparador <> "" then
		vAux = Split(data, strSeparador, -1)
		for i=Lbound(vAux) to Ubound(vAux)
			if Len(vAux(i)) = 1 then vAux(i) = "0" & vAux(i)
			next
		data = Join(vAux, strSeparador)
		strDataOriginal = Join(vAux, "/")
		end if

'   TAMANHO CORRETO?
	If (Len(data) < 8) Or (Len(data) > 10) then Exit Function
		
'   PORTUGUES / INGLES?
	dia = Day(DATA_PORTUGUES)
	eh_portugues = (dia = 12)
	If Not eh_portugues Then 
		dia = Day(DATA_INGLES)
		If (dia <> 12) Then Exit Function
		end if
		
'   INVERTE DIA/MES/ANO PARA MES/DIA/ANO (FORMATO INGLES)
	if Not eh_portugues then 
		dia = ""
		mes = ""
		ano = ""
		for i=1 to Len(data)
			c = Mid(data,i,1)
			if IsNumeric(c) then 
				if Len(dia) < 2 then
					dia = dia & c
				elseif Len(mes) < 2 then
					mes = mes & c
				elseif Len(ano) < 4 then
					ano = ano & c
					end if
			else
				if dia <> "" then do while Len(dia) < 2: dia = "0" & dia: loop
				if mes <> "" then do while Len(mes) < 2: mes = "0" & mes: loop
				end if
			next

		if dia = "" then Exit Function
		if mes = "" then Exit Function
		if Len(ano) <> 4 then Exit Function

		data = mes & "/" & dia & "/" & ano
		end if

'   ANO OK?
	ano = retorna_so_digitos(Mid(data,7,4))
	if Len(ano) <> 4 then Exit Function
    If CInt(ano) < 1900 then Exit Function
	
'   DATA OK?
	If Not IsDate(data) then Exit Function
	
	dtResp = CDate(data)
	if retorna_so_digitos(strDataOriginal) <> retorna_so_digitos(formata_data(dtResp)) then exit function
	
	StrToDate = dtResp
	
end function



' ------------------------------------------------------------------------
'   STR TO TIME
'   Converte a hora do tipo texto no formato HH:NN:SS para o tipo date.
function StrToTime(byval data) 
dim hora
dim min
dim seg
dim i
dim c
		
	StrToTime = 0
	
	if Instr(1,data,":")=0 then Exit Function
	If (Len(data) < 3) Or (Len(data) > 8) then Exit Function
		
	hora = ""
	min = ""
	seg = ""
	for i=1 to Len(data)
		c = Mid(data,i,1)
		if IsNumeric(c) then 
			if Len(hora) < 2 then
				hora = hora & c
			elseif Len(min) < 2 then
				min = min & c
			elseif Len(seg) < 2 then
				seg = seg & c
				end if
		else
			if hora <> "" then do while Len(hora) < 2: hora = "0" & hora: loop
			if min <> "" then do while Len(min) < 2: min = "0" & min: loop
			end if
		next

	do while Len(seg) < 2: seg = "0" & seg: loop

	if hora = "" then Exit Function
	if min = "" then Exit Function
	if seg = "" then Exit Function

	if CInt(hora) > 24 then Exit Function
	if CInt(min) > 60 then Exit Function
	if CInt(seg) > 60 then Exit Function
	
	data = hora & ":" & min & ":" & seg
	
'   DATA OK?
	If Not IsDate(data) then Exit Function
	
	StrToTime = CDate(data)
	
end function



' ------------------------------------------------------------------------
'   STR TO DATETIME
'   Converte a data do tipo texto no formato
'	DD/MM/YYYY HH:NN:SS para o tipo date.
function StrToDateTime(byval data) 
dim s_data
dim s_hora
dim dt_data
dim dt_hora
dim dt

	StrToDateTime = 0

'   TEM DATA E HORA?
	s_data = ""
	s_hora = ""
	if InStr(1, data, " ") <> 0 then
		s_data=Trim(Mid(data, 1, InStr(1, data, " ")))
		s_hora=Trim(Mid(data, InStr(1, data, " ")+1, Len(data)))
	else
	'   SOMENTE HORA OU SOMENTE DATA
		if InStr(1, data, ":") <> 0 then
			s_hora = data
		else
			s_data = data
			end if
		end if
		
	dt = 0
	dt_data = 0
	dt_hora = 0
	
	if s_data <> "" then dt_data = StrToDate(s_data)
	if s_hora <> "" then dt_hora = StrToTime(s_hora)
	
	if dt_data <> 0 then dt = dt_data
	if dt_hora <> 0 then dt = dt + dt_hora
	
	StrToDateTime = dt
	
end function



' ------------------------------------------------------------------------
'	Converte a data de String para Date
'		Entrada: string no formato MM/DD/YYYY HH:MM:SS AM/PM
'		Sa�da: data com tipo de dados Date
function converte_para_datetime_from_mmddyyyy_hhmmss_am_pm(byval data_hora)
dim v, intIndex
dim strData, strHorario, strAmPm
dim strDia, strMes, strAno
dim strHora, strMin, strSeg
dim strDdMmYyyyHhMmSs
	converte_para_datetime_from_mmddyyyy_hhmmss_am_pm = 0
	data_hora = Trim("" & data_hora)
	
	if data_hora = "" then exit function
	
	strData = ""
	strHorario = ""
	strAmPm = ""
	
	v = Split(data_hora, " ")
	intIndex = LBound(v)
	if intIndex <= UBound(v) then strData = Trim(v(intIndex))
	intIndex = intIndex+1
	if intIndex <= UBound(v) then strHorario = Trim(v(intIndex))
	intIndex = intIndex+1
	if intIndex <= UBound(v) then strAmPm = Trim(v(intIndex))
	
	strDia = ""
	strMes = ""
	strAno = ""
	strHora = "00"
	strMin = "00"
	strSeg = "00"
	
	if strData <> "" then
	'	DATA EST� NO FORMATO MM/DD/YYYY HH:MM:SS AM/PM
		v = Split(strData, "/")
		intIndex = LBound(v)
		if intIndex <= UBound(v) then strMes = Trim(v(intIndex))
		intIndex = intIndex+1
		if intIndex <= UBound(v) then strDia = Trim(v(intIndex))
		intIndex = intIndex+1
		if intIndex <= UBound(v) then strAno = Trim(v(intIndex))
		end if
	
	if strHorario <> "" then
	'	DATA EST� NO FORMATO MM/DD/YYYY HH:MM:SS AM/PM
		v = Split(strHorario, ":")
		intIndex = LBound(v)
		if intIndex <= UBound(v) then strHora = Trim(v(intIndex))
		intIndex = intIndex+1
		if intIndex <= UBound(v) then strMin = Trim(v(intIndex))
		intIndex = intIndex+1
		if intIndex <= UBound(v) then strSeg = Trim(v(intIndex))
		end if
	
	if Ucase(strAmPm) = "AM" then
		if strHora = "12" then
		''	12:00:00 AM -> 00:00:00 meia-noite
			strHora = "00"
		else
		'	NOP
		'	01:00:00 AM -> 01:00:00
		'	11:00:00 AM -> 11:00:00
			end if
		end if

	if Ucase(strAmPm) = "PM" then
		if strHora = "12" then
		'	NOP: 12:00:00 PM -> 12:00:00 (meio-dia)
		else
		'	01:00:00 PM -> 13:00
		'	11:00:00 PM -> 23:00
			strHora = Cstr(CInt(strHora) + 12)
			if Len(strHora) = 1 then strHora = "0" & strHora
			end if
		end if
	
	strDdMmYyyyHhMmSs = strDia & "/" & strMes & "/" & strAno & " " & strHora & ":" & strMin & ":" & strSeg
	converte_para_datetime_from_mmddyyyy_hhmmss_am_pm = StrToDateTime(strDdMmYyyyHhMmSs)
end function



' ------------------------------------------------------------------------
'	Converte a data de String para Date
'		Entrada: string no formato YYYY-MM-DD HH:MM:SS (formato 24h)
'		Sa�da: data com tipo de dados Date
function converte_para_datetime_from_yyyymmdd_hhmmss(byval data_hora)
dim v, intIndex
dim strData, strHorario
dim strDia, strMes, strAno
dim strHora, strMin, strSeg
dim strDdMmYyyyHhMmSs
	converte_para_datetime_from_yyyymmdd_hhmmss = 0
	data_hora = Trim("" & data_hora)
	
	if data_hora = "" then exit function
	
	strData = ""
	strHorario = ""
	
	v = Split(data_hora, " ")
	intIndex = LBound(v)
	if intIndex <= UBound(v) then strData = Trim(v(intIndex))
	intIndex = intIndex+1
	if intIndex <= UBound(v) then strHorario = Trim(v(intIndex))
	intIndex = intIndex+1
	if intIndex <= UBound(v) then strAmPm = Trim(v(intIndex))
	
	strDia = ""
	strMes = ""
	strAno = ""
	strHora = "00"
	strMin = "00"
	strSeg = "00"
	
	if Instr(strData, "/") <> 0 then strData = Replace(strData, "/", "-")

	if strData <> "" then
	'	DATA EST� NO FORMATO YYYY-MM-DD HH:MM:SS (formato 24h)
		v = Split(strData, "-")
		intIndex = LBound(v)
		if intIndex <= UBound(v) then strAno = Trim(v(intIndex))
		intIndex = intIndex+1
		if intIndex <= UBound(v) then strMes = Trim(v(intIndex))
		intIndex = intIndex+1
		if intIndex <= UBound(v) then strDia = Trim(v(intIndex))
		end if
	
	if strHorario <> "" then
	'	DATA EST� NO FORMATO YYYY-MM-DD HH:MM:SS (formato 24h)
		v = Split(strHorario, ":")
		intIndex = LBound(v)
		if intIndex <= UBound(v) then strHora = Trim(v(intIndex))
		intIndex = intIndex+1
		if intIndex <= UBound(v) then strMin = Trim(v(intIndex))
		intIndex = intIndex+1
		if intIndex <= UBound(v) then strSeg = Trim(v(intIndex))
		end if
	
	strDdMmYyyyHhMmSs = strDia & "/" & strMes & "/" & strAno & " " & strHora & ":" & strMin & ":" & strSeg
	converte_para_datetime_from_yyyymmdd_hhmmss = StrToDateTime(strDdMmYyyyHhMmSs)
end function




' ------------------------------------------------------------------------
'   TEXTO_ADD_CR
function texto_add_cr(byval texto)
	if texto <> "" then
		texto_add_cr=texto & chr(13)
	else
		texto_add_cr=texto
		end if
end function



' ------------------------------------------------------------------------
'   TEXTO_ADD_BR
function texto_add_br(byval texto)
	if texto <> "" then
		texto_add_br=texto & "<br>"
	else
		texto_add_br=texto
		end if
end function



' ------------------------------------------------------------------------
'   RETORNA_SEPARADOR_DECIMAL
function retorna_separador_decimal(byval numero)
dim i
dim c
dim s_num
dim s_resp
dim n_ponto
dim n_virg
dim s_ult_sep
dim n_digitos_finais
dim n_digitos_iniciais

	n_digitos_finais=0
	n_digitos_iniciais=0
	n_ponto=0
	n_virg=0
	s_ult_sep=""
	s_num = Trim("" & numero)
	for i=Len(s_num) to 1 step -1
		c=mid(s_num, i, 1)
		if (c=".") then
			n_ponto=n_ponto+1
			if (s_ult_sep="") then s_ult_sep=c
		elseif (c=",") then
			n_virg=n_virg+1
			if (s_ult_sep="") then s_ult_sep=c
			end if
		if IsNumeric(c) And (n_ponto=0) And (n_virg=0) then n_digitos_finais=n_digitos_finais+1
		if IsNumeric(c) And (s_ult_sep<>"") then n_digitos_iniciais=n_digitos_iniciais+1
		next
		
'	DEFAULT
	s_resp = ","
	if (s_ult_sep=".") then
		if (n_ponto=1) And (n_virg=0) And (n_digitos_finais=3) then
			if n_digitos_iniciais > 3 then
			'	CONSIDERA 1234.567 COMO MIL DUZENTOS E TRINTA E QUATRO E QUINHENTOS E SESSENTA E SETE MIL�SIMOS
				s_resp="."
			else
			'	NOP: CONSIDERA 123.456 COMO CENTO E VINTE E TR�S MIL E QUATROCENTOS E CINQUENTA E SEIS,
				end if
		elseif (n_ponto=1) then 
			s_resp="."
			end if
	elseif (s_ult_sep=",") then
		if (n_virg > 1) And (n_ponto=0) then s_resp="."
		end if
		
	retorna_separador_decimal=s_resp
end function



' ------------------------------------------------------------------------
'   RETORNA SEPARADOR DECIMAL SISTEMA
function retorna_separador_decimal_sistema
dim i
dim c
dim s
dim s_sep_sys
	retorna_separador_decimal_sistema = ""
	s_sep_sys = ""
	s = cstr(.5)
	for i = 1 to len(s)
		c = mid(s,i,1)
		if Not IsNumeric(c) then
			s_sep_sys = c
			exit for
			end if
		next
	if s_sep_sys = "" then exit function
	retorna_separador_decimal_sistema = s_sep_sys
end function



' ------------------------------------------------------------------------
'   CONVERTE_NUMERO
function converte_numero(byval numero)
dim i
dim s
dim c
dim s_sep
dim s_sep_sys
dim s_valor

	converte_numero = 0

	if (vartype(numero) <> vbString) And IsNumeric(numero) then
		converte_numero = numero * 1
		exit function
		end if

	numero=Trim("" & numero)
	if numero="" then exit function
	s_sep=retorna_separador_decimal(numero)
	s_sep_sys = retorna_separador_decimal_sistema()
	if s_sep_sys="" then exit function
	numero=substitui_caracteres(numero,s_sep,"V")

	s=""
	for i=1 to Len(numero)
		c=mid(numero, i, 1)
		if (Not IsNumeric(c)) And (c<>"-") And (c<>"V") then c=""
		s = s & c
		next
	
	s_valor=substitui_caracteres(s, "V", s_sep_sys)
	if Not IsNumeric(s_valor) then exit function
	
	converte_numero = s_valor * 1
end function



' ------------------------------------------------------------------------
'   JS FORMATA NUMERO
function js_formata_numero(byval numero)
dim s, s_sep_sys
	js_formata_numero=""
	if IsNull(numero) then exit function
	if Trim("" & numero)="" then exit function
	s_sep_sys = retorna_separador_decimal_sistema()
	if s_sep_sys = "" then exit function
	s = Cstr(converte_numero(numero))
	s = substitui_caracteres(s, s_sep_sys, "V")
	s = substitui_caracteres(s, ".", "")
	s = substitui_caracteres(s, ",", "")
	s = substitui_caracteres(s, "V", ".")
	js_formata_numero = s
end function



' ------------------------------------------------------------------------
'   BD FORMATA NUMERO
function bd_formata_numero(byval numero)
dim s, s_sep_sys
	bd_formata_numero=""
	if IsNull(numero) then exit function
	if Trim("" & numero)="" then exit function
	s_sep_sys = retorna_separador_decimal_sistema()
	if s_sep_sys = "" then exit function
	s = Cstr(converte_numero(numero))
	s = substitui_caracteres(s, s_sep_sys, "V")
	s = substitui_caracteres(s, ".", "")
	s = substitui_caracteres(s, ",", "")
	s = substitui_caracteres(s, "V", ".")
	bd_formata_numero = s
end function



' ------------------------------------------------------------------------
'   bd_formata_moeda
function bd_formata_moeda(ByVal numero)
dim strSeparadorDecimal
dim strValorFormatado
dim i
dim c
dim s
	strSeparadorDecimal = ""
	s = CStr(0.5)
	For i = Len(s) To 1 Step -1
		c = Mid(s, i, 1)
		If Not IsNumeric(c) Then
			strSeparadorDecimal = c
			Exit For
			End If
		Next

	If strSeparadorDecimal = "" Then strSeparadorDecimal = ","
	
'	FormatNumber(Expression, NumDigitsAfterDecimal, IncludeLeadingDigit, UseParensForNegativeNumbers, GroupDigits)
'	Lembrando que IncludeLeadingDigit indica se valores como .5 ser�o exibidos como .5 ou 0.5
'	A fun��o FormatCurrency sempre inclui o s�mbolo monet�rio.
	strValorFormatado = FormatNumber(numero, 2, -1, 0, 0)
	strValorFormatado = substitui_caracteres(strValorFormatado, strSeparadorDecimal, "V")
	strValorFormatado = substitui_caracteres(strValorFormatado, ".", "")
	strValorFormatado = substitui_caracteres(strValorFormatado, ",", "")
	strValorFormatado = substitui_caracteres(strValorFormatado, "V", ".")
	
	bd_formata_moeda = strValorFormatado
end function



' ------------------------------------------------------------------------
'   FORMATA_NUMERO
function formata_numero(byval valor, byval decimais)
dim s, c_decimal, c_grupo
	formata_numero=""
	if IsNull(valor) then exit function
	if Trim("" & valor)="" then exit function
	if Not IsNumeric(decimais) then decimais=0 else decimais=CLng(decimais)
	if decimais<0 then decimais=0
  ' FormatNumber(Expression, NumDigitsAfterDecimal, IncludeLeadingDigit, UseParensForNegativeNumbers, GroupDigits)
  ' Lembrando que IncludeLeadingDigit indica se valores como .5 ser�o exibidos como .5 ou 0.5
  ' A fun��o FormatCurrency sempre inclui o s�mbolo monet�rio.
	s = FormatNumber(1234.5, 2, -1, 0, -1)
	c_decimal = Left(Right(s, 3) ,1)
	c_grupo = Left(Right(s, 7), 1)
	s = FormatNumber(valor, decimais, -1, 0, -1)
	s = substitui_caracteres(s, c_decimal, "V")
	s = substitui_caracteres(s, c_grupo, ".")
	s = substitui_caracteres(s, "V", ",")
	formata_numero = s
end function



' ------------------------------------------------------------------------
'   FORMATA INTEIRO
function formata_inteiro(byval valor)
	formata_inteiro=formata_numero(valor, 0)
end function



' ------------------------------------------------------------------------
'   FORMATA_MOEDA
function formata_moeda(byval valor)
	formata_moeda=formata_numero(valor, 2)
end function



' ------------------------------------------------------------------------
'   FORMATA_PERC
function formata_perc(byval valor)
	formata_perc=formata_numero(valor, 2)
end function



' ------------------------------------------------------------------------
'   FORMATA_PERC1DEC
function formata_perc1dec(byval valor)
	formata_perc1dec=formata_numero(valor, 1)
end function



' ------------------------------------------------------------------------
'   FORMATA_PERC4DEC
function formata_perc4dec(byval valor)
	formata_perc4dec=formata_numero(valor, 4)
end function



' ------------------------------------------------------------------------
'   FORMATA_NUMERO6DEC
function formata_numero6dec(byval valor)
	formata_numero6dec=formata_numero(valor, 6)
end function



' ------------------------------------------------------------------------
'   FORMATA_PERC_RT
function formata_perc_RT(byval valor)
	formata_perc_RT=formata_numero(valor, 1)
end function



' ------------------------------------------------------------------------
'   FORMATA_PERC_DESC
function formata_perc_desc(byval valor)
	formata_perc_desc=formata_numero(valor, 1)
end function



' ------------------------------------------------------------------------
'   FORMATA_PERC_COMISSAO
function formata_perc_comissao(byval valor)
	formata_perc_comissao=formata_numero(valor, 1)
end function



' ------------------------------------------------------------------------
'   FORMATA_PERC_MARKUP
function formata_perc_markup(byval valor)
	formata_perc_markup=formata_numero(valor, 1)
end function



' ------------------------------------------------------------------------
'   FORMATA COEFICIENTE CUSTO FINANC FORNECEDOR
function formata_coeficiente_custo_financ_fornecedor(byval valor)
	formata_coeficiente_custo_financ_fornecedor=formata_numero(valor, MAX_DECIMAIS_COEFICIENTE_CUSTO_FINANCEIRO_FORNECEDOR)
end function



' ------------------------------------------------------------------------
'   NUMERACAO_ALTERA_BASE
function numeracao_altera_base(byval base_orig, byval base_dest, byval numero)
'	DEFAULT	
	numeracao_altera_base=numero
	if Not IsNumeric(base_orig) then exit function
	if Not IsNumeric(base_dest) then exit function
	if Not IsNumeric(numero) then exit function
	numeracao_altera_base = numero + (base_dest-base_orig)
end function



' ------------------------------------------------------------------------
'   RENUMERA_COM_BASE1
function renumera_com_base1(byval base_orig, byval numero)
	renumera_com_base1 = numeracao_altera_base(base_orig, 1, numero)
end function



' ------------------------------------------------------------------------
'   NORMALIZA_CODIGO
function normaliza_codigo(byref codigo, byval tamanho_default)
dim s
	normaliza_codigo = ""
	s=Trim("" & codigo)
	if s = "" then exit function
	do while len(s) < tamanho_default: s="0" & s: loop
	normaliza_codigo=s
end function



' ------------------------------------------------------------------------
'   NORMALIZA_A_ESQ
function normaliza_a_esq(byval numero, byval tamanho_default)
dim s
	normaliza_a_esq = ""
	s=Trim("" & numero)
	if s = "" then exit function
	do while len(s) < tamanho_default: s="0" & s: loop
	normaliza_a_esq=s
end function



' ------------------------------------------------------------------------
'   CONVERTE_SEG_TO_DEC
function converte_seg_to_dec(byval tempo_em_seg)
	converte_seg_to_dec=0
	if Not IsNumeric(tempo_em_seg) then exit function
	tempo_em_seg=CLng(tempo_em_seg)
	converte_seg_to_dec=tempo_em_seg/86400
end function



' ------------------------------------------------------------------------
'   CONVERTE_MIN_TO_DEC
function converte_min_to_dec(byval tempo_em_min)
	converte_min_to_dec=0
	if Not IsNumeric(tempo_em_min) then exit function
	tempo_em_min=CLng(tempo_em_min)
	converte_min_to_dec=converte_seg_to_dec(tempo_em_min*60)
end function



' ------------------------------------------------------------------------
'   CONVERTE_HORA_TO_DEC
function converte_hora_to_dec(byval tempo_em_hora)
	converte_hora_to_dec=0
	if Not IsNumeric(tempo_em_hora) then exit function
	tempo_em_hora=CLng(tempo_em_hora)
	converte_hora_to_dec=converte_seg_to_dec(tempo_em_hora*60*60)
end function



' ------------------------------------------------------------------------
'   FORMATA_TEXTO_LOG
function formata_texto_log(byval texto)
	texto=Trim("" & texto)
	if texto="" then texto=chr(34) & chr(34)
	formata_texto_log=texto
end function



' ------------------------------------------------------------------------
'   NORMALIZA_NUM_PEDIDO
function normaliza_num_pedido(byval id_pedido)
dim i
dim c
dim s
dim s_num
dim s_ano
dim s_filhote
	normaliza_num_pedido = ""
	id_pedido = Ucase(Trim("" & id_pedido))
	if id_pedido = "" then exit function
	s_num = ""
	for i=1 to len(id_pedido)
		if IsNumeric(mid(id_pedido,i,1)) then
			s_num=s_num & mid(id_pedido,i,1)
		else
			exit for
			end if
		next
	
	if s_num = "" then exit function
	
	s_ano = ""
	s_filhote = ""
	for i=1 to len(id_pedido)
		c = mid(id_pedido, i, 1)
		if IsLetra(c) then
			if s_ano = "" then
				s_ano = c
			elseif s_filhote = "" then
				s_filhote = c
				end if
			end if
		next

	if s_ano = "" then exit function
	s_num = normaliza_codigo(s_num, TAM_MIN_NUM_PEDIDO)
	s = s_num & s_ano
	if s_filhote <> "" then s = s & COD_SEPARADOR_FILHOTE & s_filhote
	normaliza_num_pedido = s
end function



' ------------------------------------------------------------------------
'   RETORNA_SUFIXO_PEDIDO_FILHOTE
function retorna_sufixo_pedido_filhote(byval id_pedido)
dim s
dim i	
	retorna_sufixo_pedido_filhote = ""
	id_pedido = Trim("" & id_pedido)
	id_pedido=normaliza_num_pedido(id_pedido)
	i = Instr(id_pedido, COD_SEPARADOR_FILHOTE)
	if i = 0 then exit function
	s = Mid(id_pedido, i+1)
	if len(s) <> 1 then exit function
	retorna_sufixo_pedido_filhote = s
end function



' ------------------------------------------------------------------------
'   RETORNA_NUM_PEDIDO_BASE
function retorna_num_pedido_base(byval id_pedido)
dim s
dim i
	id_pedido = Trim("" & id_pedido)
	id_pedido=normaliza_num_pedido(id_pedido)
	retorna_num_pedido_base = id_pedido
	i = Instr(id_pedido, COD_SEPARADOR_FILHOTE)
	if i = 0 then exit function
	s = Mid(id_pedido, 1, i-1)
	if s = "" then exit function
	retorna_num_pedido_base = s
end function



' ------------------------------------------------------------------------
'   IS PEDIDO FILHOTE
function IsPedidoFilhote(byval id_pedido)
dim i
	IsPedidoFilhote = False
	id_pedido = Trim("" & id_pedido)
	id_pedido=normaliza_num_pedido(id_pedido)
	i = Instr(id_pedido, COD_SEPARADOR_FILHOTE)
	if i = 0 then exit function
	IsPedidoFilhote = True
end function



' ------------------------------------------------------------------------
'   NORMALIZA_NUM_ORCAMENTO
function normaliza_num_orcamento(byval id_orcamento)
dim i
dim c
dim s
dim s_num
dim s_ano
	normaliza_num_orcamento = ""
	id_orcamento = Ucase(Trim("" & id_orcamento))
	if id_orcamento = "" then exit function
	s_num = ""
	for i=1 to len(id_orcamento)
		if IsNumeric(mid(id_orcamento,i,1)) then
			s_num=s_num & mid(id_orcamento,i,1)
		else
			exit for
			end if
		next
	
	if s_num = "" then exit function
	
	s_ano = ""
	for i=1 to len(id_orcamento)
		c = mid(id_orcamento, i, 1)
		if IsLetra(c) then
			if s_ano = "" then 
				s_ano = c
			else
				exit function
				end if
			end if
		next

	if s_ano = "" then exit function
	s_num = normaliza_codigo(s_num, TAM_MIN_NUM_ORCAMENTO)
	s = s_num & s_ano
	normaliza_num_orcamento = s
end function



' ---------------------------------------------------------------
'   LOG_PRODUTO_MONTA
function log_produto_monta(byval quantidade, byval id_fabricante, byval id_produto)
dim s
	s = " " & CStr(quantidade) & "x" & Trim(id_produto)
    if Trim(id_fabricante) <> "" then s = s & "(" & Trim(id_fabricante) & ")"
    log_produto_monta = s
end function



' ---------------------------------------------------------------
'   ESPERA
function espera(byval tempo_segundos)
dim inicio
	inicio = Now
	do while CLng(DateDiff("s", inicio, Now)) < tempo_segundos
		loop
end function



' ---------------------------------------------------------------
'   TEM DIGITO
function tem_digito(byval texto)
dim i	
dim achou
	tem_digito = False
	texto = Trim("" & texto)
	achou = False
	for i = 1 to len(texto)
		if IsNumeric(mid(texto, i, 1)) then
			achou = True
			exit for
			end if
		next
	if achou then tem_digito = True
end function



' ---------------------------------------------------------------
'   TEM VOGAL
function tem_vogal(byval texto)
dim i	
dim achou
dim letra
	tem_vogal = False
	texto = Trim("" & texto)
	achou = False
	for i = 1 to len(texto)
		letra = UCase(mid(texto, i, 1))
		if letra="A" Or letra="E" Or letra="I" Or letra="O" Or letra="U" then
			achou = True
			exit for
			end if
		next
	if achou then tem_vogal = True
end function



' ---------------------------------------------------------------
'   INICIAIS EM MAIUSCULAS
function iniciais_em_maiusculas(byval texto)
const palavras_minusculas = "|A|AS|AO|AOS|�|�S|E|O|OS|UM|UNS|UMA|UMAS|DA|DAS|DE|DO|DOS|EM|NA|NAS|NO|NOS|COM|SEM|POR|PELO|PELA|PARA|PRA|P/|S/|C/|TEM|OU|E/OU|ATE|AT�|QUE|SE|QUAL|"
const palavras_maiusculas = "|II|III|IV|VI|VII|VIII|IX|XI|XII|XIII|XIV|XV|XVI|XVII|XVIII|XIX|XX|XXI|XXII|XXIII|S/A|S/C|AC|AL|AM|AP|BA|CE|DF|ES|GO|MA|MG|MS|MT|PA|PB|PE|PI|PR|RJ|RN|RO|RR|RS|SC|SE|SP|TO|ME|EPP|"
dim letra
dim palavra
dim frase
dim s
dim i
dim i_max
dim blnAltera
	iniciais_em_maiusculas = ""
	frase = ""
	palavra = ""
	texto = "" & texto
	i_max = Len(texto)
	for i = 1 to i_max
		letra = mid(texto, i, 1)
		palavra = palavra & letra
		if (letra = " ") Or (i = i_max) Or (letra="(") Or (letra=")") Or (letra="[") Or (letra="]") Or (letra="'") Or (letra=chr(34)) Or (letra="-") then 
			s = "|" & UCase(Trim(palavra)) & "|"
			if (Instr(palavras_minusculas,s)<>0) And (frase<>"") then 
			'	SE FOR FINAL DA FRASE, DEIXA INALTERADO (EX: BLOCO A)
				if i < i_max then palavra = Lcase(palavra)
			elseif (Instr(palavras_maiusculas,s)<>0) then
				palavra=Ucase(palavra)
			else
			'	ANALISA SE CONVERTE O TEXTO OU N�O
				blnAltera = True
				if tem_digito(palavra) then
				'	ENDERE�OS CUJO N� DA RESID�NCIA S�O SEPARADOS POR V�RGULA, SEM NENHUM ESPA�O EM BRANCO
				'	CASO CONTR�RIO, CONSIDERA QUE � ALGUM TIPO DE C�DIGO
					if Instr(palavra, ",") = 0 then blnAltera = False
					end if
				if Instr(palavra, ".")<>0 then
				'	C.C.P.
					if Instr(Instr(palavra,".")+1, palavra, ".") <> 0 then blnAltera = False
					end if
				if Instr(palavra, "/")<>0 then
				'	S/C, S/A, S/C., S/A.
					if Len(palavra) <= 4 then blnAltera = False
					end if
				
			'	SIGLA?
				if Not tem_vogal(palavra) then blnAltera = False
					
				if blnAltera then palavra = Ucase(Left(palavra,1)) & Lcase(Mid(palavra,2))
				end if
			frase = frase & palavra
			palavra = ""
			end if
		next
	iniciais_em_maiusculas = frase
end function



' ---------------------------------------------------------------
'   IS EAN?
function IsEAN(byval codigo)
	codigo = Trim("" & codigo)
	IsEAN = (Len(codigo) = 13)
end function



' ------------------------------------------------------------------------
'   NORMALIZA_PRODUTO
function normaliza_produto(byval produto)
	normaliza_produto = ""
	produto = Ucase(Trim("" & produto))
	if produto = "" then exit function
'	NORMALIZA COM ZEROS � ESQUERDA SOMENTE SE O C�DIGO COME�A COM NUM�RICOS
	if Not IsNumeric(Left(produto,1)) then
		normaliza_produto = produto
		exit function
		end if
	normaliza_produto=normaliza_codigo(produto, TAM_MIN_PRODUTO)
end function



' ------------------------------------------------------------------------
'   QUICKSORT CL DUAS COLUNAS (OBS: ALGORITMO � RECURSIVO)
sub QuickSort_cl_duas_colunas(ByRef vetor, ByVal inf, ByVal sup)
dim i, j
dim ref, temp

	set ref = New cl_DUAS_COLUNAS
	set temp = New cl_DUAS_COLUNAS

  '�LA�O DE ORDENA��O
    Do
        i = inf
        j = sup
        ref.c1 = vetor((inf + sup) \ 2).c1
        ref.c2 = vetor((inf + sup) \ 2).c2
        
        Do
            Do
                If ref.c1 > vetor(i).c1 Then i = i + 1 Else Exit Do
                Loop

            Do
                If ref.c1 < vetor(j).c1 Then j = j - 1 Else Exit Do
                Loop

            If i <= j Then
                temp.c1 = vetor(i).c1
                temp.c2 = vetor(i).c2
                
                vetor(i).c1 = vetor(j).c1
                vetor(i).c2 = vetor(j).c2
                
                vetor(j).c1 = temp.c1
                vetor(j).c2 = temp.c2
                
                i = i + 1
                j = j - 1
                End If

            Loop Until i > j
        
        If inf < j Then QuickSort_cl_duas_colunas vetor, inf, j
        
        inf = i
        
        Loop Until i >= sup

end sub



' ------------------------------------------------------------------------
'   ORDENA CL DUAS COLUNAS
sub ordena_cl_duas_colunas(ByRef vetor, ByVal inf, ByVal sup)
    If inf > sup Then Exit Sub
    QuickSort_cl_duas_colunas vetor, inf, sup
end sub



' ------------------------------------------------------------------------
'   LOCALIZA CL DUAS COLUNAS (OBS: O VETOR PRECISA ESTAR ORDENADO)
function localiza_cl_duas_colunas(ByRef vetor, ByVal id, ByRef indice_localizado)
dim inf, sup, meio
dim s

	localiza_cl_duas_colunas = False

	id = Trim("" & id)
	indice_localizado = 0
	
  '�ESTABELECE LIMITES DE COMPARA��O INICIAIS
    sup = UBound(vetor)
    If sup > 0 Then inf = 1 Else Exit Function

 '�LA�O DE COMPARA��O
    Do While sup >= inf
        meio = (sup + inf) \ 2
        s = Trim("" & vetor(meio).c1)
        
      '�COMPARA CAMPO
        If (id > s) Then
            inf = meio + 1
        ElseIf (id < s) Then
            sup = meio - 1
        Else
            indice_localizado = meio
            localiza_cl_duas_colunas = True
            exit function
            End If
        Loop

end function



' ------------------------------------------------------------------------
'   QUICKSORT CL TRES COLUNAS (OBS: ALGORITMO � RECURSIVO)
sub QuickSort_cl_tres_colunas(ByRef vetor, ByVal inf, ByVal sup)
dim i, j
dim ref, temp

	set ref = New cl_TRES_COLUNAS
	set temp = New cl_TRES_COLUNAS

  '�LA�O DE ORDENA��O
    Do
        i = inf
        j = sup
        ref.c1 = vetor((inf + sup) \ 2).c1
        ref.c2 = vetor((inf + sup) \ 2).c2
        ref.c3 = vetor((inf + sup) \ 2).c3
        
        Do
            Do
                If ref.c1 > vetor(i).c1 Then i = i + 1 Else Exit Do
                Loop

            Do
                If ref.c1 < vetor(j).c1 Then j = j - 1 Else Exit Do
                Loop

            If i <= j Then
                temp.c1 = vetor(i).c1
                temp.c2 = vetor(i).c2
                temp.c3 = vetor(i).c3
                
                vetor(i).c1 = vetor(j).c1
                vetor(i).c2 = vetor(j).c2
                vetor(i).c3 = vetor(j).c3
                
                vetor(j).c1 = temp.c1
                vetor(j).c2 = temp.c2
                vetor(j).c3 = temp.c3
                
                i = i + 1
                j = j - 1
                End If

            Loop Until i > j
        
        If inf < j Then QuickSort_cl_tres_colunas vetor, inf, j
        
        inf = i
        
        Loop Until i >= sup

end sub



' ------------------------------------------------------------------------
'   ORDENA CL TRES COLUNAS
sub ordena_cl_tres_colunas(ByRef vetor, ByVal inf, ByVal sup)
    If inf > sup Then Exit Sub
    QuickSort_cl_tres_colunas vetor, inf, sup
end sub



' ------------------------------------------------------------------------
'   LOCALIZA CL TRES COLUNAS (OBS: O VETOR PRECISA ESTAR ORDENADO)
function localiza_cl_tres_colunas(ByRef vetor, ByVal id, ByRef indice_localizado)
dim inf, sup, meio
dim s

	localiza_cl_tres_colunas = False

	id = Trim("" & id)
	indice_localizado = 0
	
  '�ESTABELECE LIMITES DE COMPARA��O INICIAIS
    sup = UBound(vetor)
    If sup > 0 Then inf = 1 Else Exit Function

 '�LA�O DE COMPARA��O
    Do While sup >= inf
        meio = (sup + inf) \ 2
        s = Trim("" & vetor(meio).c1)
        
      '�COMPARA CAMPO
        If (id > s) Then
            inf = meio + 1
        ElseIf (id < s) Then
            sup = meio - 1
        Else
            indice_localizado = meio
            localiza_cl_tres_colunas = True
            exit function
            End If
        Loop

end function



' ------------------------------------------------------------------------
'   QUICKSORT CL QUATRO COLUNAS (OBS: ALGORITMO � RECURSIVO)
sub QuickSort_cl_quatro_colunas(ByRef vetor, ByVal inf, ByVal sup)
dim i, j
dim ref, temp

	set ref = New cl_QUATRO_COLUNAS
	set temp = New cl_QUATRO_COLUNAS

  '�LA�O DE ORDENA��O
    Do
        i = inf
        j = sup
        ref.c1 = vetor((inf + sup) \ 2).c1
        ref.c2 = vetor((inf + sup) \ 2).c2
        ref.c3 = vetor((inf + sup) \ 2).c3
        ref.c4 = vetor((inf + sup) \ 2).c4
        
        Do
            Do
                If ref.c1 > vetor(i).c1 Then i = i + 1 Else Exit Do
                Loop

            Do
                If ref.c1 < vetor(j).c1 Then j = j - 1 Else Exit Do
                Loop

            If i <= j Then
                temp.c1 = vetor(i).c1
                temp.c2 = vetor(i).c2
                temp.c3 = vetor(i).c3
                temp.c4 = vetor(i).c4
                
                vetor(i).c1 = vetor(j).c1
                vetor(i).c2 = vetor(j).c2
                vetor(i).c3 = vetor(j).c3
                vetor(i).c4 = vetor(j).c4
                
                vetor(j).c1 = temp.c1
                vetor(j).c2 = temp.c2
                vetor(j).c3 = temp.c3
                vetor(j).c4 = temp.c4
                
                i = i + 1
                j = j - 1
                End If

            Loop Until i > j
        
        If inf < j Then QuickSort_cl_quatro_colunas vetor, inf, j
        
        inf = i
        
        Loop Until i >= sup

end sub



' ------------------------------------------------------------------------
'   ORDENA CL QUATRO COLUNAS
sub ordena_cl_quatro_colunas(ByRef vetor, ByVal inf, ByVal sup)
    If inf > sup Then Exit Sub
    QuickSort_cl_quatro_colunas vetor, inf, sup
end sub



' ------------------------------------------------------------------------
'   LOCALIZA CL QUATRO COLUNAS (OBS: O VETOR PRECISA ESTAR ORDENADO)
function localiza_cl_quatro_colunas(ByRef vetor, ByVal id, ByRef indice_localizado)
dim inf, sup, meio
dim s

	localiza_cl_quatro_colunas = False

	id = Trim("" & id)
	indice_localizado = 0
	
  '�ESTABELECE LIMITES DE COMPARA��O INICIAIS
    sup = UBound(vetor)
    If sup > 0 Then inf = 1 Else Exit Function

 '�LA�O DE COMPARA��O
    Do While sup >= inf
        meio = (sup + inf) \ 2
        s = Trim("" & vetor(meio).c1)
        
      '�COMPARA CAMPO
        If (id > s) Then
            inf = meio + 1
        ElseIf (id < s) Then
            sup = meio - 1
        Else
            indice_localizado = meio
            localiza_cl_quatro_colunas = True
            exit function
            End If
        Loop

end function



' ------------------------------------------------------------------------
'   QUICKSORT CL CINCO COLUNAS (OBS: ALGORITMO � RECURSIVO)
sub QuickSort_cl_cinco_colunas(ByRef vetor, ByVal inf, ByVal sup)
dim i, j
dim ref, temp

	set ref = New cl_CINCO_COLUNAS
	set temp = New cl_CINCO_COLUNAS

  '�LA�O DE ORDENA��O
    Do
        i = inf
        j = sup
        ref.c1 = vetor((inf + sup) \ 2).c1
        ref.c2 = vetor((inf + sup) \ 2).c2
        ref.c3 = vetor((inf + sup) \ 2).c3
        ref.c4 = vetor((inf + sup) \ 2).c4
        ref.c5 = vetor((inf + sup) \ 2).c5
        
        Do
            Do
                If ref.c1 > vetor(i).c1 Then i = i + 1 Else Exit Do
                Loop

            Do
                If ref.c1 < vetor(j).c1 Then j = j - 1 Else Exit Do
                Loop

            If i <= j Then
                temp.c1 = vetor(i).c1
                temp.c2 = vetor(i).c2
                temp.c3 = vetor(i).c3
                temp.c4 = vetor(i).c4
                temp.c5 = vetor(i).c5
                
                vetor(i).c1 = vetor(j).c1
                vetor(i).c2 = vetor(j).c2
                vetor(i).c3 = vetor(j).c3
                vetor(i).c4 = vetor(j).c4
                vetor(i).c5 = vetor(j).c5
                
                vetor(j).c1 = temp.c1
                vetor(j).c2 = temp.c2
                vetor(j).c3 = temp.c3
                vetor(j).c4 = temp.c4
                vetor(j).c5 = temp.c5
                
                i = i + 1
                j = j - 1
                End If

            Loop Until i > j
        
        If inf < j Then QuickSort_cl_cinco_colunas vetor, inf, j
        
        inf = i
        
        Loop Until i >= sup

end sub



' ------------------------------------------------------------------------
'   ORDENA CL CINCO COLUNAS
sub ordena_cl_cinco_colunas(ByRef vetor, ByVal inf, ByVal sup)
    If inf > sup Then Exit Sub
    QuickSort_cl_cinco_colunas vetor, inf, sup
end sub



' ------------------------------------------------------------------------
'   LOCALIZA CL CINCO COLUNAS (OBS: O VETOR PRECISA ESTAR ORDENADO)
function localiza_cl_cinco_colunas(ByRef vetor, ByVal id, ByRef indice_localizado)
dim inf, sup, meio
dim s

	localiza_cl_cinco_colunas = False

	id = Trim("" & id)
	indice_localizado = 0
	
  '�ESTABELECE LIMITES DE COMPARA��O INICIAIS
    sup = UBound(vetor)
    If sup > 0 Then inf = 1 Else Exit Function

 '�LA�O DE COMPARA��O
    Do While sup >= inf
        meio = (sup + inf) \ 2
        s = Trim("" & vetor(meio).c1)
        
      '�COMPARA CAMPO
        If (id > s) Then
            inf = meio + 1
        ElseIf (id < s) Then
            sup = meio - 1
        Else
            indice_localizado = meio
            localiza_cl_cinco_colunas = True
            exit function
            End If
        Loop

end function



' ------------------------------------------------------------------------
'   QUICKSORT CL SEIS COLUNAS (OBS: ALGORITMO � RECURSIVO)
sub QuickSort_cl_seis_colunas(ByRef vetor, ByVal inf, ByVal sup)
dim i, j
dim ref, temp

	set ref = New cl_SEIS_COLUNAS
	set temp = New cl_SEIS_COLUNAS

  '�LA�O DE ORDENA��O
    Do
        i = inf
        j = sup
        ref.c1 = vetor((inf + sup) \ 2).c1
        ref.c2 = vetor((inf + sup) \ 2).c2
        ref.c3 = vetor((inf + sup) \ 2).c3
        ref.c4 = vetor((inf + sup) \ 2).c4
        ref.c5 = vetor((inf + sup) \ 2).c5
        ref.c6 = vetor((inf + sup) \ 2).c6
        
        Do
            Do
                If ref.c1 > vetor(i).c1 Then i = i + 1 Else Exit Do
                Loop

            Do
                If ref.c1 < vetor(j).c1 Then j = j - 1 Else Exit Do
                Loop

            If i <= j Then
                temp.c1 = vetor(i).c1
                temp.c2 = vetor(i).c2
                temp.c3 = vetor(i).c3
                temp.c4 = vetor(i).c4
                temp.c5 = vetor(i).c5
                temp.c6 = vetor(i).c6
                
                vetor(i).c1 = vetor(j).c1
                vetor(i).c2 = vetor(j).c2
                vetor(i).c3 = vetor(j).c3
                vetor(i).c4 = vetor(j).c4
                vetor(i).c5 = vetor(j).c5
                vetor(i).c6 = vetor(j).c6
                
                vetor(j).c1 = temp.c1
                vetor(j).c2 = temp.c2
                vetor(j).c3 = temp.c3
                vetor(j).c4 = temp.c4
                vetor(j).c5 = temp.c5
                vetor(j).c6 = temp.c6
                
                i = i + 1
                j = j - 1
                End If

            Loop Until i > j
        
        If inf < j Then QuickSort_cl_seis_colunas vetor, inf, j
        
        inf = i
        
        Loop Until i >= sup

end sub



' ------------------------------------------------------------------------
'   ORDENA CL SEIS COLUNAS
sub ordena_cl_seis_colunas(ByRef vetor, ByVal inf, ByVal sup)
    If inf > sup Then Exit Sub
    QuickSort_cl_seis_colunas vetor, inf, sup
end sub



' ------------------------------------------------------------------------
'   LOCALIZA CL SEIS COLUNAS (OBS: O VETOR PRECISA ESTAR ORDENADO)
function localiza_cl_seis_colunas(ByRef vetor, ByVal id, ByRef indice_localizado)
dim inf, sup, meio
dim s

	localiza_cl_seis_colunas = False

	id = Trim("" & id)
	indice_localizado = 0
	
  '�ESTABELECE LIMITES DE COMPARA��O INICIAIS
    sup = UBound(vetor)
    If sup > 0 Then inf = 1 Else Exit Function

 '�LA�O DE COMPARA��O
    Do While sup >= inf
        meio = (sup + inf) \ 2
        s = Trim("" & vetor(meio).c1)
        
      '�COMPARA CAMPO
        If (id > s) Then
            inf = meio + 1
        ElseIf (id < s) Then
            sup = meio - 1
        Else
            indice_localizado = meio
            localiza_cl_seis_colunas = True
            exit function
            End If
        Loop

end function



' ------------------------------------------------------------------------
'   QUICKSORT CL DEZ COLUNAS (OBS: ALGORITMO � RECURSIVO)
sub QuickSort_cl_dez_colunas(ByRef vetor, ByVal inf, ByVal sup)
dim i, j, intIdxRef
dim ref, temp

	set ref = New cl_DEZ_COLUNAS
	set temp = New cl_DEZ_COLUNAS

  '�LA�O DE ORDENA��O
    Do
        i = inf
        j = sup
        
        intIdxRef = (inf + sup) \ 2
        ref.CampoOrdenacao = vetor(intIdxRef).CampoOrdenacao
        ref.c1 = vetor(intIdxRef).c1
        ref.c2 = vetor(intIdxRef).c2
        ref.c3 = vetor(intIdxRef).c3
        ref.c4 = vetor(intIdxRef).c4
        ref.c5 = vetor(intIdxRef).c5
        ref.c6 = vetor(intIdxRef).c6
        ref.c7 = vetor(intIdxRef).c7
        ref.c8 = vetor(intIdxRef).c8
        ref.c9 = vetor(intIdxRef).c9
        ref.c10 = vetor(intIdxRef).c10
        
        Do
            Do
                If ref.CampoOrdenacao > vetor(i).CampoOrdenacao Then i = i + 1 Else Exit Do
                Loop

            Do
                If ref.CampoOrdenacao < vetor(j).CampoOrdenacao Then j = j - 1 Else Exit Do
                Loop

            If i <= j Then
				temp.CampoOrdenacao = vetor(i).CampoOrdenacao
                temp.c1 = vetor(i).c1
                temp.c2 = vetor(i).c2
                temp.c3 = vetor(i).c3
                temp.c4 = vetor(i).c4
                temp.c5 = vetor(i).c5
                temp.c6 = vetor(i).c6
                temp.c7 = vetor(i).c7
                temp.c8 = vetor(i).c8
                temp.c9 = vetor(i).c9
                temp.c10 = vetor(i).c10
                
                vetor(i).CampoOrdenacao = vetor(j).CampoOrdenacao
                vetor(i).c1 = vetor(j).c1
                vetor(i).c2 = vetor(j).c2
                vetor(i).c3 = vetor(j).c3
                vetor(i).c4 = vetor(j).c4
                vetor(i).c5 = vetor(j).c5
                vetor(i).c6 = vetor(j).c6
                vetor(i).c7 = vetor(j).c7
                vetor(i).c8 = vetor(j).c8
                vetor(i).c9 = vetor(j).c9
                vetor(i).c10 = vetor(j).c10
                
                vetor(j).CampoOrdenacao = temp.CampoOrdenacao
                vetor(j).c1 = temp.c1
                vetor(j).c2 = temp.c2
                vetor(j).c3 = temp.c3
                vetor(j).c4 = temp.c4
                vetor(j).c5 = temp.c5
                vetor(j).c6 = temp.c6
                vetor(j).c7 = temp.c7
                vetor(j).c8 = temp.c8
                vetor(j).c9 = temp.c9
                vetor(j).c10 = temp.c10
                
                i = i + 1
                j = j - 1
                End If

            Loop Until i > j
        
        If inf < j Then QuickSort_cl_dez_colunas vetor, inf, j
        
        inf = i
        
        Loop Until i >= sup

end sub



' ------------------------------------------------------------------------
'   ORDENA CL DEZ COLUNAS
sub ordena_cl_dez_colunas(ByRef vetor, ByVal inf, ByVal sup)
    If inf > sup Then Exit Sub
    QuickSort_cl_dez_colunas vetor, inf, sup
end sub



' ------------------------------------------------------------------------
'   LOCALIZA CL DEZ COLUNAS (OBS: O VETOR PRECISA ESTAR ORDENADO PELO CAMPO USADO P/ LOCALIZA��O)
function localiza_cl_dez_colunas(ByRef vetor, ByVal id, ByRef indice_localizado)
dim inf, sup, meio
dim s

	localiza_cl_dez_colunas = False

	id = Trim("" & id)
	indice_localizado = 0
	
  '�ESTABELECE LIMITES DE COMPARA��O INICIAIS
    sup = UBound(vetor)
    If sup > 0 Then inf = 1 Else Exit Function

 '�LA�O DE COMPARA��O
    Do While sup >= inf
        meio = (sup + inf) \ 2
        s = Trim("" & vetor(meio).CampoOrdenacao)
        
      '�COMPARA CAMPO
        If (id > s) Then
            inf = meio + 1
        ElseIf (id < s) Then
            sup = meio - 1
        Else
            indice_localizado = meio
            localiza_cl_dez_colunas = True
            exit function
            End If
        Loop

end function



' ------------------------------------------------------------------------
'   QUICKSORT CL VINTE COLUNAS (OBS: ALGORITMO � RECURSIVO)
sub QuickSort_cl_vinte_colunas(ByRef vetor, ByVal inf, ByVal sup)
dim i, j, intIdxRef
dim ref, temp

	set ref = New cl_VINTE_COLUNAS
	set temp = New cl_VINTE_COLUNAS

  '�LA�O DE ORDENA��O
    Do
        i = inf
        j = sup
        
        intIdxRef = (inf + sup) \ 2
        ref.CampoOrdenacao = vetor(intIdxRef).CampoOrdenacao
        ref.c1 = vetor(intIdxRef).c1
        ref.c2 = vetor(intIdxRef).c2
        ref.c3 = vetor(intIdxRef).c3
        ref.c4 = vetor(intIdxRef).c4
        ref.c5 = vetor(intIdxRef).c5
        ref.c6 = vetor(intIdxRef).c6
        ref.c7 = vetor(intIdxRef).c7
        ref.c8 = vetor(intIdxRef).c8
        ref.c9 = vetor(intIdxRef).c9
        ref.c10 = vetor(intIdxRef).c10
        ref.c11 = vetor(intIdxRef).c11
        ref.c12 = vetor(intIdxRef).c12
        ref.c13 = vetor(intIdxRef).c13
        ref.c14 = vetor(intIdxRef).c14
        ref.c15 = vetor(intIdxRef).c15
        ref.c16 = vetor(intIdxRef).c16
        ref.c17 = vetor(intIdxRef).c17
        ref.c18 = vetor(intIdxRef).c18
        ref.c19 = vetor(intIdxRef).c19
        ref.c20 = vetor(intIdxRef).c20
        
        Do
            Do
                If ref.CampoOrdenacao > vetor(i).CampoOrdenacao Then i = i + 1 Else Exit Do
                Loop

            Do
                If ref.CampoOrdenacao < vetor(j).CampoOrdenacao Then j = j - 1 Else Exit Do
                Loop

            If i <= j Then
				temp.CampoOrdenacao = vetor(i).CampoOrdenacao
                temp.c1 = vetor(i).c1
                temp.c2 = vetor(i).c2
                temp.c3 = vetor(i).c3
                temp.c4 = vetor(i).c4
                temp.c5 = vetor(i).c5
                temp.c6 = vetor(i).c6
                temp.c7 = vetor(i).c7
                temp.c8 = vetor(i).c8
                temp.c9 = vetor(i).c9
                temp.c10 = vetor(i).c10
                temp.c11 = vetor(i).c11
                temp.c12 = vetor(i).c12
                temp.c13 = vetor(i).c13
                temp.c14 = vetor(i).c14
                temp.c15 = vetor(i).c15
                temp.c16 = vetor(i).c16
                temp.c17 = vetor(i).c17
                temp.c18 = vetor(i).c18
                temp.c19 = vetor(i).c19
                temp.c20 = vetor(i).c20
                
                vetor(i).CampoOrdenacao = vetor(j).CampoOrdenacao
                vetor(i).c1 = vetor(j).c1
                vetor(i).c2 = vetor(j).c2
                vetor(i).c3 = vetor(j).c3
                vetor(i).c4 = vetor(j).c4
                vetor(i).c5 = vetor(j).c5
                vetor(i).c6 = vetor(j).c6
                vetor(i).c7 = vetor(j).c7
                vetor(i).c8 = vetor(j).c8
                vetor(i).c9 = vetor(j).c9
                vetor(i).c10 = vetor(j).c10
                vetor(i).c11 = vetor(j).c11
                vetor(i).c12 = vetor(j).c12
                vetor(i).c13 = vetor(j).c13
                vetor(i).c14 = vetor(j).c14
                vetor(i).c15 = vetor(j).c15
                vetor(i).c16 = vetor(j).c16
                vetor(i).c17 = vetor(j).c17
                vetor(i).c18 = vetor(j).c18
                vetor(i).c19 = vetor(j).c19
                vetor(i).c20 = vetor(j).c20
                
                vetor(j).CampoOrdenacao = temp.CampoOrdenacao
                vetor(j).c1 = temp.c1
                vetor(j).c2 = temp.c2
                vetor(j).c3 = temp.c3
                vetor(j).c4 = temp.c4
                vetor(j).c5 = temp.c5
                vetor(j).c6 = temp.c6
                vetor(j).c7 = temp.c7
                vetor(j).c8 = temp.c8
                vetor(j).c9 = temp.c9
                vetor(j).c10 = temp.c10
                vetor(j).c11 = temp.c11
                vetor(j).c12 = temp.c12
                vetor(j).c13 = temp.c13
                vetor(j).c14 = temp.c14
                vetor(j).c15 = temp.c15
                vetor(j).c16 = temp.c16
                vetor(j).c17 = temp.c17
                vetor(j).c18 = temp.c18
                vetor(j).c19 = temp.c19
                vetor(j).c20 = temp.c20
                
                i = i + 1
                j = j - 1
                End If

            Loop Until i > j
        
        If inf < j Then QuickSort_cl_vinte_colunas vetor, inf, j
        
        inf = i
        
        Loop Until i >= sup

end sub



' ------------------------------------------------------------------------
'   ORDENA CL VINTE COLUNAS
sub ordena_cl_vinte_colunas(ByRef vetor, ByVal inf, ByVal sup)
    If inf > sup Then Exit Sub
    QuickSort_cl_vinte_colunas vetor, inf, sup
end sub



' ------------------------------------------------------------------------
'   LOCALIZA CL VINTE COLUNAS (OBS: O VETOR PRECISA ESTAR ORDENADO PELO CAMPO USADO P/ LOCALIZA��O)
function localiza_cl_vinte_colunas(ByRef vetor, ByVal id, ByRef indice_localizado)
dim inf, sup, meio
dim s

	localiza_cl_vinte_colunas = False

	id = Trim("" & id)
	indice_localizado = 0
	
  '�ESTABELECE LIMITES DE COMPARA��O INICIAIS
    sup = UBound(vetor)
    If sup > 0 Then inf = 1 Else Exit Function

 '�LA�O DE COMPARA��O
    Do While sup >= inf
        meio = (sup + inf) \ 2
        s = Trim("" & vetor(meio).CampoOrdenacao)
        
      '�COMPARA CAMPO
        If (id > s) Then
            inf = meio + 1
        ElseIf (id < s) Then
            sup = meio - 1
        Else
            indice_localizado = meio
            localiza_cl_vinte_colunas = True
            exit function
            End If
        Loop

end function



' ------------------------------------------------------------------------
'   MES POR EXTENSO
'   Retorna o m�s por extenso ou apenas c/ os 3 primeiros caracteres.
function mes_por_extenso(ByVal i, byval por_extenso)
dim s
dim j
    If IsNumeric(i) Then j = CInt(i) Else j = 0

    select case j
        case 1: s = "JANEIRO"
        case 2: s = "FEVEREIRO"
        case 3: s = "MAR�O"
        case 4: s = "ABRIL"
        case 5: s = "MAIO"
        case 6: s = "JUNHO"
        case 7: s = "JULHO"
        case 8: s = "AGOSTO"
        case 9: s = "SETEMBRO"
        case 10: s = "OUTUBRO"
        case 11: s = "NOVEMBRO"
        case 12: s = "DEZEMBRO"
        case else: s = ""
        end select

    If Not por_extenso Then s = Mid(s, 1, 3)
    mes_por_extenso = s
end function



' ------------------------------------------------------------------------
'   DIA DA SEMANA
'   Retorna o dia da semana por extenso ou apenas c/ os 3 primeiros caracteres.
function dia_da_semana(byval data, byval por_extenso)
dim s
	dia_da_semana = ""
	if Not IsDate(data) then exit function
	
	select case WeekDay(data)
		case 1: s = "DOMINGO"
		case 2: s = "SEGUNDA"
		case 3: s = "TER�A"
		case 4: s = "QUARTA"
		case 5: s = "QUINTA"
		case 6: s = "SEXTA"
		case 7: s = "S�BADO"
		case else: s = ""
		end select

    If Not por_extenso Then s = Mid(s, 1, 3)
	dia_da_semana = s
end function



' ------------------------------------------------------------------------
'   FILTRA TEXTO JAVASCRIPT
'   Monta o texto para ser tratado em JavaScript.
function filtra_texto_js(byval texto, byval delimitador_texto)
dim s
	s = texto	
	if delimitador_texto = chr(34) then
		s = replace(s, chr(34), delimitador_texto & "+KEY_ASPAS+" & delimitador_texto)
		s = replace(s, "'", delimitador_texto & "+KEY_APOSTROFE+" & delimitador_texto)
	else
		s = replace(s, "'", delimitador_texto & "+KEY_APOSTROFE+" & delimitador_texto)
		s = replace(s, chr(34), delimitador_texto & "+KEY_ASPAS+" & delimitador_texto)
		end if
	s = replace(s, vbCrLf, delimitador_texto & "+KEY_CRLF+" & delimitador_texto)
	s = replace(s, vbLf & vbCr, delimitador_texto & "+KEY_LFCR+" & delimitador_texto)
	s = replace(s, vbCr, delimitador_texto & "+KEY_RETURN+" & delimitador_texto)
	s = replace(s, vbLf, delimitador_texto & "+KEY_LINEFEED+" & delimitador_texto)
	filtra_texto_js = s
end function



' ------------------------------------------------------------------------
'   VISANET OPERACAO DESCRICAO
'   Retorna a descri��o para o c�digo de opera��o.
function visanet_operacao_descricao(byval codigo_operacao)
dim s_resp

	select case codigo_operacao
		case OP_VISANET_PAGAMENTO
			s_resp = "Pagamento"
		case OP_VISANET_CANCELAMENTO
			s_resp = "Cancelamento"
		case else
			s_resp = ""
	end select
	
	visanet_operacao_descricao = s_resp
	
end function



' ------------------------------------------------------------------------
'   VISANET TRANSACAO PAGAMENTO AUTORIZADA
'   Indica se o c�digo de retorno � de autoriza��o ou rejei��o.
function visanet_transacao_pagamento_autorizada(byval codigo_LR)

	visanet_transacao_pagamento_autorizada = False
	
	If (Cstr(codigo_LR)<>"0") And (Cstr(codigo_LR)<>"00") And (Cstr(codigo_LR)<>"11") then exit function
	
	visanet_transacao_pagamento_autorizada = True
	
end function



' ------------------------------------------------------------------------
'   VISANET TRANSACAO CANCELAMENTO APROVADA
'   Indica se o c�digo de retorno � de aprova��o ou rejei��o.
function visanet_transacao_cancelamento_aprovada(byval codigo_LR)
	visanet_transacao_cancelamento_aprovada = False
	If (Cstr(codigo_LR)<>"0") then exit function
	visanet_transacao_cancelamento_aprovada = True
end function



' ------------------------------------------------------------------------
'   VISANET DESCRICAO ORIGEM EMISSAO
'   Retorna a descri��o para a origem do cart�o.
function visanet_descricao_origem_emissao(byval codigo_LR)
dim s_resp
	
	if (Cstr(codigo_LR) = "0") Or (Cstr(codigo_LR) = "00") then
		s_resp = "Cart�o Emitido no Brasil"
	else
		s_resp = "Cart�o Emitido no Exterior"
		end if

	visanet_descricao_origem_emissao = s_resp
end function



' ------------------------------------------------------------------------
'   VISANET DESCRICAO AUTENTICACAO
'   Retorna a descri��o para o c�digo de autentica��o.
function visanet_descricao_autenticacao(byval codigo_AUTHENT)
dim s_resp

	if codigo_AUTHENT = "1" then
		s_resp = "Autentica��o com Sucesso"
	elseif codigo_AUTHENT = "2" then
		s_resp = "Autentica��o Negada"
	elseif codigo_AUTHENT = "3" then
		s_resp = "Falha na Autentica��o"
	elseif codigo_AUTHENT = "" then
		s_resp = "Transa��o Sem Autentica��o"
	else 
		s_resp = ""
		end if
	
	visanet_descricao_autenticacao = s_resp
	
end function



' ------------------------------------------------------------------------
'   VISANET DESCRICAO PARCELAMENTO
'   Retorna a descri��o para a forma de pagamento selecionada.
function visanet_descricao_parcelamento(byval codigo_parcelamento, byval valor_total)
dim s_resp
dim cod_produto
dim qtde_parcelas
dim vl_parcela
dim vl_total

	codigo_parcelamento = Trim("" & codigo_parcelamento)
	
	cod_produto = Left(codigo_parcelamento, 1)

	vl_total = converte_numero(valor_total)
	qtde_parcelas = converte_numero(mid(codigo_parcelamento, 2))
	if qtde_parcelas <> 0 then vl_parcela = vl_total / qtde_parcelas

	select case cod_produto
	'	VISA: CR�DITO � VISTA
		case "1"
			s_resp = SIMBOLO_MONETARIO & " " & formata_moeda(valor_total) & " � Vista"
	'	VISA: PARCELADO LOJISTA
		case "2"
			s_resp = formata_inteiro(qtde_parcelas) & " x de " & SIMBOLO_MONETARIO & " " & formata_moeda(vl_parcela) & " iguais"
	'	VISA: PARCELADO EMISSOR
		case "3"
			s_resp = formata_inteiro(qtde_parcelas) & " x de " & SIMBOLO_MONETARIO & " " & formata_moeda(vl_parcela) & " mais juros"
	'	VISA: � VISTA - D�BITO
		case "A"
			s_resp = SIMBOLO_MONETARIO & " " & formata_moeda(valor_total) & " D�bito � Vista"
	'	MASTERCARD: CR�DITO � VISTA
		case "4"
			s_resp = SIMBOLO_MONETARIO & " " & formata_moeda(valor_total) & " � Vista"
	'	MASTERCARD: PARCELADO LOJISTA
		case "5"
			s_resp = formata_inteiro(qtde_parcelas) & " x de " & SIMBOLO_MONETARIO & " " & formata_moeda(vl_parcela) & " iguais"
	'	MASTERCARD: PARCELADO EMISSOR
		case "6"
			s_resp = formata_inteiro(qtde_parcelas) & " x de " & SIMBOLO_MONETARIO & " " & formata_moeda(vl_parcela) & " mais juros"
	'	ELO: CR�DITO � VISTA
		case "7"
			s_resp = SIMBOLO_MONETARIO & " " & formata_moeda(valor_total) & " � Vista"
	'	ELO: PARCELADO LOJISTA
		case "8"
			s_resp = formata_inteiro(qtde_parcelas) & " x de " & SIMBOLO_MONETARIO & " " & formata_moeda(vl_parcela) & " iguais"
	'	ELO: PARCELADO EMISSOR
		case "9"
			s_resp = formata_inteiro(qtde_parcelas) & " x de " & SIMBOLO_MONETARIO & " " & formata_moeda(vl_parcela) & " mais juros"
		case else
			s_resp = ""
	end select
	
	visanet_descricao_parcelamento = s_resp
	
end function


' ------------------------------------------------------------------------
'   IS MESMO ANO E MES
'   Verifica se as duas datas est�o dentro do mesmo ano e m�s.
function IsMesmoAnoEMes(byval data1, byval data2)

	IsMesmoAnoEMes = False
	
	if (Not IsDate(data1)) Or (Not IsDate(data2)) then exit function
	
	if VarType(data1) = vbString then data1 = StrToDate(data1)
	if VarType(data2) = vbString then data2 = StrToDate(data2)
	
	if VarType(data1) <> vbDate then data1 = CDate(data1)
	if VarType(data2) <> vbDate then data2 = CDate(data2)
	
	if Year(data1) <> Year(data2) then exit function
	if Month(data1) <> Month(data2) then exit function
	
	IsMesmoAnoEMes = True
	
end function


' ------------------------------------------------------------------------
'   FORMATA ENDERECO
'   Formata os campos do endere�o em um texto formatado.
function formata_endereco(endereco, endereco_numero, endereco_complemento, bairro, cidade, uf, cep)
dim s_aux, strResposta
	strResposta = ""
	if Trim(endereco) <> "" then
		strResposta = Trim(endereco)
		s_aux=Trim(endereco_numero)
		if s_aux<>"" then strResposta=strResposta & ", " & s_aux
		s_aux=Trim(endereco_complemento)
		if s_aux<>"" then strResposta=strResposta & " " & s_aux
		s_aux=Trim(bairro)
		if s_aux<>"" then strResposta=strResposta & " - " & s_aux
		s_aux=Trim(cidade)
		if s_aux<>"" then strResposta=strResposta & " - " & s_aux
		s_aux=Trim(uf)
		if s_aux<>"" then strResposta=strResposta & " - " & s_aux
		s_aux=Trim(cep)
		if s_aux<>"" then strResposta=strResposta & " - " & cep_formata(s_aux)
		end if
	formata_endereco=strResposta
end function


' ------------------------------------------------------------------------
'   FORMATA ENDERECO DE ENTREGA DE UM PEDIDO
'   Formata os campos do endere�o em um texto formatado.
function pedido_formata_endereco_entrega(r_pedido, r_cliente)
dim s_cabecalho, s_aux, s_tel_aux_1, s_tel_aux_2, s_telefones, s_endereco, s_email
    with r_pedido
		s_endereco = formata_endereco(.EndEtg_endereco, .EndEtg_endereco_numero, .EndEtg_endereco_complemento, .EndEtg_bairro, .EndEtg_cidade, .EndEtg_uf, .EndEtg_cep)
		end with
	
    pedido_formata_endereco_entrega=s_endereco

    'tem endere�o de entrega diferente?
    if r_pedido.st_end_entrega = 0 then exit function

    'se a memoriza��o n�o estiver ativa ou o registro foi criado no formato antigo, paramos por aqui
    if not isActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos or r_pedido.st_memorizacao_completa_enderecos = 0 then exit function

    'PF, somente e-mails adicionais
    if r_cliente.tipo = ID_PF then 
		'EndEtg_email e EndEtg_email_xml 
		s_email = ""
		if Trim("" & r_pedido.EndEtg_email) <> "" or Trim("" & r_pedido.EndEtg_email_xml) <> ""  then
			s_email = "<br>"
			if Trim("" & r_pedido.EndEtg_email) <> "" then
				s_email = s_email & "E-mail: " & r_pedido.EndEtg_email & " "
				end if
			if Trim("" & r_pedido.EndEtg_email_xml) <> "" then
				s_email = s_email & "E-mail (XML): " & r_pedido.EndEtg_email_xml & " "
				end if
			end if

        pedido_formata_endereco_entrega = s_endereco + s_email
		exit function
		end if

    'memoriza��o ativa, colocamos os campos adicionais
    if r_pedido.EndEtg_tipo_pessoa = ID_PF then
        'Nome, CPF, Produto rural, ICMS, IE
        'Exemplo: Teste de Nome Para Entrega - CPF: 089.617.758/04 - Produtor rural: Sim (IE: 244.355.757.113)
        s_cabecalho = r_pedido.EndEtg_nome + "<br>CPF: " + cnpj_cpf_formata(r_pedido.EndEtg_cnpj_cpf)
        s_aux = ""
        if r_pedido.EndEtg_produtor_rural_status = converte_numero(COD_ST_CLIENTE_PRODUTOR_RURAL_SIM) then
                if r_pedido.EndEtg_contribuinte_icms_status = converte_numero(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO) then
                    s_aux = "Sim (N�o contribuinte)"
                elseif r_pedido.EndEtg_contribuinte_icms_status = converte_numero(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM) then
                    s_aux = "Sim (IE: " & r_pedido.EndEtg_ie & ")"
                elseif r_pedido.EndEtg_contribuinte_icms_status = converte_numero(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_ISENTO) then
                    s_aux = "Sim (Isento)"
                end if
            elseif r_pedido.EndEtg_produtor_rural_status = converte_numero(COD_ST_CLIENTE_PRODUTOR_RURAL_NAO) then
                s_aux = "N�o"
            end if
        if s_aux <> "" then s_cabecalho = s_cabecalho  + " - Produtor rural: " + s_aux 
        s_cabecalho = s_cabecalho  + "<br>"

        'telefones, formato: 
        'Telefone (11) 1234-1234 - Celular (99) 90090-0099 
        s_tel_aux_1 = formata_ddd_telefone_ramal(r_pedido.EndEtg_ddd_res, r_pedido.EndEtg_tel_res, "")
        s_tel_aux_2 = formata_ddd_telefone_ramal(r_pedido.EndEtg_ddd_cel, r_pedido.EndEtg_tel_cel, "")

        s_telefones = ""
        if s_tel_aux_1 <> "" or s_tel_aux_2 <> "" then s_telefones = "<br>"
        if s_tel_aux_1 <> "" then 
            s_telefones = s_telefones + "Telefone " + s_tel_aux_1
            if s_tel_aux_2 <> "" then s_telefones = s_telefones + " - "
            end if
        
        if s_tel_aux_2 <> "" then s_telefones = s_telefones + "Celular " + s_tel_aux_2
    
		'EndEtg_email e EndEtg_email_xml 
		s_email = ""
		if Trim("" & r_pedido.EndEtg_email) <> "" or Trim("" & r_pedido.EndEtg_email_xml) <> ""  then
			s_email = "<br>"
			if Trim("" & r_pedido.EndEtg_email) <> "" then
				s_email = s_email & "E-mail: " & r_pedido.EndEtg_email & " "
				end if
			if Trim("" & r_pedido.EndEtg_email_xml) <> "" then
				s_email = s_email & "E-mail (XML): " & r_pedido.EndEtg_email_xml & " "
				end if
			end if

        pedido_formata_endereco_entrega = s_cabecalho + s_endereco + s_telefones + s_email
        exit function
        end if

    'o endere�o de entrega � de PJ
    'Nome, CNPJ, ICMS, IE
    'Nome de teste de outra empresa - CNPJ: 01.051.970/0001-89 - Contribuinte ICMS: Sim (IE: 244.355.757.113)
    s_cabecalho = r_pedido.EndEtg_nome + "<br>CNPJ: " + cnpj_cpf_formata(r_pedido.EndEtg_cnpj_cpf)
    s_aux = ""
    if r_pedido.EndEtg_contribuinte_icms_status = converte_numero(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO) then
        s_aux = "N�o"
    elseif r_pedido.EndEtg_contribuinte_icms_status = converte_numero(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM) then
        s_aux = "Sim (IE: " & r_pedido.EndEtg_ie & ")"
    elseif r_pedido.EndEtg_contribuinte_icms_status = converte_numero(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_ISENTO) then
        s_aux = "Isento"
    end if
    if s_aux <> "" then s_cabecalho = s_cabecalho  + " - Contribuinte ICMS: " + s_aux 
    s_cabecalho = s_cabecalho  + "<br>"

    'Telefone (11) 1234-1234 - Celular (99) 90090-0099 
    s_tel_aux_1 = formata_ddd_telefone_ramal(r_pedido.EndEtg_ddd_com, r_pedido.EndEtg_tel_com, r_pedido.EndEtg_ramal_com)
    s_tel_aux_2 = formata_ddd_telefone_ramal(r_pedido.EndEtg_ddd_com_2, r_pedido.EndEtg_tel_com_2, r_pedido.EndEtg_ramal_com_2)

    s_telefones = ""
    if s_tel_aux_1 <> "" or s_tel_aux_2 <> "" then s_telefones = "<br>Telefone "
    if s_tel_aux_1 <> "" then 
        s_telefones = s_telefones + s_tel_aux_1
        if s_tel_aux_2 <> "" then s_telefones = s_telefones + " - "
        end if
        
    if s_tel_aux_2 <> "" then s_telefones = s_telefones + s_tel_aux_2

	'EndEtg_email e EndEtg_email_xml 
	s_email = ""
	if Trim("" & r_pedido.EndEtg_email) <> "" or Trim("" & r_pedido.EndEtg_email_xml) <> ""  then
		s_email = "<br>"
		if Trim("" & r_pedido.EndEtg_email) <> "" then
			s_email = s_email & "E-mail: " & r_pedido.EndEtg_email & " "
			end if
		if Trim("" & r_pedido.EndEtg_email_xml) <> "" then
			s_email = s_email & "E-mail (XML): " & r_pedido.EndEtg_email_xml & " "
			end if
		end if

    pedido_formata_endereco_entrega = s_cabecalho + s_endereco + s_telefones + s_email

end function


' ------------------------------------------------------------------------
'   FORMATA DDD TELEFONE RAMAL
'   Formata os campos de telefone.
function formata_ddd_telefone_ramal(ddd, telefone, ramal)
dim s_tel, i, s_aux, strResposta
	strResposta = ""

'	FORMATA A PARCELA RELATIVA AO N�MERO DO TELEFONE
	s_tel = "" & telefone
	s_tel = retorna_so_digitos(s_tel)
	
	if ((s_tel="") Or (len(s_tel)>9) Or (Not telefone_ok(s_tel))) then 
		'NOP
	else
		i=len(s_tel)-4
		s_tel = mid(s_tel, 1, i) & "-" & mid(s_tel, i+1, len(s_tel))
		end if

	strResposta = s_tel

'	FORMATA AGRUPANDO O DDD E O RAMAL
	if strResposta <> "" then
		s_aux = Trim("" & ddd)
		if s_aux <> "" then strResposta = "(" & s_aux & ") " & strResposta
		s_aux = Trim("" & ramal)
		if s_aux <> "" then strResposta = strResposta & "  (R. " & s_aux & ")"
		end if

	formata_ddd_telefone_ramal=strResposta
end function



' ------------------------------------------------------------------------
'   RETIRAZEROSAESQUERDA
'   Retira todos os zeros que estejam � esquerda.
'	Ex:  0000123 = 123
'		-00000.23 = -0.23
'		 0000,12 = 0,12
'		 0000000 = 0
function RetiraZerosAEsquerda(byval numero)
dim strSinal
	numero=Trim("" & numero)
	strSinal = ""
	if Left(numero, 1) = "-" then 
		strSinal="-"
		numero=mid(numero, 2)
		end if
	do while (Left(numero, 1) = "0") And (len(numero) > 1)
		numero=mid(numero, 2)
		loop
	if Left(numero, 1) = "," Or Left(numero, 1) = "." then 
		numero= "0" & numero
		end if
	RetiraZerosAEsquerda = strSinal & numero
end function



' ------------------------------------------------------------------------
'   FORMATA NUM OS TELA
'   Formata o n� identifica��o da ordem de servi�o p/ exibi��o na tela.
function formata_num_OS_tela(byval numero)
	numero=Trim("" & numero)
	numero=retorna_so_digitos(numero)
	numero=RetiraZerosAEsquerda(numero)
	numero=formata_inteiro(numero)
	numero=normaliza_codigo(numero, 3)
	formata_num_OS_tela=numero
end function



' ------------------------------------------------------------------------
'   GERA TICKET SESSION CTRL
'   Gera um n� de ticket p/ esta sess�o.
function GeraTicketSessionCtrl(Byval strUsuario)
dim strTicketBase
dim strChave
	strTicketBase = Trim(strUsuario) & formata_data_yyyymmdd(Date) & formata_hora_hhnnss(Now)
	strChave = gera_chave(FATOR_CRIPTO_SESSION_CTRL)
	GeraTicketSessionCtrl = CriptografaTexto(strTicketBase, strChave)
end function



' ------------------------------------------------------------------------
'   MONTA SESSION CTRL INFO
'   Monta o conjunto de dados criptografados usados p/ recuperar
'   a sess�o expirada.
function MontaSessionCtrlInfo(ByVal strUsuario, ByVal strModulo, ByVal strLoja, ByVal strTicket, ByVal dtLogon, ByVal dtUltAtividade)
dim strSessionCtrlParametro
dim strSessionCtrlParametroCripto
dim strChaveCripto
	strUsuario = Trim("" & strUsuario)
	strModulo = Trim("" & strModulo)
	strLoja = Trim("" & strLoja)
	strTicket = Trim("" & strTicket)
	strSessionCtrlParametro = strUsuario & "|" & strModulo & "|" & strLoja & "|" & strTicket & "|" & CStr(CDbl(dtLogon)) & "|" & CStr(CDbl(dtUltAtividade))
	strChaveCripto = gera_chave(FATOR_CRIPTO_SESSION_CTRL)
	strSessionCtrlParametroCripto = CriptografaTexto(strSessionCtrlParametro, strChaveCripto)
	MontaSessionCtrlInfo = strSessionCtrlParametroCripto
end function



' ------------------------------------------------------------------------
'   ATUALIZA SESSION CTRL INFO DATAHORAULTATIVIDADE
'   Atualiza a data/hora da �ltima atividade p/ que o c�lculo do tempo
'   de sess�o ociosa seja correto.
function AtualizaSessionCtrlInfoDataHoraUltAtividade(ByVal strSessionCtrlInfo)
dim strChave
dim strDecriptografado
dim vVetor
	strSessionCtrlInfo = Trim("" & strSessionCtrlInfo)
	AtualizaSessionCtrlInfoDataHoraUltAtividade=strSessionCtrlInfo
	if strSessionCtrlInfo = "" then exit function
	'Decriptografa o par�metro
	strChave = gera_chave(FATOR_CRIPTO_SESSION_CTRL)
	strDecriptografado = DecriptografaTexto(strSessionCtrlInfo, strChave)
	'Separa os campos
	vVetor = Split(strDecriptografado, "|", -1)
	if LBound(vVetor)=UBound(vVetor) then exit function
	'Atualiza data/hora da �ltima atividade
	vVetor(Ubound(vVetor)) = CStr(CDbl(Now))
	strDecriptografado = Join(vVetor, "|")
	AtualizaSessionCtrlInfoDataHoraUltAtividade = CriptografaTexto(strDecriptografado, strChave)
end function



' ------------------------------------------------------------------------
'   MONTA CAMPO FORM SESSION CTRL INFO
'   Cria um campo input para armazenar o valor de SessionCtrlInfo
'   (tratamento para sess�o expirada).
function MontaCampoFormSessionCtrlInfo(ByVal strSessionCtrlInfo)
	strSessionCtrlInfo = AtualizaSessionCtrlInfoDataHoraUltAtividade(strSessionCtrlInfo)
	MontaCampoFormSessionCtrlInfo = "<INPUT type=HIDDEN name='SessionCtrlInfo' value='" & strSessionCtrlInfo & "'>"
end function



' ------------------------------------------------------------------------
'   MONTA CAMPO QUERYSTRING SESSION CTRL INFO
'   Cria um campo QueryString para ser passado pela URL informando o 
'   valor de SessionCtrlInfo (tratamento para sess�o expirada).
function MontaCampoQueryStringSessionCtrlInfo(ByVal strSessionCtrlInfo)
	strSessionCtrlInfo = AtualizaSessionCtrlInfoDataHoraUltAtividade(strSessionCtrlInfo)
	MontaCampoQueryStringSessionCtrlInfo = "SessionCtrlInfo=" & strSessionCtrlInfo
end function


' ------------------------------------------------------------------------
'   MontaNumPedidoExibicaoTitleBrowser
function MontaNumPedidoExibicaoTitleBrowser(ByVal strNumPedido)
	MontaNumPedidoExibicaoTitleBrowser = " - Pedido N� " & strNumPedido
end function


' ------------------------------------------------------------------------
'   MontaNumOrcamentoExibicaoTitleBrowser
function MontaNumOrcamentoExibicaoTitleBrowser(ByVal strNumOrcamento)
	MontaNumOrcamentoExibicaoTitleBrowser = " - Or�amento N� " & strNumOrcamento
end function


' ------------------------------------------------------------------------
'   MontaNumPrePedidoExibicaoTitleBrowser
function MontaNumPrePedidoExibicaoTitleBrowser(ByVal strNumOrcamento)
	MontaNumPrePedidoExibicaoTitleBrowser = " Pr�-Pedido N� " & strNumOrcamento
end function


' ------------------------------------------------------------------------
'   QuotedStr
'   Retorna um texto tratado para inserir aspas simples em comandos SQL.
function QuotedStr(ByVal strTexto)
dim intCounter
dim strResp
dim strChar
	strResp = ""
	for intCounter=1 to Len(strTexto)
		strChar = Mid(strTexto, intCounter, 1)
		if strChar = "'" then 
			strResp = strResp & String(2, strChar)
		else
			strResp = strResp & strChar
			end if
		next
	QuotedStr = strResp
end function


' ------------------------------------------------------------------------
'   JsQuotedStr
'   Retorna um texto tratado para exibir aspas simples em strings
'	no JavaScript.
function JsQuotedStr(ByVal strTexto)
dim intCounter
dim strResp
dim strChar
	strResp = ""
	for intCounter=1 to Len(strTexto)
		strChar = Mid(strTexto, intCounter, 1)
		if strChar = "'" then 
			strResp = strResp & "\'"
		else
			strResp = strResp & strChar
			end if
		next
	JsQuotedStr = strResp
end function


' ------------------------------------------------------------------------
'   Monta o header com a identifica��o do pedido que � exibido no topo da
'	tela do pedido.
function MontaHeaderIdentificacaoPedido(ByVal strNumPedido, ByVal r_pedido, ByVal intLargTable)
dim strResp
dim strStEntregaData
dim strClassData
dim strClassCell
dim strRecebidoData
dim loja
dim s_colon

    loja =  Trim(Session("loja_atual"))
	MontaHeaderIdentificacaoPedido = ""
	
	strClassData="STP"
	strClassCell=""
	strStEntregaData=""
	strRecebidoData=""
	if r_pedido.st_entrega=ST_ENTREGA_ENTREGUE then 
		if Cstr(r_pedido.PedidoRecebidoStatus) = Cstr(COD_ST_PEDIDO_RECEBIDO_SIM) then
			strClassData="HoraPed"
			strClassCell=" class='MD ME' style='border-width:2px;border-color:" & x_status_entrega_cor(r_pedido.st_entrega, strNumPedido) & ";' "
			strStEntregaData = formata_data(r_pedido.entregue_data)
			strRecebidoData = "<br>" & chr(13) & _
								"<span class='" & strClassData & "' style='color:#9932CC;text-align:center;margin:0px 3px 0px 3px;'>" & _
								formata_data(r_pedido.PedidoRecebidoData) & _
								"</span>" & chr(13)
		else
			strStEntregaData="  (" & formata_data(r_pedido.entregue_data) & ")"
			end if
	elseif r_pedido.st_entrega=ST_ENTREGA_CANCELADO then
		strStEntregaData="  (" & formata_data(r_pedido.cancelado_data) & ")"
		end if

	strResp = _
		"<table width='" & Cstr(intLargTable) & "' cellpadding='0' cellspacing='0' style='border-bottom:1px solid black;'>" & chr(13)
    
    if (r_pedido.pedido_bs_x_marketplace <> "") then
        if r_pedido.pedido_bs_x_marketplace <> "" then s_colon = ":" else s_colon = ""
		strResp = strResp & _
		    "	<tr>" & chr(13) & _
		    "		<td valign='bottom' align='right' width='100%' colspan='3' style='padding-top: 4px;padding-bottom:0'>" & chr(13) & _
            "           <span class='C' style='font-size:10pt'>" & obtem_descricao_tabela_t_codigo_descricao("PedidoECommerce_Origem", r_pedido.marketplace_codigo_origem) & s_colon & "</span><span class='C' style='font-size:12pt'>" & r_pedido.pedido_bs_x_marketplace & "</span>" & chr(13) & _
            "       </td>" & chr(13) & _
            "   </tr>" & chr(13)
    end if

    if (r_pedido.pedido_ac <> "") then
		if r_pedido.pedido_ac <> "" then s_colon = ":" else s_colon = ""
        strResp = strResp & _
		    "	<tr>" & chr(13) & _
		    "		<td valign='bottom' align='right' width='100%' colspan='3' style='padding-top: 4px;padding-bottom:0'>" & chr(13) & _
            "           <span class='C' style='font-size:10pt;color:purple'>Magento" & s_colon & "</span><span class='C' style='font-size:12pt;color:purple'>" & r_pedido.pedido_ac & "</span>" & chr(13) & _
            "       </td>" & chr(13) & _
            "   </tr>" & chr(13)
    
    end if
    strResp = strResp & _
        "   <tr>" & chr(13) & _
		"		<td valign='bottom' align='left' width='33%' style='padding-bottom: 4px'>" & chr(13) & _
		"			<table cellpadding='0' cellspacing='0'>" & chr(13) & _
		"				<tr>" & chr(13) & _
		"					<td valign='bottom' align='left'>" & chr(13) & _
		"						<span class='STP' style='color:" & x_status_entrega_cor(r_pedido.st_entrega, strNumPedido) & ";'>" & _
									Ucase(x_status_entrega(r_pedido.st_entrega)) & chr(13) & _
		"						</span>" & chr(13) & _
		"					</td>" & chr(13) & _
		"					<td align='left' style='width:6px;padding-bottom:4px'>" & chr(13) & _
		"					</td>" & chr(13) & _
		"					<td align='center' valign='bottom'" & strClassCell & "  style='padding-bottom: 4px'>" & chr(13) & _
		"						<span class='" & strClassData & "' style='text-align:center;color:" & x_status_entrega_cor(r_pedido.st_entrega, strNumPedido) & ";margin:0px 3px 0px 3px;'" & _
								">" & strStEntregaData & "</span>" & chr(13) & _
								strRecebidoData & _
		"					</td>" & chr(13) & _
		"				</tr>" & chr(13) & _
		"			</table>" & chr(13) & _
		"		</td>" & chr(13) & _
		"		<td valign='bottom' align='center' width='33%'>" & chr(13) & _
		"			<span class='STP'>" & formata_data(r_pedido.data) & chr(13) & _
		"				<span class='HoraPed'>" & formata_hhnnss_para_hh_nn(r_pedido.hora) & "</span>" & chr(13) & _
		"			</span>" & chr(13) & _
		"		</td>" & chr(13) & _
		"		<td valign='bottom' align='right' nowrap>" & chr(13) & _
		"			<span class='PEDIDO'>Pedido N�&nbsp;" & strNumPedido & "</span>" & chr(13)& _
		"		</td>" & chr(13) & _
		"	</tr>" & chr(13) & _
		"</table>" & chr(13)
		
	MontaHeaderIdentificacaoPedido = strResp
	
end function


' ------------------------------------------------------------------------
'   Monta o header com a identifica��o do or�amento que � exibido no
'	topo da tela do or�amento.
function MontaHeaderIdentificacaoOrcamento(ByVal strNumOrcamento, ByVal r_orcamento, ByVal intLargTable)
dim strResp
dim strData
	
	MontaHeaderIdentificacaoOrcamento = ""

	strResp = strResp & _
		"<table width='649' cellpadding='4' cellspacing='0' style='border-bottom:1px solid black'>" & chr(13) & _
		"	<tr>" & chr(13)
		
	if r_orcamento.st_orcamento <> "" then
		strResp = strResp & _
			"		<td valign='bottom' align='left'>" & chr(13) & _
			"			<span class='STP' style='color:" & x_st_orcamento_cor(r_orcamento.st_orcamento) & ";'>" & chr(13) & _
							Ucase(x_st_orcamento(r_orcamento.st_orcamento))

		strData = ""
		if Cstr(r_orcamento.st_orcamento)=Cstr(ST_ORCAMENTO_CANCELADO) then strData=formata_data(r_orcamento.cancelado_data)
		if strData <> "" then strData = "  (" & strData & ")"
		strResp = strResp & strData
		
		strResp = strResp & _
			"			</span>" & chr(13) & _
			"		</td>" & chr(13) & _
			"		<td valign='bottom' align='center' style='width:120px;'>" & chr(13) & _
			"			<span class='STP'>" & chr(13) & _
							formata_data(r_orcamento.data) & _
			"				<span class='HoraPed'>" & chr(13) & _
								formata_hhnnss_para_hh_nn(r_orcamento.hora) & chr(13) & _
			"				</span>" & chr(13) & _
			"			</span>" & chr(13) & _
			"		</td>" & chr(13)
	else
		strResp = strResp & _
			"		<td align='left' valign='bottom' style='width:120px;'>" & chr(13) & _
			"			<span class='STP'>" & chr(13) & _
							formata_data(r_orcamento.data) & _
			"				<span class='HoraPed'>" & chr(13) & _
								formata_hhnnss_para_hh_nn(r_orcamento.hora) & _
			"				</span>" & chr(13) & _
			"			</span>" & chr(13) & _
			"		</td>" & chr(13)
		end if

	if r_orcamento.st_orc_virou_pedido = 1 then
		strResp = strResp & _
			"		<td align='left' valign='bottom' nowrap>" & chr(13) & _
			"			<span class='STP' style='color:red;'>" & chr(13) & _
			"				<a href='javascript:fPEDConcluir(" & chr(34) & r_orcamento.pedido & chr(34) & ")' title='clique para consultar o pedido' style='color:red;'>" & chr(13) & _
								"Pedido&nbsp;&nbsp;" & r_orcamento.pedido & "&nbsp;&nbsp;(" & formata_data(r_pedido.data) & ")</a>" & chr(13) & _
			"			</span>" & chr(13) & _
			"		</td>" & chr(13)
		end if

	strResp = strResp & _
		"		<td valign='bottom' align='right' nowrap>" & chr(13) & _
		"			<span class='PEDIDO'>Or�amento N�&nbsp;" & strNumOrcamento & "</span>" & chr(13) & _
		"		</td>" & chr(13) & _
		"	</tr>" & chr(13) & _
		"</table>" & chr(13)

	MontaHeaderIdentificacaoOrcamento = strResp
end function


' ------------------------------------------------------------------------
'   Monta o header com a identifica��o do or�amento que � exibido no
'	topo da tela do or�amento.
function MontaHeaderIdentificacaoPrePedido(ByVal strNumOrcamento, ByVal r_orcamento, ByVal intLargTable)
dim strResp
dim strData
	
	MontaHeaderIdentificacaoPrePedido = ""

	strResp = strResp & _
		"<table width='649' cellpadding='4' cellspacing='0' style='border-bottom:1px solid black'>" & chr(13) & _
		"	<tr>" & chr(13)
		
	if r_orcamento.st_orcamento <> "" then
		strResp = strResp & _
			"		<td valign='bottom' align='left'>" & chr(13) & _
			"			<span class='STP' style='color:" & x_st_orcamento_cor(r_orcamento.st_orcamento) & ";'>" & chr(13) & _
							Ucase(x_st_orcamento(r_orcamento.st_orcamento))

		strData = ""
		if Cstr(r_orcamento.st_orcamento)=Cstr(ST_ORCAMENTO_CANCELADO) then strData=formata_data(r_orcamento.cancelado_data)
		if strData <> "" then strData = "  (" & strData & ")"
		strResp = strResp & strData
		
		strResp = strResp & _
			"			</span>" & chr(13) & _
			"		</td>" & chr(13) & _
			"		<td valign='bottom' align='center' style='width:120px;'>" & chr(13) & _
			"			<span class='STP'>" & chr(13) & _
							formata_data(r_orcamento.data) & _
			"				<span class='HoraPed'>" & chr(13) & _
								formata_hhnnss_para_hh_nn(r_orcamento.hora) & chr(13) & _
			"				</span>" & chr(13) & _
			"			</span>" & chr(13) & _
			"		</td>" & chr(13)
	else
		strResp = strResp & _
			"		<td align='left' valign='bottom' style='width:120px;'>" & chr(13) & _
			"			<span class='STP'>" & chr(13) & _
							formata_data(r_orcamento.data) & _
			"				<span class='HoraPed'>" & chr(13) & _
								formata_hhnnss_para_hh_nn(r_orcamento.hora) & _
			"				</span>" & chr(13) & _
			"			</span>" & chr(13) & _
			"		</td>" & chr(13)
		end if

	if r_orcamento.st_orc_virou_pedido = 1 then
		strResp = strResp & _
			"		<td align='left' valign='bottom' nowrap>" & chr(13) & _
			"			<span class='STP' style='color:red;'>" & chr(13) & _
			"				<a href='javascript:fPEDConcluir(" & chr(34) & r_orcamento.pedido & chr(34) & ")' title='clique para consultar o pedido' style='color:red;'>" & chr(13) & _
								"Pedido&nbsp;&nbsp;" & r_orcamento.pedido & "&nbsp;&nbsp;(" & formata_data(r_pedido.data) & ")</a>" & chr(13) & _
			"			</span>" & chr(13) & _
			"		</td>" & chr(13)
		end if

	strResp = strResp & _
		"		<td valign='bottom' align='right' nowrap>" & chr(13) & _
		"			<span class='PEDIDO'>Pr�-Pedido N�&nbsp;" & strNumOrcamento & "</span>" & chr(13) & _
		"		</td>" & chr(13) & _
		"	</tr>" & chr(13) & _
		"</table>" & chr(13)

	MontaHeaderIdentificacaoPrePedido = strResp
end function


' ------------------------------------------------------------------------
'   A carga da planilha Excel de produtos permite formatar a descri��o
'	com negrito, it�lico e sublinhado.
'	Essas formata��es s�o armazenadas no BD com as tags HTML:
'		Negrito: <b></b>
'		It�lico: <i></i>
'		Sublinhado: <u></u>
'	Entretanto, como em v�rios locais a descri��o toda j� � exibida em
'	negrito, pode-se ressaltar de outras formas, como colocar o texto
'	em tamanho maior, ou em outra cor, etc.
function produto_formata_descricao_em_html(ByVal strDescricaoHtml)
dim strReplaceAbreBold
dim strReplaceFechaBold
	strReplaceAbreBold="<span style=" & chr(34) & "font-size:130%;font-weight:bolder;" & chr(34) & ">"
	strReplaceFechaBold="</span>"
	strDescricaoHtml = Replace(strDescricaoHtml, "<b>", strReplaceAbreBold)
	strDescricaoHtml = Replace(strDescricaoHtml, "<B>", strReplaceAbreBold)
	strDescricaoHtml = Replace(strDescricaoHtml, "</b>", strReplaceFechaBold)
	strDescricaoHtml = Replace(strDescricaoHtml, "</B>", strReplaceFechaBold)
	produto_formata_descricao_em_html = strDescricaoHtml
end function

function DecodificaCorHtmlValorMonetario(vl)
dim strCor
	strCor = "black"
	if IsNumeric(vl) then
		if CCur(vl) > 0 then
			strCor = "green"
		elseif CCur(vl) < 0 then
			strCor = "red"
			end if
		end if
	DecodificaCorHtmlValorMonetario=strCor
end function

function descricaoCustoFinancFornecTipoParcelamento(ByVal strCodigoTipoParcelamento)
dim strResp
	strResp = ""
	strCodigoTipoParcelamento = Trim("" & strCodigoTipoParcelamento)
	if strCodigoTipoParcelamento = COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__COM_ENTRADA then
		strResp = "Com Entrada"
	elseif strCodigoTipoParcelamento = COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__SEM_ENTRADA then
		strResp = "Sem Entrada"
	elseif strCodigoTipoParcelamento = COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__A_VISTA then
		strResp = "� Vista"
		end if
	descricaoCustoFinancFornecTipoParcelamento = strResp
end function

function decodificaCustoFinancFornecQtdeParcelas(ByVal strCodigoTipoParcelamento, ByVal strQtdeParcelas)
dim strResp	
	strResp = ""
	strQtdeParcelas = Trim("" & strQtdeParcelas)
	if strCodigoTipoParcelamento = COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__SEM_ENTRADA then
		strResp = "0+" & strQtdeParcelas
	elseif strCodigoTipoParcelamento = COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__COM_ENTRADA then
		strResp = "1+" & strQtdeParcelas
		end if
	decodificaCustoFinancFornecQtdeParcelas=strResp
end function


' ------------------------------------------------------------------------
'   finNaturezaDescricao
'   Retorna a descri��o da natureza da opera��o a partir do c�digo.
function finNaturezaDescricao(byval codigo)
dim strResp
	strResp=""
	codigo = Trim("" & codigo)
	if Cstr(codigo) = Cstr(COD_FIN_NATUREZA__CREDITO) then
		strResp = "Cr�dito"
	elseif Cstr(codigo) = Cstr(COD_FIN_NATUREZA__DEBITO) then
		strResp = "D�bito"
		end if
	finNaturezaDescricao=strResp
end function


' ------------------------------------------------------------------------
'   finNaturezaCor
'   Retorna a cor de exibi��o para a natureza da opera��o a 
'   partir do c�digo.
function finNaturezaCor(byval codigo)
dim strResp
	strResp="#000000"
	codigo = Trim("" & codigo)
	if Cstr(codigo) = Cstr(COD_FIN_NATUREZA__CREDITO) then
		strResp = "#006600"
	elseif Cstr(codigo) = Cstr(COD_FIN_NATUREZA__DEBITO) then
		strResp = "#FF0000"
		end if
	finNaturezaCor=strResp
end function


' ------------------------------------------------------------------------
'   finStAtivoDescricao
'   Retorna a descri��o do status a partir do c�digo.
function finStAtivoDescricao(byval codigo)
dim strResp
	strResp=""
	codigo = Trim("" & codigo)
	if Cstr(codigo) = Cstr(COD_FIN_ST_ATIVO__INATIVO) then
		strResp = "Inativo"
	elseif Cstr(codigo) = Cstr(COD_FIN_ST_ATIVO__ATIVO) then
		strResp = "Ativo"
		end if
	finStAtivoDescricao=strResp
end function


' ------------------------------------------------------------------------
'   finStAtivoCor
'   Retorna a cor de exibi��o para o status a partir do c�digo.
function finStAtivoCor(byval codigo)
dim strResp
	strResp="#000000"
	codigo = Trim("" & codigo)
	if Cstr(codigo) = Cstr(COD_FIN_ST_ATIVO__ATIVO) then
		strResp = "#006600"
	elseif Cstr(codigo) = Cstr(COD_FIN_ST_ATIVO__INATIVO) then
		strResp = "#FF0000"
		end if
	finStAtivoCor=strResp
end function


' ------------------------------------------------------------------------
'   isInscricaoEstadualValida
'   Indica se o n�mero de inscri��o estadual � v�lido
function isInscricaoEstadualValida(byval inscricaoEstadual, byval uf)
dim strInscricaoEstadualNormalizado
dim objIE
dim blnResultado
dim blnOk
dim i
dim c
dim qtdeDigitos

	isInscricaoEstadualValida=False
	
	inscricaoEstadual=ucase(trim(inscricaoEstadual))
	uf=ucase(trim(uf))
	
	if inscricaoEstadual="ISENTO" then
		strInscricaoEstadualNormalizado = inscricaoEstadual
	else
	'	VERIFICA SE H� CARACTERES INV�LIDOS
		blnOk=True
		qtdeDigitos=0
		for i=1 to len(inscricaoEstadual)
			c=mid(inscricaoEstadual, i, 1)
			if (Not IsDigit(c)) And (c<>".") And (c<>"-") And (c<>"/") then
				blnOk=false
				exit for
				end if
			if IsDigit(c) then qtdeDigitos=qtdeDigitos+1
			next
			
		if not blnOk then exit function
		if qtdeDigitos < 2 then exit function
		if qtdeDigitos > 14 then exit function
	
	'	INFORMAR SOMENTE D�GITOS PARA A DLL DE VALIDA��O DO SINTEGRA
		strInscricaoEstadualNormalizado = retorna_so_digitos(inscricaoEstadual)
		end if
		
	set objIE = CreateObject("ComPlusWrapper_DllInscE32.ComPlusWrapper_DllInscE32")
	blnResultado = objIE.isInscricaoEstadualOk(strInscricaoEstadualNormalizado, uf)
	set objIE = Nothing

	isInscricaoEstadualValida = blnResultado
end function


' ------------------------------------------------------------------------
'   isTextoValido
'   Indica se o texto possui somente caracteres v�lidos
function isTextoValido(byval texto, byref caracteresInvalidos)
dim i
dim c
dim u
dim blnErro

	isTextoValido=False
	caracteresInvalidos=""
	
	blnErro=False
	for i=1 to len(texto)
		c=mid(texto,i,1)
		if (Asc(c) < 32) Or (Asc(c) > 127) then
			u=ucase(c)
			if u="�" then
			'	NOP
			elseif u="�" then
			'	NOP
			elseif u="�" then
			'	NOP
			elseif u="�" then
			'	NOP
			elseif u="�" then
			'	NOP
			elseif u="�" then
			'	NOP
			elseif u="�" then
			'	NOP
			elseif u="�" then
			'	NOP
			elseif u="�" then
			'	NOP
			elseif u="�" then
			'	NOP
			elseif u="�" then
			'	NOP
			elseif u="�" then
			'	NOP
			elseif u="�" then
			'	NOP
			elseif u="�" then
			'	NOP
			elseif u="�" then
			'	NOP
			elseif u="�" then
			'	NOP
			elseif u="�" then
			'	NOP
			elseif u="�" then
			'	NOP
			elseif u="�" then
			'	NOP
			elseif u="�" then
			'	NOP
			elseif u="�" then
			'	NOP
			elseif u="�" then
			'	NOP
			elseif u="�" then
			'	NOP
			else
				blnErro=True
				if caracteresInvalidos <> "" then caracteresInvalidos = caracteresInvalidos & " "
				caracteresInvalidos = caracteresInvalidos & c
				end if
			end if
		next
		
	if blnErro then exit function
	
	isTextoValido=True
end function



' ------------------------------------------------------------------------
'   RETIRA ACENTUACAO
'   Retira a acentua��o do texto
function retira_acentuacao(byval texto)
dim s_resp
dim i
dim c

	s_resp=""
	
	for i=1 to len(texto)
		c = mid(texto,i,1)
		if (Asc(c) < 32) Or (Asc(c) > 127) then
			if (c="�")Or(c="�")Or(c="�")Or(c="�")Or(c="�") then
				c="A"
			elseif (c="�")Or(c="�")Or(c="�")Or(c="�")Or(c="�") then
				c="a"
			elseif (c="�")Or(c="�")Or(c="�")Or(c="�") then
				c="E"
			elseif (c="�")Or(c="�")Or(c="�")Or(c="�") then
				c="e"
			elseif (c="�")Or(c="�")Or(c="�")Or(c="�") then
				c="I"
			elseif (c="�")Or(c="�")Or(c="�")Or(c="�") then
				c="i"
			elseif (c="�")Or(c="�")Or(c="�")Or(c="�")Or(c="�") then
				c="O"
			elseif (c="�")Or(c="�")Or(c="�")Or(c="�")Or(c="�") then
				c="o"
			elseif (c="�")Or(c="�")Or(c="�")Or(c="�") then
				c="U"
			elseif (c="�")Or(c="�")Or(c="�")Or(c="�") then
				c="u"
			elseif c="�" then
				c="C"
			elseif c="�" then
				c="c"
			elseif c="�" then
				c="N"
			elseif c="�" then
				c="n"
			elseif c="�" then
				c="y"
				end if
			end if
		
		s_resp = s_resp & c
		next

	retira_acentuacao = s_resp
end function



' ------------------------------------------------------------------------
'   NIVEL_ACESSO_BLOCO_NOTAS_PEDIDO_DESCRICAO
'   Retorna a descri��o do n�vel de acesso ao bloco de notas do pedido 
'	a partir do c�digo.
function nivel_acesso_bloco_notas_pedido_descricao(byval codigo)
dim strResp
	strResp=""
	codigo = Trim("" & codigo)
	if Cstr(codigo) = Cstr(COD_NIVEL_ACESSO_BLOCO_NOTAS_PEDIDO__PUBLICO) then
		strResp = "P�blico"
	elseif Cstr(codigo) = Cstr(COD_NIVEL_ACESSO_BLOCO_NOTAS_PEDIDO__RESTRITO) then
		strResp = "Restrito"
	elseif Cstr(codigo) = Cstr(COD_NIVEL_ACESSO_BLOCO_NOTAS_PEDIDO__SIGILOSO) then
		strResp = "Sigiloso"
		end if
	nivel_acesso_bloco_notas_pedido_descricao=strResp
end function


' ------------------------------------------------------------------------
'   NIVEL_ACESSO_CHAMADO_PEDIDO_DESCRICAO
'   Retorna a descri��o do n�vel de acesso ao chamado do pedido 
'	a partir do c�digo.
function nivel_acesso_chamado_pedido_descricao(byval codigo)
dim strResp
	strResp=""
	codigo = Trim("" & codigo)
	if Cstr(codigo) = Cstr(COD_NIVEL_ACESSO_CHAMADO_PEDIDO__PUBLICO) then
		strResp = "P�blico"
	elseif Cstr(codigo) = Cstr(COD_NIVEL_ACESSO_CHAMADO_PEDIDO__PUBLICO_INTERNO) then
		strResp = "P�blico (interno)"
    elseif Cstr(codigo) = Cstr(COD_NIVEL_ACESSO_CHAMADO_PEDIDO__RESTRITO) then
		strResp = "Restrito"
	elseif Cstr(codigo) = Cstr(COD_NIVEL_ACESSO_CHAMADO_PEDIDO__SIGILOSO) then
		strResp = "Sigiloso"
		end if
	nivel_acesso_chamado_pedido_descricao=strResp
end function


' ------------------------------------------------------------------------
'   NIVEL_ACESSO_BLOCO_NOTAS_PEDIDO_COR
'   Retorna a cor para exibi��o da descri��o do n�vel de acesso ao bloco
'	de notas do pedido.
function nivel_acesso_bloco_notas_pedido_cor(byval codigo)
dim strResp
	strResp="#000000"
	codigo = Trim("" & codigo)
	if Cstr(codigo) = Cstr(COD_NIVEL_ACESSO_BLOCO_NOTAS_PEDIDO__PUBLICO) then
		strResp = "#006400"
	elseif Cstr(codigo) = Cstr(COD_NIVEL_ACESSO_BLOCO_NOTAS_PEDIDO__RESTRITO) then
		strResp = "#FF8C00"
	elseif Cstr(codigo) = Cstr(COD_NIVEL_ACESSO_BLOCO_NOTAS_PEDIDO__SIGILOSO) then
		strResp = "#FF0000"
		end if
	nivel_acesso_bloco_notas_pedido_cor=strResp
end function


' ------------------------------------------------------------------------
'   OBTEM PATH PDF DANFE
'   Retorna o path do diret�rio onde est�o os arquivos PDF das DANFE's
'	do referido emitente.
function obtem_path_pdf_danfe(byval idNFeEmitente)
dim s_resp, intIdNFeEmitente

	obtem_path_pdf_danfe = ""
	
	if Not IsNumeric(idNFeEmitente) then exit function
	intIdNFeEmitente= CInt(idNFeEmitente)
	
	if intIdNFeEmitente = ID_NFE_EMITENTE__OLD01_01 then
		s_resp = DIR_TARGET_ONE_PDF_DANFE_EMITENTE__OLD01_01
	elseif intIdNFeEmitente = ID_NFE_EMITENTE__OLD03_01 then
		s_resp = DIR_TARGET_ONE_PDF_DANFE_EMITENTE__OLD03_01
	elseif intIdNFeEmitente = ID_NFE_EMITENTE__OLD02_02 then
		s_resp = DIR_TARGET_ONE_PDF_DANFE_EMITENTE__OLD02_02
	elseif intIdNFeEmitente = ID_NFE_EMITENTE__DIS_01 then
		s_resp = DIR_TARGET_ONE_PDF_DANFE_EMITENTE__DIS_01
	elseif intIdNFeEmitente = ID_NFE_EMITENTE__DIS_03 then
		s_resp = DIR_TARGET_ONE_PDF_DANFE_EMITENTE__DIS_03
	elseif intIdNFeEmitente = ID_NFE_EMITENTE__DIS_903 then
		s_resp = DIR_TARGET_ONE_PDF_DANFE_EMITENTE__DIS_903
	else
		s_resp = ""
		end if
		
	obtem_path_pdf_danfe=s_resp
end function



' ------------------------------------------------------------------------
'   OBTEM PATH XML NFE
'   Retorna o path do diret�rio onde est�o os arquivos XML das NFe's
'	do referido emitente.
function obtem_path_xml_nfe(byval idNFeEmitente)
dim s_resp, intIdNFeEmitente

	obtem_path_xml_nfe = ""
	
	if Not IsNumeric(idNFeEmitente) then exit function
	intIdNFeEmitente= CInt(idNFeEmitente)
	
	if intIdNFeEmitente = ID_NFE_EMITENTE__OLD01_01 then
		s_resp = DIR_TARGET_ONE_XML_NFE_EMITENTE__OLD01_01
	elseif intIdNFeEmitente = ID_NFE_EMITENTE__OLD03_01 then
		s_resp = DIR_TARGET_ONE_XML_NFE_EMITENTE__OLD03_01
	elseif intIdNFeEmitente = ID_NFE_EMITENTE__OLD02_02 then
		s_resp = DIR_TARGET_ONE_XML_NFE_EMITENTE__OLD02_02
	elseif intIdNFeEmitente = ID_NFE_EMITENTE__DIS_01 then
		s_resp = DIR_TARGET_ONE_XML_NFE_EMITENTE__DIS_01
	elseif intIdNFeEmitente = ID_NFE_EMITENTE__DIS_03 then
		s_resp = DIR_TARGET_ONE_XML_NFE_EMITENTE__DIS_03
	elseif intIdNFeEmitente = ID_NFE_EMITENTE__DIS_903 then
		s_resp = DIR_TARGET_ONE_XML_NFE_EMITENTE__DIS_903
	else
		s_resp = ""
		end if
		
	obtem_path_xml_nfe=s_resp
end function



' ------------------------------------------------------------------------
'   MONTA DESCRICAO FORMA PAGTO
'   Monta a descri��o para a forma de pagamento especificada.
'	O par�metro deve ser um recordset contendo os campos que armazenam os
'	dados da forma de pagamento.
function monta_descricao_forma_pagto(byref r)
dim strResp

	strResp = ""
	
	if Trim("" & r("tipo_parcelamento")) = COD_FORMA_PAGTO_A_VISTA then
		strResp = "� Vista  (" & x_opcao_forma_pagamento(r("av_forma_pagto")) & ")"
	elseif Trim("" & r("tipo_parcelamento")) = COD_FORMA_PAGTO_PARCELA_UNICA then
		strResp = "Parcela �nica:  " & SIMBOLO_MONETARIO & " " & formata_moeda(r("pu_valor")) & "  (" & x_opcao_forma_pagamento(r("pu_forma_pagto")) & ")  vencendo ap�s " & Cstr(r("pu_vencto_apos")) & " dias"
	elseif Trim("" & r("tipo_parcelamento")) = COD_FORMA_PAGTO_PARCELADO_CARTAO then
		strResp = "Parcelado no Cart�o (internet) em " & Cstr(r("pc_qtde_parcelas")) & " x  " & SIMBOLO_MONETARIO & " " & formata_moeda(r("pc_valor_parcela"))
	elseif Trim("" & r("tipo_parcelamento")) = COD_FORMA_PAGTO_PARCELADO_CARTAO_MAQUINETA then
		strResp = "Parcelado no Cart�o (maquineta) em " & Cstr(r("pc_maquineta_qtde_parcelas")) & " x  " & SIMBOLO_MONETARIO & " " & formata_moeda(r("pc_maquineta_valor_parcela"))
	elseif Trim("" & r("tipo_parcelamento")) = COD_FORMA_PAGTO_PARCELADO_COM_ENTRADA then
		strResp = "Entrada:  " & SIMBOLO_MONETARIO & " " & formata_moeda(r("pce_entrada_valor")) & "  (" & x_opcao_forma_pagamento(r("pce_forma_pagto_entrada")) & ")" & _
				  chr(13) & "<br>" & chr(13) & _
				  "Presta��es:  " & Cstr(r("pce_prestacao_qtde")) & " x  " & SIMBOLO_MONETARIO & " " & formata_moeda(r("pce_prestacao_valor")) & _
				  "  (" & x_opcao_forma_pagamento(r("pce_forma_pagto_prestacao")) & ")  vencendo a cada " & _
				  Cstr(r("pce_prestacao_periodo")) & " dias"
	elseif Trim("" & r("tipo_parcelamento")) = COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA then
		strResp = "1� Presta��o:  " & SIMBOLO_MONETARIO & " " & formata_moeda(r("pse_prim_prest_valor")) & "  (" & x_opcao_forma_pagamento(r("pse_forma_pagto_prim_prest")) & ")  vencendo ap�s " & Cstr(r("pse_prim_prest_apos")) & " dias" & _
				  chr(13) & "<br>" & chr(13) & _
				  "Demais Presta��es:  " & Cstr(r("pse_demais_prest_qtde")) & " x  " & SIMBOLO_MONETARIO & " " & formata_moeda(r("pse_demais_prest_valor")) & _
				  "  (" & x_opcao_forma_pagamento(r("pse_forma_pagto_demais_prest")) & ")  vencendo a cada " & _
				  Cstr(r("pse_demais_prest_periodo")) & " dias"
		end if
		
	monta_descricao_forma_pagto = strResp
end function


' ------------------------------------------------------------------------
'	IS FORMA PAGTO SOMENTE CARTAO
'	Analisa e indica se a forma de pagamento utiliza somente o cart�o
'	como meio de pagamento para todas as parcelas.
'	Retorno:
'		True = o cart�o � o meio de pagamento p/ todas as parcelas
'		False = o cart�o N�O � o �nico meio de pagamento utilizado
function is_forma_pagto_somente_cartao(byref r)
dim blnResp
	blnResp = False
	if Trim("" & r("tipo_parcelamento")) = COD_FORMA_PAGTO_A_VISTA then
		if Trim("" & r("av_forma_pagto")) = ID_FORMA_PAGTO_CARTAO then blnResp = True
	elseif Trim("" & r("tipo_parcelamento")) = COD_FORMA_PAGTO_PARCELA_UNICA then
		if Trim("" & r("pu_forma_pagto")) = ID_FORMA_PAGTO_CARTAO then blnResp = True
	elseif Trim("" & r("tipo_parcelamento")) = COD_FORMA_PAGTO_PARCELADO_CARTAO then
		blnResp = True
	elseif Trim("" & r("tipo_parcelamento")) = COD_FORMA_PAGTO_PARCELADO_CARTAO_MAQUINETA then
		'NOP
	elseif Trim("" & r("tipo_parcelamento")) = COD_FORMA_PAGTO_PARCELADO_COM_ENTRADA then
		if (Trim("" & r("pce_forma_pagto_entrada")) = ID_FORMA_PAGTO_CARTAO) And (Trim("" & r("pce_forma_pagto_prestacao")) = ID_FORMA_PAGTO_CARTAO) then blnResp = True
	elseif Trim("" & r("tipo_parcelamento")) = COD_FORMA_PAGTO_PARCELADO_SEM_ENTRADA then
		if (Trim("" & r("pse_forma_pagto_prim_prest")) = ID_FORMA_PAGTO_CARTAO) And (Trim("" & r("pse_forma_pagto_demais_prest")) = ID_FORMA_PAGTO_CARTAO) then blnResp = True
		end if
	is_forma_pagto_somente_cartao = blnResp
end function


' ------------------------------------------------------------------------
'   IS PLACA VEICULO OK
'   Indica se a placa do ve�culo est� em formato v�lido
function isPlacaVeiculoOk(byval numeroPlaca)
dim i, c, letras, numeros
	
	isPlacaVeiculoOk = False
	
	numeroPlaca = Trim("" & numeroPlaca)
	if Len(numeroPlaca) = 0 then exit function
	
	letras = ""
	numeros = ""
	for i=1 to Len(numeroPlaca)
		c = Mid(numeroPlaca, i, 1)
		if c = " " then
		'	O ESPA�O EM BRANCO APARECEU EM POSI��O INESPERADA?
			if Len(letras) <> 3 then exit function
			if Len(numeros) > 0 then exit function
		elseif isLetra(c) then
		'	APARECEU UMA LETRA DEPOIS DE J� TER INICIADO A PARTE DOS D�GITOS?
			if Len(numeros) > 0 then exit function
			letras = letras + c
		elseif isDigit(c) then
		'	APARECEU UM D�GITO EM POSI��O INESPERADA?
			if Len(letras) <> 3 then exit function
			numeros = numeros + c
		else
		'	CARACTER INV�LIDO!
			exit function
			end if
		next
	
	if Len(letras) <> 3 then exit function
	if Len(numeros) <> 4 then exit function
	
	isPlacaVeiculoOk = True
end function



' ------------------------------------------------------------------------
'   xml_read_node
function xml_read_node(Byval node_path, Byref blnNodeNotFound)
dim oNode
	blnNodeNotFound = False

	set oNode = objXML.documentElement.selectSingleNode(node_path)
	if oNode is nothing then
		blnNodeNotFound = True
		xml_read_node = ""
		exit function
		end if
	
	xml_read_node = oNode.text
end function



' ------------------------------------------------------------------------
'   cria_instancia_cl_CIELO_REQUISICAO_TRANSACAO_TX
function cria_instancia_cl_CIELO_REQUISICAO_TRANSACAO_TX(Byval bandeira)
dim trx
	set trx = new cl_CIELO_REQUISICAO_TRANSACAO_TX

	bandeira = Lcase(Trim("" & bandeira))
	
	trx.dadosPedidoMoeda = "986"
	trx.dadosPedidoIdioma = "PT"

	trx.dadosEcNumero = CIELO__NUMERO_CIELO
	trx.dadosEcChave = CIELO__CHAVE_CIELO

'	Define se a transa��o ser� automaticamente capturada caso seja autorizada
	trx.capturar = "false"
	
'	Indicador de autoriza��o autom�tica:
'	0 = n�o autorizar
'	1 = autorizar somente se autenticada
'	2 = autorizar autenticada e n�o-autenticada
'	3 = autorizar sem passar por autentica��o - v�lido somente para cr�dito)
'	Para Diners, Discover, Elo, Amex, Aura e JCB o valor ser� sempre �3�, pois estas bandeiras n�o possuem programa de autentica��o.
	if (bandeira = Lcase(CIELO_BANDEIRA__DINERS)) Or _
		(bandeira = Lcase(CIELO_BANDEIRA__DISCOVER)) Or _
		(bandeira = Lcase(CIELO_BANDEIRA__ELO)) Or _
		(bandeira = Lcase(CIELO_BANDEIRA__AMEX)) Or _
		(bandeira = Lcase(CIELO_BANDEIRA__AURA)) Or _
		(bandeira = Lcase(CIELO_BANDEIRA__JCB)) then
		trx.autorizar = "3"
	else
		trx.autorizar = "2"
		end if
	
	set cria_instancia_cl_CIELO_REQUISICAO_TRANSACAO_TX = trx
end function


' ------------------------------------------------------------------------
'   cria_instancia_cl_CIELO_REQUISICAO_CONSULTA_TX
function cria_instancia_cl_CIELO_REQUISICAO_CONSULTA_TX
dim trx
	set trx = new cl_CIELO_REQUISICAO_CONSULTA_TX

	trx.dadosEcNumero = CIELO__NUMERO_CIELO
	trx.dadosEcChave = CIELO__CHAVE_CIELO

	set cria_instancia_cl_CIELO_REQUISICAO_CONSULTA_TX = trx
end function


' ------------------------------------------------------------------------
'   CieloXmlHeader
private function CieloXmlHeader
dim xml
	xml = "<?xml version=""1.0"" encoding=""" & CIELO_XML_ENCODING & """ ?>"
	CieloXmlHeader = xml
end function


' ------------------------------------------------------------------------
'   CieloXmlRequisicaoTransacaoDadosEc
private function CieloXmlRequisicaoTransacaoDadosEc(ByRef trx)
dim xml
	xml = "<dados-ec>" & chr(13) & _
				"<numero>" & _
					trx.dadosEcNumero & _
				"</numero>" & chr(13) & _
				"<chave>" & _
					trx.dadosEcChave & _
				"</chave>" & chr(13) & _
			"</dados-ec>"
	
	CieloXmlRequisicaoTransacaoDadosEc = xml
end function


' ------------------------------------------------------------------------
'   CieloXmlRequisicaoTransacaoDadosPedido
private function CieloXmlRequisicaoTransacaoDadosPedido(ByRef trx)
dim xml
	xml = "<dados-pedido>" & chr(13) & _
				"<numero>" & _
					trx.dadosPedidoNumero & _
				"</numero>" & chr(13) & _
				"<valor>" & _
					trx.dadosPedidoValor & _
				"</valor>" & chr(13) & _
				"<moeda>" & _
					trx.dadosPedidoMoeda & _
				"</moeda>" & chr(13) & _
				"<data-hora>" & _
					trx.dadosPedidoData & _
				"</data-hora>" & chr(13)
	
	if Trim("" & trx.dadosPedidoDescricao) <> "" then
		xml = xml & _
				"<descricao>" & _
				"<![CDATA[" & _
					trx.dadosPedidoDescricao & _
				"]]>" & _
				"</descricao>" & chr(13)
		end if
	
	xml = xml & _
				"<idioma>" & _
					trx.dadosPedidoIdioma & _
				"</idioma>" & chr(13) & _
			"</dados-pedido>"
	
	CieloXmlRequisicaoTransacaoDadosPedido = xml
end function


' ------------------------------------------------------------------------
'   CieloXmlRequisicaoTransacaoFormaPagamento
private function CieloXmlRequisicaoTransacaoFormaPagamento(ByRef trx)
dim xml
	xml = "<forma-pagamento>" & chr(13) & _
				"<bandeira>" & _
					trx.formaPagamentoBandeira & _
				"</bandeira>" & chr(13) & _
				"<produto>" & _
					trx.formaPagamentoProduto & _
				"</produto>" & chr(13) & _
				"<parcelas>" & _
					trx.formaPagamentoParcelas & _
				"</parcelas>" & chr(13) & _
			"</forma-pagamento>"

	CieloXmlRequisicaoTransacaoFormaPagamento = xml
end function


' ------------------------------------------------------------------------
'   CieloXmlRequisicaoTransacaoUrlRetorno
private function CieloXmlRequisicaoTransacaoUrlRetorno(ByRef trx)
dim xml
	xml = "<url-retorno>" & trx.urlRetorno & "</url-retorno>"
	CieloXmlRequisicaoTransacaoUrlRetorno = xml
end function


' ------------------------------------------------------------------------
'   CieloXmlRequisicaoTransacaoAutorizar
private function CieloXmlRequisicaoTransacaoAutorizar(ByRef trx)
dim xml
	xml = "<autorizar>" & trx.autorizar & "</autorizar>"
	CieloXmlRequisicaoTransacaoAutorizar = xml
end function


' ------------------------------------------------------------------------
'   CieloXmlRequisicaoTransacaoCapturar
private function CieloXmlRequisicaoTransacaoCapturar(ByRef trx)
dim xml
	xml = "<capturar>" & trx.capturar & "</capturar>"
	CieloXmlRequisicaoTransacaoCapturar = xml
end function


' ------------------------------------------------------------------------
'   CieloXmlMontaRequisicaoTransacao
function CieloXmlMontaRequisicaoTransacao(ByRef trx, ByRef requisicao_transacao_id)
dim xml
	requisicao_transacao_id = Lcase(gera_uid)
	xml =	CieloXmlHeader & chr(13) & _
			"<requisicao-transacao id=""" & requisicao_transacao_id & """ versao=""" & CIELO_VERSAO_TRANSACAO & """>" & chr(13) & _
				CieloXmlRequisicaoTransacaoDadosEc(trx) & chr(13) & _
				CieloXmlRequisicaoTransacaoDadosPedido(trx) & chr(13) & _
				CieloXmlRequisicaoTransacaoFormaPagamento(trx) & chr(13) & _
				CieloXmlRequisicaoTransacaoUrlRetorno(trx) & chr(13) & _
				CieloXmlRequisicaoTransacaoAutorizar(trx) & chr(13) & _
				CieloXmlRequisicaoTransacaoCapturar(trx) & chr(13) & _
			"</requisicao-transacao>"
	
	CieloXmlMontaRequisicaoTransacao = xml
end function


' ------------------------------------------------------------------------
'   CieloXmlMontaRequisicaoConsulta
function CieloXmlMontaRequisicaoConsulta(ByRef trx, ByRef requisicao_consulta_id)
dim xml
	requisicao_consulta_id = Lcase(gera_uid)
	xml =	CieloXmlHeader & chr(13) & _
			"<requisicao-consulta id=""" & requisicao_consulta_id & """ versao=""" & CIELO_VERSAO_TRANSACAO & """>" & chr(13) & _
				"<tid>" & trx.tid & "</tid>" & chr(13) & _
				CieloXmlRequisicaoTransacaoDadosEc(trx) & chr(13) & _
			"</requisicao-consulta>"
	
	CieloXmlMontaRequisicaoConsulta = xml
end function


' ------------------------------------------------------------------------
'   CieloEnviaTransacao
'	Option: 2 = SXH_OPTION_IGNORE_SERVER_SSL_CERT_ERROR_FLAGS
'	The SXH_OPTION_IGNORE_SERVER_SSL_CERT_ERROR_FLAGS option is a DWORD mask of various flags that can be set to change this default behavior.
'	The default value is to ignore all problems. You must set this option before calling the send method. The flags are as follows:
'		SXH_SERVER_CERT_IGNORE_UNKNOWN_CA = 256
'		Unknown certificate authority
'		SXH_SERVER_CERT_IGNORE_WRONG_USAGE = 512
'		Malformed certificate such as a certificate with no subject name.
'		SXH_SERVER_CERT_IGNORE_CERT_CN_INVALID = 4096
'		Mismatch between the visited hostname and the certificate name being used on the server.
'		SXH_SERVER_CERT_IGNORE_CERT_DATE_INVALID = 8192
'		The date in the certificate is invalid or has expired.
'		SXH_SERVER_CERT_IGNORE_ALL_SERVER_ERRORS = 13056
'		All certificate errors.
'	To turn off a flag, you subtract it from the default value, which is the sum of all flags.
'	For example, to catch an invalid date in a certificate, you turn off the SXH_SERVER_CERT_IGNORE_CERT_DATE_INVALID flag as follows:
'	shx.setOption(2) = (shx.getOption(2) - SXH_SERVER_CERT_IGNORE_CERT_DATE_INVALID)
function CieloEnviaTransacao(Byval xml)
dim xmlhttp
	set xmlhttp = Server.CreateObject("MSXML2.ServerXMLHTTP")
	xmlhttp.open "POST", CIELO_WEB_SERVICE_ENDERECO, False
	xmlhttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
	xmlhttp.setOption 2, 13056
	xmlhttp.send xml
	CieloEnviaTransacao = xmlhttp.responseText
end function


' ------------------------------------------------------------------------
'   CieloDescricaoStatus
function CieloDescricaoStatus(Byval codigoStatus)
dim s_resp
	codigoStatus = Trim("" & codigoStatus)
		
	select case codigoStatus
		case CIELO_TRANSACAO_STATUS__CRIADA
			s_resp = "Criada"
		case CIELO_TRANSACAO_STATUS__EM_ANDAMENTO
			s_resp = "Em andamento"
		case CIELO_TRANSACAO_STATUS__AUTENTICADA
			s_resp = "Autenticada"
		case CIELO_TRANSACAO_STATUS__NAO_AUTENTICADA
			s_resp = "N�o autenticada"
		case CIELO_TRANSACAO_STATUS__AUTORIZADA
			s_resp = "Autorizada"
		case CIELO_TRANSACAO_STATUS__NAO_AUTORIZADA
			s_resp = "N�o autorizada"
		case CIELO_TRANSACAO_STATUS__CAPTURADA
			s_resp = "Capturada"
		case CIELO_TRANSACAO_STATUS__NAO_CAPTURADA
			s_resp = "N�o capturada"
		case CIELO_TRANSACAO_STATUS__CANCELADA
			s_resp = "Cancelada"
		case CIELO_TRANSACAO_STATUS__EM_AUTENTICACAO
			s_resp = "Em autentica��o"
		case ""
			s_resp = ""
		case else
			s_resp = "C�DIGO DESCONHECIDO: " & codigoStatus
	end select
	
	CieloDescricaoStatus = s_resp
end function


' ------------------------------------------------------------------------
'   CieloDescricaoOperacao
'   Retorna a descri��o para o c�digo de opera��o.
function CieloDescricaoOperacao(byval codigo_operacao)
dim s_resp

	select case codigo_operacao
		case OP_CIELO_OPERACAO__PAGAMENTO
			s_resp = "Pagamento"
		case OP_CIELO_OPERACAO__CANCELAMENTO
			s_resp = "Cancelamento"
		case else
			s_resp = ""
	end select
	
	CieloDescricaoOperacao = s_resp
end function


' ------------------------------------------------------------------------
'   CieloDescricaoParcelamento
'   Retorna a descri��o para a forma de pagamento selecionada.
function CieloDescricaoParcelamento(byval cod_produto, byval qtde_parcelas, byval valor_total)
dim s_resp
dim vl_parcela
dim vl_total

	cod_produto = Trim("" & cod_produto)
	vl_total = converte_numero(valor_total)
	if qtde_parcelas <> 0 then vl_parcela = vl_total / qtde_parcelas

	select case cod_produto
	'	CR�DITO � VISTA
		case "1"
			s_resp = SIMBOLO_MONETARIO & " " & formata_moeda(valor_total) & " � Vista (no Cr�dito)"
	'	PARCELADO LOJA
		case "2"
			s_resp = formata_inteiro(qtde_parcelas) & " x de " & SIMBOLO_MONETARIO & " " & formata_moeda(vl_parcela) & " iguais"
	'	PARCELADO ADMINISTRADORA
		case "3"
			s_resp = formata_inteiro(qtde_parcelas) & " x de " & SIMBOLO_MONETARIO & " " & formata_moeda(vl_parcela) & " mais juros"
	'	D�BITO
		case "A"
			s_resp = SIMBOLO_MONETARIO & " " & formata_moeda(valor_total) & " � Vista (no D�bito)"
		case else
			s_resp = ""
	end select

	CieloDescricaoParcelamento = s_resp
end function


' ------------------------------------------------------------------------
'   CieloDescricaoBandeira
function CieloDescricaoBandeira(Byval bandeira)
dim s_resp
	bandeira = Lcase(Trim("" & bandeira))
	if bandeira = "visa" then
		s_resp = "Visa"
	elseif bandeira = "mastercard" then
		s_resp = "Mastercard"
	elseif bandeira = "amex" then
		s_resp = "Amex"
	elseif bandeira = "elo" then
		s_resp = "Elo"
	elseif bandeira = "hipercard" then
		s_resp = "Hipercard"
	elseif bandeira = "diners" then
		s_resp = "Diners"
	elseif bandeira = "discover" then
		s_resp = "Discover"
	elseif bandeira = "aura" then
		s_resp = "Aura"
	elseif bandeira = "jcb" then
		s_resp = "JCB"
	elseif bandeira <> "" then
		s_resp = "Bandeira desconhecida (" & bandeira & ")"
	else
		s_resp = ""
		end if
		
	CieloDescricaoBandeira = s_resp
end function


' ------------------------------------------------------------------------
'	CieloObtemIdRegistroBdPrazoPagtoLoja
'	Dada a bandeira do cart�o, retorna o ID do registro da tabela
'	t_PRAZO_PAGTO_VISANET que cont�m os dados do parcelamento pela loja.
function CieloObtemIdRegistroBdPrazoPagtoLoja(Byval bandeira)
dim s_resp
	s_resp = ""
	bandeira = Ucase(Trim("" & bandeira))
	if bandeira = Ucase(CIELO_BANDEIRA__VISA) then
		s_resp = COD_VISANET_PRAZO_PAGTO_LOJA
	elseif bandeira = Ucase(CIELO_BANDEIRA__MASTERCARD) then
		s_resp = COD_MASTERCARD_PRAZO_PAGTO_LOJA
	elseif bandeira = Ucase(CIELO_BANDEIRA__AMEX) then
		s_resp = COD_AMEX_PRAZO_PAGTO_LOJA
	elseif bandeira = Ucase(CIELO_BANDEIRA__ELO) then
		s_resp = COD_ELO_PRAZO_PAGTO_LOJA
	elseif bandeira = Ucase(CIELO_BANDEIRA__HIPERCARD) then
		s_resp = COD_HIPERCARD_PRAZO_PAGTO_LOJA
	elseif bandeira = Ucase(CIELO_BANDEIRA__DINERS) then
		s_resp = COD_DINERS_PRAZO_PAGTO_LOJA
	elseif bandeira = Ucase(CIELO_BANDEIRA__DISCOVER) then
		s_resp = COD_DISCOVER_PRAZO_PAGTO_LOJA
	elseif bandeira = Ucase(CIELO_BANDEIRA__AURA) then
		s_resp = COD_AURA_PRAZO_PAGTO_LOJA
	elseif bandeira = Ucase(CIELO_BANDEIRA__JCB) then
		s_resp = COD_JCB_PRAZO_PAGTO_LOJA
	elseif bandeira = Ucase(CIELO_BANDEIRA__CELULAR) then
		s_resp = COD_CELULAR_PRAZO_PAGTO_LOJA
		end if
	CieloObtemIdRegistroBdPrazoPagtoLoja = s_resp
end function


' ------------------------------------------------------------------------
'	CieloObtemIdRegistroBdPrazoPagtoEmissor
'	Dada a bandeira do cart�o, retorna o ID do registro da tabela
'	t_PRAZO_PAGTO_VISANET que cont�m os dados do parcelamento pelo
'	emissor do cart�o.
function CieloObtemIdRegistroBdPrazoPagtoEmissor(Byval bandeira)
dim s_resp
	s_resp = ""
	bandeira = Ucase(Trim("" & bandeira))
	if bandeira = Ucase(CIELO_BANDEIRA__VISA) then
		s_resp = COD_VISANET_PRAZO_PAGTO_EMISSOR
	elseif bandeira = Ucase(CIELO_BANDEIRA__MASTERCARD) then
		s_resp = COD_MASTERCARD_PRAZO_PAGTO_EMISSOR
	elseif bandeira = Ucase(CIELO_BANDEIRA__AMEX) then
		s_resp = COD_AMEX_PRAZO_PAGTO_EMISSOR
	elseif bandeira = Ucase(CIELO_BANDEIRA__ELO) then
		s_resp = COD_ELO_PRAZO_PAGTO_EMISSOR
	elseif bandeira = Ucase(CIELO_BANDEIRA__HIPERCARD) then
		s_resp = COD_HIPERCARD_PRAZO_PAGTO_EMISSOR
	elseif bandeira = Ucase(CIELO_BANDEIRA__DINERS) then
		s_resp = COD_DINERS_PRAZO_PAGTO_EMISSOR
	elseif bandeira = Ucase(CIELO_BANDEIRA__DISCOVER) then
		s_resp = COD_DISCOVER_PRAZO_PAGTO_EMISSOR
	elseif bandeira = Ucase(CIELO_BANDEIRA__AURA) then
		s_resp = COD_AURA_PRAZO_PAGTO_EMISSOR
	elseif bandeira = Ucase(CIELO_BANDEIRA__JCB) then
		s_resp = COD_JCB_PRAZO_PAGTO_EMISSOR
	elseif bandeira = Ucase(CIELO_BANDEIRA__CELULAR) then
		s_resp = COD_CELULAR_PRAZO_PAGTO_EMISSOR
		end if
	CieloObtemIdRegistroBdPrazoPagtoEmissor = s_resp
end function


' ------------------------------------------------------------------------
'	CieloObtemNomeArquivoLogo
'	Dada a bandeira do cart�o, retorna o nome do arquivo que cont�m o
'	logotipo.
function CieloObtemNomeArquivoLogo(Byval bandeira)
dim s_resp
	s_resp = ""
	bandeira = Ucase(Trim("" & bandeira))
	if bandeira = Ucase(CIELO_BANDEIRA__VISA) then
		s_resp = "LogoVisa.gif"
	elseif bandeira = Ucase(CIELO_BANDEIRA__MASTERCARD) then
		s_resp = "mastercard.gif"
	elseif bandeira = Ucase(CIELO_BANDEIRA__AMEX) then
		s_resp = "Amex.gif"
	elseif bandeira = Ucase(CIELO_BANDEIRA__ELO) then
		s_resp = "Elo.gif"
	elseif bandeira = Ucase(CIELO_BANDEIRA__HIPERCARD) then
		s_resp = "Hipercard.gif"
	elseif bandeira = Ucase(CIELO_BANDEIRA__DINERS) then
		s_resp = "Diners.gif"
	elseif bandeira = Ucase(CIELO_BANDEIRA__DISCOVER) then
		s_resp = "Discover.gif"
	elseif bandeira = Ucase(CIELO_BANDEIRA__AURA) then
		s_resp = "Aura.gif"
	elseif bandeira = Ucase(CIELO_BANDEIRA__JCB) then
		s_resp = "JCB.gif"
	elseif bandeira = Ucase(CIELO_BANDEIRA__CELULAR) then
		s_resp = "Celular.gif"
	else
		s_resp = "Unknown.gif"
		end if
	CieloObtemNomeArquivoLogo = s_resp
end function


' ------------------------------------------------------------------------
'	CieloQtdeBandeirasHabilitadas
'	Calcula a quantidade de bandeiras ativas que est�o dispon�veis para
'	serem usadas nas transa��es.
function CieloQtdeBandeirasHabilitadas
dim qtdeBandeiras
	qtdeBandeiras = 0
	if CIELO_BANDEIRA_HABILITADA__VISA then qtdeBandeiras = qtdeBandeiras + 1
	if CIELO_BANDEIRA_HABILITADA__MASTERCARD then qtdeBandeiras = qtdeBandeiras + 1
	if CIELO_BANDEIRA_HABILITADA__AMEX then qtdeBandeiras = qtdeBandeiras + 1
	if CIELO_BANDEIRA_HABILITADA__ELO then qtdeBandeiras = qtdeBandeiras + 1
	if CIELO_BANDEIRA_HABILITADA__DINERS then qtdeBandeiras = qtdeBandeiras + 1
	if CIELO_BANDEIRA_HABILITADA__DISCOVER then qtdeBandeiras = qtdeBandeiras + 1
	if CIELO_BANDEIRA_HABILITADA__AURA then qtdeBandeiras = qtdeBandeiras + 1
	if CIELO_BANDEIRA_HABILITADA__JCB then qtdeBandeiras = qtdeBandeiras + 1
	CieloQtdeBandeirasHabilitadas = qtdeBandeiras
end function


' ------------------------------------------------------------------------
'	CieloArrayBandeiras
'	Cria e retorna um array contendo as bandeiras existentes, ou seja,
'	independentemente da bandeira estar habilitada ou n�o.
function CieloArrayBandeiras
	CieloArrayBandeiras = Array(CIELO_BANDEIRA__VISA, _
							CIELO_BANDEIRA__MASTERCARD, _
							CIELO_BANDEIRA__AMEX, _
							CIELO_BANDEIRA__ELO, _
							CIELO_BANDEIRA__HIPERCARD, _
							CIELO_BANDEIRA__DINERS, _
							CIELO_BANDEIRA__DISCOVER, _
							CIELO_BANDEIRA__AURA, _
							CIELO_BANDEIRA__JCB, _
							CIELO_BANDEIRA__CELULAR)
end function


' ------------------------------------------------------------------------
'	CieloSelecaoBandeiraQtdePorLinha
'	Calcula quantas bandeiras devem ser exibidas por linha na tela de
'	escolha da bandeira a ser usada no pagamento.
function CieloSelecaoBandeiraQtdePorLinha
dim qtdeBandeiras
dim qtdePorLinha
	qtdeBandeiras=CieloQtdeBandeirasHabilitadas
	select case qtdeBandeiras
		case 1, 2, 3, 4
			qtdePorLinha = qtdeBandeiras
		case 5
			qtdePorLinha = 3	' L1 = 3, L2 = 2
		case 6
			qtdePorLinha = 3	' L1 = 3, L2 = 3
		case 7
			qtdePorLinha = 4	' L1 = 4, L2 = 3
		case 8
			qtdePorLinha = 4	' L1 = 4, L2 = 4
		case 9
			qtdePorLinha = 3	' L1 = 3, L2 = 3, L3 = 3
		case 10
			qtdePorLinha = 4	' L1 = 4, L2 = 4, L3 = 2
		case 11
			qtdePorLinha = 4	' L1 = 4, L2 = 4, L3 = 3
		case 12
			qtdePorLinha = 4	' L1 = 4, L2 = 4, L3 = 4
		case else
			qtdePorLinha = 4
	end select
	
	CieloSelecaoBandeiraQtdePorLinha = qtdePorLinha
end function


' ------------------------------------------------------------------------
'	IIF
'	SIMULA A FUN�O IIF DO VISUAL BASIC
'		blnCondicao: par�metro que deve conter True ou False
'		retornoTrue: par�metro cujo conte�do ser� retornado se a condi��o for True
'		retornoFalse: par�metro cujo conte�do ser� retornado se a condi��o for False
'	Ex: s = iif( (Trim(strParametro)=""), "&nbsp;", strParametro)
function iif(byval blnCondicao, byval retornoTrue, byval retornoFalse)
	if vartype(blnCondicao) <> vbBoolean then
		iif = Null
		exit function
		end if
	
	if blnCondicao then
		iif = retornoTrue
	else
		iif = retornoFalse
		end if
end function


' ------------------------------------------------------------------------
'   PRIMEIRO NAO VAZIO
function primeiroNaoVazio(lista())
dim i
	primeiroNaoVazio = ""
	for i = Lbound(lista) to Ubound(lista)
		if Trim("" & lista(i)) <> "" then
			primeiroNaoVazio = lista(i)
			exit function
			end if
		next
end function


' ------------------------------------------------------------------------
'   TrimRightCrLf
function TrimRightCrLf(byval texto)
dim c, strResp
	strResp = "" & texto
	do while len(strResp) > 0
		c = Right(strResp, 1)
		if (c = vbCr) Or (c = vbLf) Or (c = " ") then
			if Len(strResp) = 1 then
				strResp = ""
			else
				strResp = Left(strResp, Len(strResp)-1)
				end if
		else
			exit do
			end if
		loop
	TrimRightCrLf = strResp
end function

'-------------------------------------------------------------------------
'   Codigo Produto Complemento

function codigoProdutoComplemento(produto)
dim i
dim v_codigo()

   redim  preserve v_codigo(Len(produto))
  for i = 1  to Ubound(v_codigo)
      v_codigo(i) = Mid(produto,i,1)
      if IsDigit(v_codigo(i)) = true then
   v_codigo(i)= chr(v_codigo(i))
   v_codigo(i)= asc(v_codigo(i))
     codigoProdutoComplemento = codigoProdutoComplemento + cstr(57 - v_codigo(i))
    
    elseif isLetra(v_codigo(i)) = true then
       v_codigo(i)= chr(UCase(v_codigo(i)))
       v_codigo(i)= asc(v_codigo(i))
       codigoProdutoComplemento = codigoProdutoComplemento + cstr(90 - v_codigo(i))
       
    else
        codigoProdutoComplemento = codigoProdutoComplemento + v_codigo(i)  
   
    end if
    next
end function


'-------------------------------------------------------------------------
function floor(x)
        dim temp
 
        temp = Round(x)
 
        if temp > x then
            temp = temp - 1
        end if
 
        floor = temp
    end function


' ------------------------------------------------------------------------
'   limpa_cl_t_PARAMETRO
sub limpa_cl_t_PARAMETRO(ByRef rx)
	if rx is nothing then exit sub
	rx.id = ""
	rx.campo_inteiro = 0
	rx.campo_monetario = 0
	rx.campo_real = 0
	rx.campo_data = Null
	rx.campo_texto = ""
	rx.dt_hr_ult_atualizacao = Null
	rx.usuario_ult_atualizacao = ""
end sub


' ------------------------------------------------------------------------
'   limpa_cl_NFE_EMITENTE
sub limpa_cl_NFE_EMITENTE(ByRef rx)
	if rx is nothing then exit sub
	rx.id = 0
	rx.id_boleto_cedente = 0
	rx.braspag_id_boleto_cedente = 0
	rx.st_ativo = 0
	rx.apelido = ""
	rx.cnpj = ""
	rx.razao_social = ""
	rx.endereco = ""
	rx.endereco_numero = ""
	rx.endereco_complemento = ""
	rx.bairro = ""
	rx.cidade = ""
	rx.uf = ""
	rx.cep = ""
	rx.NFe_st_emitente_padrao = 0
	rx.NFe_serie_NF = 0
	rx.NFe_numero_NF = 0
	rx.NFe_T1_servidor_BD = ""
	rx.NFe_T1_nome_BD = ""
	rx.NFe_T1_usuario_BD = ""
	rx.NFe_T1_senha_BD = ""
	rx.dt_cadastro = Null
	rx.dt_hr_cadastro = Null
	rx.usuario_cadastro = ""
	rx.dt_ult_atualizacao = Null
	rx.dt_hr_ult_atualizacao = Null
	rx.usuario_ult_atualizacao = ""
	rx.st_habilitado_ctrl_estoque = 0
	rx.ordem = 0
	rx.texto_fixo_especifico = ""
end sub

' ------------------------------------------------------------------------
'   isEnderecoIgual
function isEnderecoIgual(ByVal end_logradouro_1, ByVal end_numero_1, ByVal end_cep_1, ByVal end_logradouro_2, ByVal end_numero_2, ByVal end_cep_2)
const PREFIXOS = "|R|RUA|AV|AVEN|AVENIDA|TV|TRAV|TRAVESSA|AL|ALAM|ALAMEDA|PC|PRACA|PQ|PARQUE|EST|ESTR|ESTRADA|CJ|CONJ|CONJUNTO|"
dim v1, v2
dim s, s1, s2
dim i, j
dim blnFlag, blnNumeroIgual
dim v_end_numero_1, v_end_numero_2
dim n_end_numero_1, n_end_numero_2

	isEnderecoIgual = False
	
'	Normaliza
	end_logradouro_1 = retira_acentuacao(Ucase(Trim("" & end_logradouro_1)))
	end_numero_1 = Ucase(Trim("" & end_numero_1))
	end_cep_1 = retorna_so_digitos(Trim("" & end_cep_1))
	end_logradouro_2 = retira_acentuacao(Ucase(Trim("" & end_logradouro_2)))
	end_numero_2 = Ucase(Trim("" & end_numero_2))
	end_cep_2 = retorna_so_digitos(Trim("" & end_cep_2))
	
	if end_cep_1 <> end_cep_2 then exit function
	
'	COMPARA OS N�MEROS DO ENDERE�O, LEVANDO EM CONSIDERA��O CASOS COMO: 476/478
	blnNumeroIgual = False
	
	if end_numero_1 = end_numero_2 then blnNumeroIgual = True
	
	if Not blnNumeroIgual then
		v_end_numero_1 = Split(end_numero_1, "/")
		n_end_numero_1 = 0
		for i=LBound(v_end_numero_1) to UBound(v_end_numero_1)
			if retorna_so_digitos(v_end_numero_1(i)) <> "" then n_end_numero_1 = n_end_numero_1 + 1
			next
		
		v_end_numero_2 = Split(end_numero_2, "/")
		n_end_numero_2 = 0
		for i=LBound(v_end_numero_2) to UBound(v_end_numero_2)
			if retorna_so_digitos(v_end_numero_2(i)) <> "" then n_end_numero_2 = n_end_numero_2 + 1
			next
		
		if (n_end_numero_1 = 1) And (n_end_numero_2 = 1) then
			if end_numero_1 <> end_numero_2 then exit function
		else
			for i=LBound(v_end_numero_1) to UBound(v_end_numero_1)
				if retorna_so_digitos(v_end_numero_1(i)) <> "" then
					for j=LBound(v_end_numero_2) to UBound(v_end_numero_2)
						if retorna_so_digitos(v_end_numero_2(j)) <> "" then
							if Trim(v_end_numero_1(i)) = Trim(v_end_numero_2(j)) then
								blnNumeroIgual = True
								exit for
								end if
							end if
						next
					if blnNumeroIgual then exit for
					end if
				next
			end if
		end if
	
	if Not blnNumeroIgual then exit function

	end_logradouro_1 = Replace(end_logradouro_1, ",", " ")
	end_logradouro_1 = Replace(end_logradouro_1, ".", " ")
	end_logradouro_1 = Replace(end_logradouro_1, "-", " ")
	end_logradouro_1 = Replace(end_logradouro_1, ";", " ")
	end_logradouro_1 = Replace(end_logradouro_1, ":", " ")
	end_logradouro_1 = Replace(end_logradouro_1, "=", " ")
	
	end_logradouro_2 = Replace(end_logradouro_2, ",", " ")
	end_logradouro_2 = Replace(end_logradouro_2, ".", " ")
	end_logradouro_2 = Replace(end_logradouro_2, "-", " ")
	end_logradouro_2 = Replace(end_logradouro_2, ";", " ")
	end_logradouro_2 = Replace(end_logradouro_2, ":", " ")
	end_logradouro_2 = Replace(end_logradouro_2, "=", " ")
	
	v1 = Split(end_logradouro_1, " ")
	v2 = Split(end_logradouro_2, " ")
	
	s1 = ""
	for i=Lbound(v1) to Ubound(v1)
		blnFlag = False
		s = Trim("" & v1(i))
		if s <> "" then
			if s1 = "" then
				if Instr(PREFIXOS, "|" & s & "|") = 0 then blnFlag = True
			else
				blnFlag = True
				end if
			
			if blnFlag then
				if s1 <> "" then s1 = s1 & " "
				s1 = s1 & s
				end if
			end if
		next
	
	s2 = ""
	for i=Lbound(v2) to Ubound(v2)
		blnFlag = False
		s = Trim("" & v2(i))
		if s <> "" then
			if s2 = "" then
				if Instr(PREFIXOS, "|" & s & "|") = 0 then blnFlag = True
			else
				blnFlag = True
				end if
			
			if blnFlag then
				if s2 <> "" then s2 = s2 & " "
				s2 = s2 & s
				end if
			end if
		next
	
	if s1 <> s2 then exit function
	
	isEnderecoIgual = True
end Function


' ------------------------------------------------------------------------
'   xml_monta_campo
'
function xml_monta_campo(Byval conteudo, Byval tag_name, Byval qtde_tabs)
dim strResposta
dim opening_tag, closing_tag
	
	xml_monta_campo = ""
	if Trim("" & conteudo) = "" then exit function
	
	opening_tag = "<" & tag_name & ">"
	closing_tag = "</" & tag_name & ">"
	
	strResposta = ""
	if IsNumeric(qtde_tabs) then
		strResposta = String(CLng(qtde_tabs), vbTab)
		end if
	
	strResposta = strResposta & opening_tag & conteudo & closing_tag & chr(13)
	xml_monta_campo = strResposta
end function


' ------------------------------------------------------------------------
'   obtemDescricaoCtrlPagtoModulo
'
function obtemDescricaoCtrlPagtoModulo(byval codigo)
dim strCodigo, strResp
	strCodigo = Trim("" & codigo)
	if strCodigo = CTRL_PAGTO_MODULO__BOLETO then
		strResp = "Boleto"
	elseif strCodigo = CTRL_PAGTO_MODULO__CHEQUE then
		strResp = "Cheque"
	elseif strCodigo = CTRL_PAGTO_MODULO__VISA then
		strResp = "Cart�o"
	elseif strCodigo = CTRL_PAGTO_MODULO__BRASPAG_CARTAO then
		strResp = "Cart�o"
	elseif strCodigo = CTRL_PAGTO_MODULO__BRASPAG_CLEARSALE then
		strResp = "Cart�o"
	elseif strCodigo = CTRL_PAGTO_MODULO__BRASPAG_WEBHOOK then
		strResp = "Boleto (EC)"
	else
		if strCodigo <> "" then
			strResp = strCodigo & " - C�digo Desconhecido"
		else
			strResp = ""
			end if
		end if
	obtemDescricaoCtrlPagtoModulo = strResp
end function


' ------------------------------------------------------------------------
'   DecodeUTF8
'
Public Function DecodeUTF8(byval texto)
dim stmANSI
	Set stmANSI = Server.CreateObject("ADODB.Stream")
	texto = texto & ""
	
	On Error Resume Next

	With stmANSI
		.Open
		.Position = 0
		.CharSet = "Windows-1252"
		.WriteText texto
		.Position = 0
		.CharSet = "UTF-8"
	End With

	DecodeUTF8 = stmANSI.ReadText
	stmANSI.Close

	If Err.number <> 0 Then
		DecodeUTF8 = texto
	End If
	On error Goto 0
End Function


' ------------------------------------------------------------------------
'   inicializa_cl_CTRL_ESTOQUE_PEDIDO_ITEM_NOVO
'
sub inicializa_cl_CTRL_ESTOQUE_PEDIDO_ITEM_NOVO(byref o)
	o.fabricante = ""
	o.produto = ""
	o.descricao = ""
	o.descricao_html = ""
	o.qtde_solicitada = 0
	o.qtde_estoque = 0
	o.qtde_estoque_global = 0
end sub


' ------------------------------------------------------------------------
'   gera_letra_pedido_filhote
'   Gera a letra do sufixo para o pedido filhote:
'       indice_numeracao_filhote = 0 => "" (string vazia)
'       indice_numeracao_filhote = 1 => A
'       indice_numeracao_filhote = 2 => B
'       indice_numeracao_filhote = 3 => C
'       Etc
function gera_letra_pedido_filhote(byval indice_numeracao_filhote)
dim s_letra
	gera_letra_pedido_filhote = ""
	if converte_numero(indice_numeracao_filhote) <= 0 then exit function
	s_letra = Chr((Asc("A")-1) + indice_numeracao_filhote)
	gera_letra_pedido_filhote = s_letra
end function


' ------------------------------------------------------------------------
'   elimina_html_entities
'   Elimina todas as ocorr�ncias de HTML Entities dentro de uma string
'   usando Regular Expression
function elimina_html_entities(ByVal str)
dim regEx, matches, match

    set regEx = New RegExp

    with regEx
        .Pattern = "(&#(\d{1,4});|&(\w+);)"
        .Global = True
    end with

    set matches = regEx.Execute(str)

    For Each match in matches
        str = Replace(str, match.Value, converte_html_name_caractere(match.Value))
    Next

    elimina_html_entities = Trim(str)
    
end function


' ------------------------------------------------------------------------
'   converte_html_name_caractere
'   Converte o HTML Name de um caractere especial
function converte_html_name_caractere(byval codigo)
dim s_resp

	s_resp=""
	
	select case codigo
        case "&quot;"
            s_resp = chr(34)
        case "&amp;"
            s_resp = chr(38)
        case "&lt;"
            s_resp = chr(60)
        case "&gt;"
            s_resp = chr(62)
        case "&copy;"
            s_resp = chr(169)
        case "&ordf;"
            s_resp = chr(170)
        case "&reg;"
            s_resp = chr(174)
        case else
            s_resp = ""
    end select

	converte_html_name_caractere = s_resp
end function


' ------------------------------------------------------------------------
'   copia_cl_ITEM_ORCAMENTO_para_cl_ITEM_ORCAMENTO_NOVO
'
function copia_cl_ITEM_ORCAMENTO_para_cl_ITEM_ORCAMENTO_NOVO(byval v_origem, byref v_destino, byref msg_erro)
dim i
	copia_cl_ITEM_ORCAMENTO_para_cl_ITEM_ORCAMENTO_NOVO = False
	msg_erro = ""
	Err.Clear
	redim v_destino(UBound(v_origem))
	for i=LBound(v_destino) to UBound(v_destino)
		set v_destino(i) = New cl_ITEM_ORCAMENTO_NOVO
		v_destino(i).orcamento = v_origem(i).orcamento
		v_destino(i).fabricante = v_origem(i).fabricante
		v_destino(i).produto = v_origem(i).produto
		v_destino(i).qtde = v_origem(i).qtde
		v_destino(i).qtde_spe = v_origem(i).qtde_spe
		v_destino(i).desc_dado = v_origem(i).desc_dado
		v_destino(i).preco_venda = v_origem(i).preco_venda
		v_destino(i).preco_NF = v_origem(i).preco_NF
		v_destino(i).preco_fabricante = v_origem(i).preco_fabricante
		v_destino(i).preco_lista = v_origem(i).preco_lista
		v_destino(i).margem = v_origem(i).margem
		v_destino(i).desc_max = v_origem(i).desc_max
		v_destino(i).comissao = v_origem(i).comissao
		v_destino(i).descricao = v_origem(i).descricao
		v_destino(i).descricao_html = v_origem(i).descricao_html
		v_destino(i).obs = v_origem(i).obs
		v_destino(i).ean = v_origem(i).ean
		v_destino(i).grupo = v_origem(i).grupo
		v_destino(i).peso = v_origem(i).peso
		v_destino(i).qtde_volumes = v_origem(i).qtde_volumes
		v_destino(i).abaixo_min_status = v_origem(i).abaixo_min_status
		v_destino(i).abaixo_min_autorizacao = v_origem(i).abaixo_min_autorizacao
		v_destino(i).abaixo_min_autorizador = v_origem(i).abaixo_min_autorizador
		v_destino(i).markup_fabricante = v_origem(i).markup_fabricante
		v_destino(i).sequencia = v_origem(i).sequencia
		v_destino(i).abaixo_min_superv_autorizador = v_origem(i).abaixo_min_superv_autorizador
		v_destino(i).vl_custo2 = v_origem(i).vl_custo2
		v_destino(i).custoFinancFornecCoeficiente = v_origem(i).custoFinancFornecCoeficiente
		v_destino(i).custoFinancFornecPrecoListaBase = v_origem(i).custoFinancFornecPrecoListaBase
		v_destino(i).cubagem = v_origem(i).cubagem
		v_destino(i).ncm = v_origem(i).ncm
		v_destino(i).cst = v_origem(i).cst
		v_destino(i).descontinuado = v_origem(i).descontinuado
		v_destino(i).qtde_estoque_total_disponivel = 0
		v_destino(i).qtde_estoque_vendido = 0
		v_destino(i).qtde_estoque_sem_presenca = 0
		next

	if Err <> 0 then
		msg_erro = CStr(Err.number) & " - " & Err.Description
		exit function
		end if

	copia_cl_ITEM_ORCAMENTO_para_cl_ITEM_ORCAMENTO_NOVO = True
end function


' ------------------------------------------------------------------------
'   obtem_descricao_icms_contribuinte_x_produtor_rural
'
function obtem_descricao_icms_contribuinte_x_produtor_rural(byval tipo, byval st_pj, byval st_pf)
dim strResp

    tipo = Trim(tipo)
    st_pj = Trim(st_pj)
    st_pf = Trim(st_pf)

    if tipo = ID_PJ then
        select case st_pj
            case COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO
                strResp = "N�o Contribuinte"
            case COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM
                strResp = "Contribuinte"
            case COD_ST_CLIENTE_CONTRIBUINTE_ICMS_ISENTO
                strResp = "Isento"
            case else
                strResp = ""
        end select
    elseif tipo = ID_PF then
        select case st_pf
            case COD_ST_CLIENTE_PRODUTOR_RURAL_NAO
                strResp = "Pessoa F�sica"
            case COD_ST_CLIENTE_PRODUTOR_RURAL_SIM
                strResp = "Produtor Rural"
            case else
                strResp = ""
        end select
    end if

    obtem_descricao_icms_contribuinte_x_produtor_rural = strResp
end function


' ------------------------------------------------------------------------
'   inicializa_cl_CADASTRO_MULTI_CD_REGRA
'
sub inicializa_cl_CADASTRO_MULTI_CD_REGRA(byref oRegra)
	dim i, j, k
	set oRegra = new cl_CADASTRO_MULTI_CD_REGRA
	oRegra.id = 0
	oRegra.st_inativo = 0
	oRegra.apelido = ""
	oRegra.descricao = ""
	for i=LBound(oRegra.vUF) to UBound(oRegra.vUF)
		set oRegra.vUF(i) = new cl_CADASTRO_MULTI_CD_REGRA_UF
		oRegra.vUF(i).uf = ""
		oRegra.vUF(i).st_inativo = 0
		for j=LBound(oRegra.vUF(i).vPessoa) to UBound(oRegra.vUF(i).vPessoa)
			set oRegra.vUF(i).vPessoa(j) = new cl_CADASTRO_MULTI_CD_REGRA_PESSOA
			oRegra.vUF(i).vPessoa(j).tipo_pessoa = ""
			oRegra.vUF(i).vPessoa(j).st_inativo = 0
			oRegra.vUF(i).vPessoa(j).spe_id_nfe_emitente = 0
			for k=LBound(oRegra.vUF(i).vPessoa(j).vCD) to UBound(oRegra.vUF(i).vPessoa(j).vCD)
				set oRegra.vUF(i).vPessoa(j).vCD(k) = new cl_CADASTRO_MULTI_CD_REGRA_CD
				oRegra.vUF(i).vPessoa(j).vCD(k).id_nfe_emitente = 0
				oRegra.vUF(i).vPessoa(j).vCD(k).st_inativo = 0
				oRegra.vUF(i).vPessoa(j).vCD(k).ordem_prioridade = 0
				next
			next
		next
end sub


' ------------------------------------------------------------------------
'   inicializa_cl_SELECAO_MULTI_CD_REGRA
'
sub inicializa_cl_SELECAO_MULTI_CD_REGRA(byref oRegra)
dim i
	
	'� obrigat�rio que o objeto j� tenha sido instanciado, caso contr�rio,
	'mesmo que esta rotina instancie o objeto, no retorno da rotina a vari�vel
	'continuar� como Empty. Provavelmente isso acontece somente no caso em
	'que isso � feito tentando instanciar um objeto que � uma propriedade de
	'outro objeto.
	if IsEmpty(oRegra) then exit sub
	
	oRegra.id = 0
	oRegra.st_inativo = 0
	oRegra.apelido = ""
	oRegra.descricao = ""

	set oRegra.regraUF = new cl_SELECAO_MULTI_CD_REGRA_UF
	oRegra.regraUF.id = 0
	oRegra.regraUF.id_wms_regra_cd = 0
	oRegra.regraUF.uf = ""
	oRegra.regraUF.st_inativo = 0
	
	set oRegra.regraUF.regraPessoa = new cl_SELECAO_MULTI_CD_REGRA_PESSOA
	oRegra.regraUF.regraPessoa.id = 0
	oRegra.regraUF.regraPessoa.id_wms_regra_cd_x_uf = 0
	oRegra.regraUF.regraPessoa.tipo_pessoa = ""
	oRegra.regraUF.regraPessoa.st_inativo = 0
	oRegra.regraUF.regraPessoa.spe_id_nfe_emitente = 0
	for i=LBound(oRegra.regraUF.regraPessoa.vCD) to UBound(oRegra.regraUF.regraPessoa.vCD)
		set oRegra.regraUF.regraPessoa.vCD(i) = new cl_SELECAO_MULTI_CD_REGRA_CD
		oRegra.regraUF.regraPessoa.vCD(i).id = 0
		oRegra.regraUF.regraPessoa.vCD(i).id_wms_regra_cd_x_uf_x_pessoa = 0
		oRegra.regraUF.regraPessoa.vCD(i).id_nfe_emitente = 0
		oRegra.regraUF.regraPessoa.vCD(i).st_inativo = 0
		oRegra.regraUF.regraPessoa.vCD(i).ordem_prioridade = 0
		set oRegra.regraUF.regraPessoa.vCD(i).estoque = New cl_CTRL_ESTOQUE_PEDIDO_ITEM_NOVO
		inicializa_cl_CTRL_ESTOQUE_PEDIDO_ITEM_NOVO oRegra.regraUF.regraPessoa.vCD(i).estoque
		next
end sub


' ------------------------------------------------------------------------
'   inicializa_cl_PEDIDO_SELECAO_PRODUTO_REGRA
'
sub inicializa_cl_PEDIDO_SELECAO_PRODUTO_REGRA(byref o)
	set o = new cl_PEDIDO_SELECAO_PRODUTO_REGRA
	o.fabricante = ""
	o.produto = ""
	o.st_regra_ok = False
	o.msg_erro = ""
	
	'N�o � poss�vel refatorar a rotina criando uma fun��o espec�fica p/ inicializar a classe cl_SELECAO_MULTI_CD_REGRA
	'se a propriedade o.regra estiver com Empty. Apesar da rotina refatorada processar corretamente, no retorno da rotina
	'a propriedade o.regra continua com Empty.
	set o.regra = new cl_SELECAO_MULTI_CD_REGRA
	inicializa_cl_SELECAO_MULTI_CD_REGRA o.regra
end sub


' ------------------------------------------------------------------------
'   multi_cd_regra_determina_tipo_pessoa
'
function multi_cd_regra_determina_tipo_pessoa(byval tipo_cliente, byval contribuinte_icms_status, byval produtor_rural_status)
dim tipo_pessoa
	tipo_pessoa = ""
	
	if tipo_cliente = ID_PF then
		if converte_numero(produtor_rural_status) = converte_numero(COD_ST_CLIENTE_PRODUTOR_RURAL_SIM) then
			tipo_pessoa = COD_WMS_MULTI_CD_REGRA__TIPO_PESSOA__PRODUTOR_RURAL
		elseif converte_numero(produtor_rural_status) = converte_numero(COD_ST_CLIENTE_PRODUTOR_RURAL_NAO) then
			tipo_pessoa = COD_WMS_MULTI_CD_REGRA__TIPO_PESSOA__PESSOA_FISICA
			end if
	elseif tipo_cliente = ID_PJ then
		if converte_numero(contribuinte_icms_status) = converte_numero(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_SIM) then
			tipo_pessoa = COD_WMS_MULTI_CD_REGRA__TIPO_PESSOA__PESSOA_JURIDICA_CONTRIBUINTE
		elseif converte_numero(contribuinte_icms_status) = converte_numero(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_NAO) then
			tipo_pessoa = COD_WMS_MULTI_CD_REGRA__TIPO_PESSOA__PESSOA_JURIDICA_NAO_CONTRIBUINTE
		elseif converte_numero(contribuinte_icms_status) = converte_numero(COD_ST_CLIENTE_CONTRIBUINTE_ICMS_ISENTO) then
			tipo_pessoa = COD_WMS_MULTI_CD_REGRA__TIPO_PESSOA__PESSOA_JURIDICA_ISENTO
			end if
		end if

	multi_cd_regra_determina_tipo_pessoa = tipo_pessoa
end function



' ------------------------------------------------------------------------
'   ec_dados_formata_nome
'
function ec_dados_formata_nome(byval firstName, byval middleName, byval lastName, byval maxLength)
dim s_resp, s_aux, tamMax
	tamMax = converte_numero(maxLength)
	s_resp = Trim("" & firstName)
	s_aux = Trim("" & middleName)
	if (s_resp <> "") And (s_aux <> "") And (UCase(s_resp) <> UCase(s_aux)) then s_resp = s_resp & " " & s_aux
	s_aux = Trim("" & lastName)
	if (s_resp <> "") And (s_aux <> "") And (UCase(s_resp) <> UCase(s_aux)) then s_resp = s_resp & " " & s_aux

	if (tamMax > 0) then
		if Len(s_resp) > tamMax then
			'SE H� TAMANHO M�XIMO DEFINIDO E O NOME COMPLETO EXCEDE O LIMITE, UTILIZA APENAS O PRIMEIRO NOME E SOBRENOME
			s_resp = Trim("" & firstName)
			s_aux = Trim("" & lastName)
			if (s_resp <> "") And (s_aux <> "") And (UCase(s_resp) <> UCase(s_aux)) then s_resp = s_resp & " " & s_aux
			end if
		
		if Len(s_resp) > tamMax then
			'SE O PRIMEIRO NOME E SOBRENOME EXCEDEM O TAMANHO M�XIMO, TRUNCA
			s_resp = Left(s_resp, tamMax)
			end if
		end if
	
	ec_dados_formata_nome = s_resp
end function


' ------------------------------------------------------------------------
'   ec_dados_decodifica_telefone_formatado
'
function ec_dados_decodifica_telefone_formatado(byval telefone_formatado, byref ddd, byref telefone)
const BRANCO = " "
dim v, s_telefone_formatado, s_tel_aux
	ec_dados_decodifica_telefone_formatado = False
	ddd = ""
	telefone = ""

	s_telefone_formatado = Trim("" & telefone_formatado)

	if Len(s_telefone_formatado) < 8 then exit function

	s_telefone_formatado = Replace(s_telefone_formatado, "(", BRANCO)
	s_telefone_formatado = Replace(s_telefone_formatado, ")", BRANCO)
	s_telefone_formatado = Trim(s_telefone_formatado)
	do while Instr(s_telefone_formatado, BRANCO & BRANCO) <> 0
		s_telefone_formatado = Replace(s_telefone_formatado, BRANCO & BRANCO, BRANCO)
		loop

	if Instr(s_telefone_formatado, BRANCO) = 0 then
		'N�O ENCONTROU SEPARA��O ENTRE DDD E TELEFONE
		s_tel_aux = retorna_so_digitos(s_telefone_formatado)
		if Len(s_tel_aux) = 11 then
			'ASSUME QUE PROVAVELMENTE SE TRATA DE DDD + N� DE 9 D�GITOS
			ddd = Left(s_tel_aux, 2)
			telefone = Right(s_tel_aux, 9)
		elseif Len(s_tel_aux) = 10 then
			'ASSUME QUE PROVAVELMENTE SE TRATA DE DDD + N� DE 8 D�GITOS
			ddd = Left(s_tel_aux, 2)
			telefone = Right(s_tel_aux, 8)
		else
			'RETORNA O CONTE�DO RECEBIDO SEM FAZER A SEPARA��O ENTRE DDD E TELEFONE
			telefone = Trim("" & telefone_formatado)
			end if
	else
		v = Split(s_telefone_formatado, BRANCO)
		ddd = Trim("" & v(LBound(v)))
		telefone = Trim("" & v(LBound(v)+1))
		end if

	ec_dados_decodifica_telefone_formatado = True
end function


' ------------------------------------------------------------------------
'   ec_dados_filtra_email
'
function ec_dados_filtra_email(byval email)
dim s_email
	ec_dados_filtra_email = ""
	email = Trim("" & email)
	s_email = LCase(email)
	if Instr(s_email, "@skyhub.") <> 0 then exit function
	if Instr(s_email, "@mktp.") <> 0 then exit function
	if Instr(s_email, "@email.") <> 0 then exit function
	if Instr(s_email, "@teste.com") <> 0 then exit function
	if Instr(s_email, "@mail.mercadolivre.com") <> 0 then exit function
	if Instr(s_email, "n�o informado") <> 0 then exit function
	ec_dados_filtra_email = email
end function


' ------------------------------------------------------------------------
'   ec_dados_normaliza_telefones
'
function ec_dados_normaliza_telefones(byref telephone_ddd, byref telephone_num, byref celular_ddd, byref celular_num, byref fax_ddd, byref fax_num)
dim ddd_aux, num_aux
	ec_dados_normaliza_telefones = False

	telephone_ddd = retorna_so_digitos(Trim("" & telephone_ddd))
	telephone_num = retorna_so_digitos(Trim("" & telephone_num))
	celular_ddd = retorna_so_digitos(Trim("" & celular_ddd))
	celular_num = retorna_so_digitos(Trim("" & celular_num))
	fax_ddd = retorna_so_digitos(Trim("" & fax_ddd))
	fax_num = retorna_so_digitos(Trim("" & fax_num))

	'Verifica se o n� de celular est� no campo errado
	if (Len(celular_num) < 9) And ( (Len(telephone_num) >= 9) Or (Len(fax_num) >= 9) ) then
		if Len(telephone_num) >= 9 then
			ddd_aux = celular_ddd
			num_aux = celular_num
			celular_ddd = telephone_ddd
			celular_num = telephone_num
			telephone_ddd = ddd_aux
			telephone_num = num_aux
		elseif Len(fax_num) >= 9 then
			ddd_aux = celular_ddd
			num_aux = celular_num
			celular_ddd = fax_ddd
			celular_num = fax_num
			fax_ddd = ddd_aux
			fax_num = num_aux
			end if
		end if

	'Verifica se o n� do telefone est� no campo de fax
	if (Len(fax_num) > 0) And (Len(telephone_num) = 0) then
		telephone_ddd = fax_ddd
		telephone_num = fax_num
		fax_ddd = ""
		fax_num = ""
		end if

	'Verifica repeti��o de n�meros
	if (Len(telephone_num) > 0) And (telephone_ddd = celular_ddd) And (telephone_num = celular_num) then
		'Os n�meros s�o iguais, verifica se � n�mero de telefone fixo ou celular
		if Len(celular_num) >= 9 then
			'� um n�mero de celular, ent�o limpa o campo do telefone fixo
			telephone_ddd = ""
			telephone_num = ""
		else
			'� um n�mero de telefone fixo, ent�o limpa o campo de celular
			celular_ddd = ""
			celular_num = ""
			end if
		end if

	if (Len(fax_num) > 0) And (fax_ddd = celular_ddd) And (fax_num = celular_num) then
		'Os n�meros s�o iguais, verifica se � n�mero de telefone fixo ou celular
		if Len(celular_num) >= 9 then
			'� um n�mero de celular, ent�o limpa o campo do fax
			fax_ddd = ""
			fax_num = ""
		else
			'� um n�mero de telefone fixo, ent�o limpa o campo de celular
			celular_ddd = ""
			celular_num = ""
			end if
		end if

	if (Len(telephone_num) > 0) And (telephone_ddd = fax_ddd) And (telephone_num = fax_num) then
		'Os n�meros s�o iguais, mant�m somente o campo telephone
		fax_ddd = ""
		fax_num = ""
		end if

	ec_dados_normaliza_telefones = True
end function

' ___________________________________
' OBT�M DESCRI��O STATUS DEVOLU��O
'
function obtem_descricao_status_devolucao(byval st_codigo, byref st_devolucao_descricao, byref st_devolucao_cor)
dim s_descricao, s_cor
    st_codigo = Trim("" & st_codigo)
    if st_codigo = "" then exit function

    st_devolucao_descricao = ""
    st_devolucao_cor = ""

    s_descricao = ""
    s_cor = ""
    select case st_codigo
        case COD_ST_PEDIDO_DEVOLUCAO__CADASTRADA
                s_descricao = "Cadastrada"
                s_cor = "#0348E1"
            case COD_ST_PEDIDO_DEVOLUCAO__EM_ANDAMENTO
                s_descricao = "Em Andamento"
                s_cor = "#E26534"
            case COD_ST_PEDIDO_DEVOLUCAO__MERCADORIA_RECEBIDA
                s_descricao = "Mercadoria Recebida"
                s_cor = "#008080"
            case COD_ST_PEDIDO_DEVOLUCAO__FINALIZADA
                s_descricao = "Finalizada"
                s_cor = "#4FAB5B"
            case COD_ST_PEDIDO_DEVOLUCAO__REPROVADA
                s_descricao = "Reprovada"
                s_cor = "#B91832"
            case COD_ST_PEDIDO_DEVOLUCAO__CANCELADA
                s_descricao = "Cancelada"
                s_cor = "#C7465A"
            case else
                s_descricao = "Indefinido"
                s_cor = "#000000"
    end select
    st_devolucao_descricao = s_descricao
    st_devolucao_cor = s_cor
end function
%>