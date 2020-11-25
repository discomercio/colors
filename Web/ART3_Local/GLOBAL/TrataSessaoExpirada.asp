<%
' ======================================================================================
' Tratamento para Sess�o Expirada
' A rotina assume que:
'
'	1) Sempre haver� um par�metro chamado "SessionCtrl" que possui um
'	   conte�do criptografado.
'	   A exce��o � a p�gina de retorno da transa��o c/ a Cielo, pois n�o h� como
'	   incluir o par�metro "SessionCtrl" devido ao tamanho limitado da URL
'	   informada � Cielo. Lembrando que o acionamento dessa p�gina � feito pela
'	   Cielo. Portanto, neste caso, se a sess�o tiver expirado, a solu��o foi
'	   recriar a vari�vel Session("SessionCtrlInfo") a partir dos dados gravados
'	   no BD junto c/ o registro da transa��o.
'
'	2) O conte�do de "SessionCtrl" s�o os seguintes campos (respeitando
'	   a ordem) separados pelo caracter "|":
'			A) Usu�rio
'			B) M�dulo
'			C) Loja
'			D) Ticket
'			E) Data/hora do logon
'			F) Data/hora da �ltima atividade (no servidor)
'
'	3) Ao detectar que a sess�o expirou, a rotina tenta ler o par�metro
'	   "SessionCtrl". A partir desses dados, recria o objeto "Session"
'	   com o conte�do original.
'
'	4) O valor do campo "Ticket" � comparado com o valor gravado no banco
'	   de dados no �ltimo logon feito pelo usu�rio. Com isso � feita uma
'	   autentica��o para determinar se a sess�o continua v�lida, pois
'	   se o usu�rio tiver feito o logoff, o campo no banco de dados estar�
'	   vazio.
'
'	5) O campo "Data/hora da �ltima atividade" permite um controle manual
'	   p/ definir um tempo de timeout da sess�o. Por exemplo, podemos
'	   definir que ap�s 60min desde a �ltima atividade, a sess�o n�o ir�
'	   mais ser restabelecida, ou seja, estar� expirada de fato.
'
'	6) Os campos do tipo "data/hora" s�o informados como n�mero decimal. Ex:
'	   27/12/2007 15:00:04 = 39443,6250462963
'	   E s�o convertidos atrav�s da fun��es:
'	   strDataHora = Cstr(CDbl(Now))
'	   dteDataHora = CDate(strDataHora)
'
' Altera��o em 28/10/2020, para integrar o site em MVC da loja com o ASP.
'	Ao chamar o ASP, o site em MVC passa o par�metro OrigemSolicitacao=LojaMvc
'	Neste caso, o ASP deve usar a loja indicada na URL; se a loja da sess�o for
'	diferente da solicitada, deve recriar a sess�o a partir do SessionCtrlInfo. 
'	Se recriar a sess�o (porque a sess�o na existe ou porque � outra loja), 
'	n�o deve fazer o log de sess�o restaurada.
'
' ======================================================================================

dim str__SessionCtrlInfo
dim str__SessionCtrlParametroDecripto
dim arr__SessionCtrlParametro
dim str__SessionCtrlChaveCripto
dim str__SessionCtrlCampoAux
dim str__SessionCtrlCampoUsuario
dim str__SessionCtrlCampoModulo
dim str__SessionCtrlCampoLoja
dim str__SessionCtrlCampoTicket
dim str__SessionCtrlCampoDtHrLogon
dim str__SessionCtrlCampoDtHrUltAtividade
dim str__SessionCtrlSQL
dim str__SessionCtrlSenhaDecripto
dim str__SessionCtrlSenhaCripto
dim str__SessionCtrlChaveCriptoSenha
dim int__SessionCtrlIndice
dim bln__SessionCtrlRestaurarSessao
dim cn__SessionCtrl
dim rs__SessionCtrl, rs2__SessionCtrl

dim bln_OrigemSolicitacaoLojaMvc

bln_OrigemSolicitacaoLojaMvc = false
if Trim(Request("OrigemSolicitacao")) = "LojaMvc" then
	bln_OrigemSolicitacaoLojaMvc = true
	end if


if Trim(Session("usuario_atual")) = "" or bln_OrigemSolicitacaoLojaMvc then
	str__SessionCtrlInfo = Trim(Request("SessionCtrlInfo"))
	if str__SessionCtrlInfo = "" then str__SessionCtrlInfo = Trim(Session("SessionCtrlInfo"))
	
	if str__SessionCtrlInfo <> "" then
		'Decriptografa o par�metro
		str__SessionCtrlChaveCripto = gera_chave(FATOR_CRIPTO_SESSION_CTRL)
		str__SessionCtrlParametroDecripto = DecriptografaTexto(str__SessionCtrlInfo, str__SessionCtrlChaveCripto)
		'Separa os campos
		arr__SessionCtrlParametro = Split(str__SessionCtrlParametroDecripto, "|", -1)

		int__SessionCtrlIndice=(Lbound(arr__SessionCtrlParametro)-1)
		
		'Usu�rio
		int__SessionCtrlIndice = int__SessionCtrlIndice+1
		if int__SessionCtrlIndice <= Ubound(arr__SessionCtrlParametro) then
			str__SessionCtrlCampoAux = Trim(arr__SessionCtrlParametro(int__SessionCtrlIndice))
		else
			str__SessionCtrlCampoAux = ""
			end if
		str__SessionCtrlCampoUsuario = str__SessionCtrlCampoAux

		'M�dulo
		int__SessionCtrlIndice = int__SessionCtrlIndice+1
		if int__SessionCtrlIndice <= Ubound(arr__SessionCtrlParametro) then
			str__SessionCtrlCampoAux = Trim(arr__SessionCtrlParametro(int__SessionCtrlIndice))
		else
			str__SessionCtrlCampoAux = ""
			end if
		str__SessionCtrlCampoModulo = str__SessionCtrlCampoAux

		'Loja
		int__SessionCtrlIndice = int__SessionCtrlIndice+1
		if int__SessionCtrlIndice <= Ubound(arr__SessionCtrlParametro) then
			str__SessionCtrlCampoAux = Trim(arr__SessionCtrlParametro(int__SessionCtrlIndice))
		else
			str__SessionCtrlCampoAux = ""
			end if
		str__SessionCtrlCampoLoja = str__SessionCtrlCampoAux

		'Ticket
		int__SessionCtrlIndice = int__SessionCtrlIndice+1
		if int__SessionCtrlIndice <= Ubound(arr__SessionCtrlParametro) then
			str__SessionCtrlCampoAux = Trim(arr__SessionCtrlParametro(int__SessionCtrlIndice))
		else
			str__SessionCtrlCampoAux = ""
			end if
		str__SessionCtrlCampoTicket = str__SessionCtrlCampoAux
		
		'Data/hora logon
		int__SessionCtrlIndice = int__SessionCtrlIndice+1
		if int__SessionCtrlIndice <= Ubound(arr__SessionCtrlParametro) then
			str__SessionCtrlCampoAux = Trim(arr__SessionCtrlParametro(int__SessionCtrlIndice))
		else
			str__SessionCtrlCampoAux = ""
			end if
		str__SessionCtrlCampoDtHrLogon = str__SessionCtrlCampoAux
		
		'Data/hora da �ltima atividade (no servidor)
		int__SessionCtrlIndice = int__SessionCtrlIndice+1
		if int__SessionCtrlIndice <= Ubound(arr__SessionCtrlParametro) then
			str__SessionCtrlCampoAux = Trim(arr__SessionCtrlParametro(int__SessionCtrlIndice))
		else
			str__SessionCtrlCampoAux = ""
			end if
		str__SessionCtrlCampoDtHrUltAtividade = str__SessionCtrlCampoAux

		bln__SessionCtrlRestaurarSessao = True
		
		if (str__SessionCtrlCampoUsuario = "") Or _
		   (str__SessionCtrlCampoModulo = "") Or _
		   (str__SessionCtrlCampoTicket = "") Or _
		   (str__SessionCtrlCampoDtHrLogon = "") Or _
		   (str__SessionCtrlCampoDtHrUltAtividade = "") then 
			bln__SessionCtrlRestaurarSessao = False
			end if
		
		if str__SessionCtrlCampoModulo = SESSION_CTRL_MODULO_LOJA then
			if str__SessionCtrlCampoLoja = "" then bln__SessionCtrlRestaurarSessao = False

			'se a sess�o j� existe ent�o entramos aqui porque bln_OrigemSolicitacaoLojaMvc
			'quando ela j� existe, s� for�amos a recriar a sess�o se foi solicitada outra loja.
			'note que ele usa a loja passada na URL, e n�o o campo t_USUARIO.SessionCtrlLoja (mais abaixo)
			'O que � passado na URL pode estar diferente do banco, e queremos garantir que seja o que 
			'a LojaMvc esteja mostrando na tela.
			if Trim(Session("usuario_atual")) = str__SessionCtrlCampoUsuario and bln_OrigemSolicitacaoLojaMvc and Trim(Session("loja_atual")) = str__SessionCtrlCampoLoja then
				bln__SessionCtrlRestaurarSessao = False
				end if

			end if

		if bdd_conecta(cn__SessionCtrl) then 
			
			if bln__SessionCtrlRestaurarSessao then
				'Verifica se o tempo de sess�o inativa realmente j� foi excedido
				'Caso seja bln_OrigemSolicitacaoLojaMvc, nunca ir� ocorrer porque a LojaMvc sempre manda a data da �ltima atividade como agora
				if (CDbl(Now) - CDbl(str__SessionCtrlCampoDtHrUltAtividade)) > (SESSION_CTRL_TIMEOUT_SESSAO_MIN * (1/(24*60)))then
					bln__SessionCtrlRestaurarSessao = False
					'Limpa o campo ticket p/ assegurar que a sess�o est� expirada e tamb�m
					'para n�o exibir a mensagem no pr�ximo logon de que esta sess�o n�o foi 
					'encerrada corretamente.
					str__SessionCtrlSQL = "UPDATE t_USUARIO SET" & _
												" SessionCtrlTicket = NULL" & _
										  " WHERE" & _
												" usuario = '" & str__SessionCtrlCampoUsuario & "'"
					cn__SessionCtrl.Execute(str__SessionCtrlSQL)
					end if
				end if

			if bln__SessionCtrlRestaurarSessao then
				'Verifica se o ticket refere-se � sess�o atual
				str__SessionCtrlSQL = "SELECT " & _
					 					"*" & _
									  " FROM t_USUARIO" & _
									  " WHERE" & _
										" usuario = '" & str__SessionCtrlCampoUsuario & "'"
				set rs__SessionCtrl = cn__SessionCtrl.Execute(str__SessionCtrlSQL)
				if Not rs__SessionCtrl.Eof then
					if Trim("" & rs__SessionCtrl("SessionCtrlTicket")) <> str__SessionCtrlCampoTicket then bln__SessionCtrlRestaurarSessao = False
					if bln__SessionCtrlRestaurarSessao then
						'*********************
						' Recria a sess�o!!!
						'*********************
						Session("usuario_atual") = str__SessionCtrlCampoUsuario
						Session("SessionCtrlInfo") = str__SessionCtrlInfo
						Session("SessionCtrlTicket") = str__SessionCtrlCampoTicket
						'Senha
						str__SessionCtrlSenhaCripto = Trim("" & rs__SessionCtrl("datastamp"))
						str__SessionCtrlChaveCriptoSenha = gera_chave(FATOR_BD)
						decodifica_dado str__SessionCtrlSenhaCripto, str__SessionCtrlSenhaDecripto, str__SessionCtrlChaveCriptoSenha
						Session("senha_atual") = str__SessionCtrlSenhaDecripto
						'Permiss�es de acesso
						Session("lista_operacoes_permitidas") = obtem_operacoes_permitidas_usuario(cn__SessionCtrl, str__SessionCtrlCampoUsuario)
						Session("nivel_acesso_bloco_notas") = obtem_nivel_acesso_bloco_notas_pedido(cn__SessionCtrl, str__SessionCtrlCampoUsuario)
						Session("usuario_nome_atual") = Trim("" & rs__SessionCtrl("nome"))
						Session("DataHoraLogon") = rs__SessionCtrl("SessionCtrlDtHrLogon")
						Session("DataHoraUltRefreshSession") = Now
						Session("SessionCtrlRecuperadoAuto") = "S"
						if str__SessionCtrlCampoModulo = SESSION_CTRL_MODULO_LOJA then
							Session("loja_atual") = str__SessionCtrlCampoLoja
							Session("vendedor_loja") = (rs__SessionCtrl("vendedor_loja") <> 0)
							Session("vendedor_externo") = (rs__SessionCtrl("vendedor_externo") <> 0)
							str__SessionCtrlSQL = "SELECT * FROM t_LOJA WHERE CONVERT(smallint, loja) = " & str__SessionCtrlCampoLoja
							set rs2__SessionCtrl = cn__SessionCtrl.Execute(str__SessionCtrlSQL)
							if Not rs2__SessionCtrl.Eof then
								Session("loja_nome_atual") = Trim("" & rs2__SessionCtrl("nome"))
								end if
							rs2__SessionCtrl.Close
							set rs2__SessionCtrl = nothing
							end if
							
						'Log da sess�o restaurada
						if not bln_OrigemSolicitacaoLojaMvc then
							str__SessionCtrlSQL = "INSERT INTO t_SESSAO_RESTAURADA (" & _
														"Usuario, " & _
														"DataHora, " & _
														"Modulo, " & _
														"Loja, " & _
														"DtHrInicioSessao" & _
													") VALUES (" & _
														"'" & QuotedStr(str__SessionCtrlCampoUsuario) & "', " & _
														bd_formata_data_hora(Now) & ", " & _
														"'" & str__SessionCtrlCampoModulo & "', " & _
														"'" & str__SessionCtrlCampoLoja & "', " & _
														bd_formata_data_hora(rs__SessionCtrl("SessionCtrlDtHrLogon")) & _
													")"
							cn__SessionCtrl.Execute(str__SessionCtrlSQL)
							end if	'if not bln_OrigemSolicitacaoLojaMvc
						end if	'if bln__SessionCtrlRestaurarSessao then ' Recria a sess�o!!!
					end if  'if Not rs__SessionCtrl.Eof
				
				rs__SessionCtrl.Close
				set rs__SessionCtrl = nothing
				end if  'if (bln__SessionCtrlRestaurarSessao)

			cn__SessionCtrl.Close
			set cn__SessionCtrl = nothing
			end if  'if bdd_conecta(cn__SessionCtrl)

		end if  'if (str__SessionCtrlInfo <> "")

	end if  'if Trim(Session("usuario_atual")) = "" or bln_OrigemSolicitacaoLojaMvc
%>
