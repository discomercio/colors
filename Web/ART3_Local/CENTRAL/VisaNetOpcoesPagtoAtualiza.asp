<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp"        -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ============================================================
'	  V I S A N E T O P C O E S P A G T O A T U A L I Z A . A S P
'     ============================================================
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
'			I N I C I A L I Z A     P Á G I N A     A S P     N O     S E R V I D O R
' _____________________________________________________________________________________________

	On Error GoTo 0
	Err.Clear
	
	class cl_CIELO_CADASTRA_OPCOES_PAGTO
		dim bandeira
		dim strLojaQtdeParcelas
		dim intLojaQtdeParcelas
		dim strLojaParcelaMinima
		dim vlLojaParcelaMinima
		dim strEmissorQtdeParcelas
		dim intEmissorQtdeParcelas
		dim strEmissorParcelaMinima
		dim vlEmissorParcelaMinima
		end class
	
	dim s, s_aux, usuario, alerta
	
	usuario = trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, r
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

	if Not operacao_permitida(OP_CEN_OPCOES_PAGTO_VISANET, s_lista_operacoes_permitidas) then
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	Dim blnNovoRegistro
	Dim s_log
	Dim campos_a_omitir
	Dim vLog1()
	Dim vLog2()
	s_log = ""
	campos_a_omitir = ""
	
	dim vDados
	dim iQtdeBandeira, ic
	dim vBandeira
	iQtdeBandeira = 0
	vBandeira = CieloArrayBandeiras
	
	redim vDados(0)
	set vDados(Ubound(vDados)) = new cl_CIELO_CADASTRA_OPCOES_PAGTO
	vDados(UBound(vDados)).bandeira = ""
	
	for ic=Lbound(vBandeira) to Ubound(vBandeira)
		if Trim(vBandeira(ic)) <> "" then
			iQtdeBandeira = iQtdeBandeira + 1
			if Trim(vDados(UBound(vDados)).bandeira) <> "" then
				redim preserve vDados(Ubound(vDados)+1)
				set vDados(Ubound(vDados)) = new cl_CIELO_CADASTRA_OPCOES_PAGTO
				end if
		'	BANDEIRA
			vDados(UBound(vDados)).bandeira = vBandeira(ic)
		'	LOJA: QTDE DE PARCELAS
			s = Trim(Request.Form("C_QTDE_" & CieloObtemIdRegistroBdPrazoPagtoLoja(vBandeira(ic))))
			vDados(Ubound(vDados)).strLojaQtdeParcelas = s
			vDados(Ubound(vDados)).intLojaQtdeParcelas = converte_numero(s)
		'	LOJA: PARCELA MÍNIMA
			s = Trim(Request.Form("C_VL_" & CieloObtemIdRegistroBdPrazoPagtoLoja(vBandeira(ic))))
			vDados(Ubound(vDados)).strLojaParcelaMinima = s
			vDados(Ubound(vDados)).vlLojaParcelaMinima = converte_numero(s)
		'	EMISSOR: QTDE DE PARCELAS
			s = Trim(Request.Form("C_QTDE_" & CieloObtemIdRegistroBdPrazoPagtoEmissor(vBandeira(ic))))
			vDados(Ubound(vDados)).strEmissorQtdeParcelas = s
			vDados(Ubound(vDados)).intEmissorQtdeParcelas = converte_numero(s)
		'	EMISSOR: PARCELA MÍNIMA
			s = Trim(Request.Form("C_VL_" & CieloObtemIdRegistroBdPrazoPagtoEmissor(vBandeira(ic))))
			vDados(Ubound(vDados)).strEmissorParcelaMinima = s
			vDados(Ubound(vDados)).vlEmissorParcelaMinima = converte_numero(s)
			end if
		next
	
	dim erro_consistencia, erro_fatal
	
	erro_consistencia=false
	erro_fatal=false
	
	alerta = ""
	
	for ic=Lbound(vDados) to Ubound(vDados)
		if Trim(vDados(ic).bandeira) <> "" then
			if vDados(ic).intLojaQtdeParcelas < 0 then
				alerta=texto_add_br(alerta)
				alerta=alerta & "Bandeira " & CieloDescricaoBandeira(vDados(ic).bandeira) & ": parcelamento pela loja informa quantidade inválida (" & Cstr(vDados(ic).intLojaQtdeParcelas) & ")."
				end if
			if vDados(ic).vlLojaParcelaMinima < 0 then
				alerta=texto_add_br(alerta)
				alerta=alerta & "Bandeira " & CieloDescricaoBandeira(vDados(ic).bandeira) & ": parcelamento pela loja informa valor de parcela mínima inválido (" & formata_moeda(vDados(ic).vlLojaParcelaMinima) & ")."
				end if
			if vDados(ic).intEmissorQtdeParcelas < 0 then
				alerta=texto_add_br(alerta)
				alerta=alerta & "Bandeira " & CieloDescricaoBandeira(vDados(ic).bandeira) & ": parcelamento pelo emissor do cartão informa quantidade inválida (" & Cstr(vDados(ic).intEmissorQtdeParcelas) & ")."
				end if
			if vDados(ic).vlEmissorParcelaMinima < 0 then
				alerta=texto_add_br(alerta)
				alerta=alerta & "Bandeira " & CieloDescricaoBandeira(vDados(ic).bandeira) & ": parcelamento pelo emissor do cartão informa valor de parcela mínima inválido (" & formata_moeda(vDados(ic).vlEmissorParcelaMinima) & ")."
				end if
			end if
		next
	
	if alerta <> "" then erro_consistencia=True
	
	Err.Clear
	
	dim msg_erro
	
	if alerta = "" then
	'	~~~~~~~~~~~~~
		cn.BeginTrans
	'	~~~~~~~~~~~~~
		if Not cria_recordset_otimista(r, msg_erro) then
		'	~~~~~~~~~~~~~~~~
			cn.RollbackTrans
		'	~~~~~~~~~~~~~~~~
			Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
			end if
		
		for ic=Lbound(vDados) to Ubound(vDados)
			if alerta <> "" then exit for
			
			if Trim(vDados(ic).bandeira) <> "" then
			'	PARCELAMENTO PELA LOJA
				if alerta = "" then
					blnNovoRegistro = False
					s = "SELECT * FROM t_PRAZO_PAGTO_VISANET WHERE tipo = '" & CieloObtemIdRegistroBdPrazoPagtoLoja(vDados(ic).bandeira) & "'"
					if r.State <> 0 then r.Close
					r.Open s, cn
					if r.EOF then
						blnNovoRegistro = True
						r.Addnew
						r("tipo") = CieloObtemIdRegistroBdPrazoPagtoLoja(vDados(ic).bandeira)
						r("descricao") = "Parcelamento pela loja (sem juros)"
					else
						log_via_vetor_carrega_do_recordset r, vLog1, campos_a_omitir
						end if
					
					r("qtde_parcelas")=vDados(ic).intLojaQtdeParcelas
					r("vl_min_parcela")=vDados(ic).vlLojaParcelaMinima
					r("atualizacao_data")=Date
					r("atualizacao_hora")=retorna_so_digitos(formata_hora(Now))
					r("atualizacao_usuario")=usuario
					r.Update
					
					If Err = 0 then
						if blnNovoRegistro then
							log_via_vetor_carrega_do_recordset r, vLog1, campos_a_omitir
							s_log = log_via_vetor_monta_inclusao(vLog1)
							if s_log <> "" then
								grava_log usuario, "", "", "", OP_LOG_VISANET_OPCOES_PAGTO_INCLUSAO, s_log
								end if
						else
							log_via_vetor_carrega_do_recordset r, vLog2, campos_a_omitir
							s_log = log_via_vetor_monta_alteracao(vLog1, vLog2)
							if s_log <> "" then
								s_log="tipo=" & Trim("" & r("tipo")) & "; " & s_log
								grava_log usuario, "", "", "", OP_LOG_VISANET_OPCOES_PAGTO_ALTERACAO, s_log
								end if
							end if
					else
						erro_fatal=True
						alerta = "FALHA AO GRAVAR OS DADOS (" & Cstr(Err) & ": " & Err.Description & ")."
						end if
					end if
				
			'	PARCELAMENTO PELO EMISSOR DO CARTÃO
				if alerta = "" then
					s = "SELECT * FROM t_PRAZO_PAGTO_VISANET WHERE tipo = '" & CieloObtemIdRegistroBdPrazoPagtoEmissor(vDados(ic).bandeira) & "'"
					if r.State <> 0 then r.Close
					r.Open s, cn
					if r.EOF then
						blnNovoRegistro = True
						r.Addnew
						r("tipo") = CieloObtemIdRegistroBdPrazoPagtoEmissor(vDados(ic).bandeira)
						r("descricao") = "Parcelamento pelo emissor do cartão (com juros)"
					else
						log_via_vetor_carrega_do_recordset r, vLog1, campos_a_omitir
						end if
					
					r("qtde_parcelas")=vDados(ic).intEmissorQtdeParcelas
					r("vl_min_parcela")=vDados(ic).vlEmissorParcelaMinima
					r("atualizacao_data")=Date
					r("atualizacao_hora")=retorna_so_digitos(formata_hora(Now))
					r("atualizacao_usuario")=usuario
					r.Update
					
					If Err = 0 then 
						if blnNovoRegistro then
							log_via_vetor_carrega_do_recordset r, vLog1, campos_a_omitir
							s_log = log_via_vetor_monta_inclusao(vLog1)
							if s_log <> "" then
								grava_log usuario, "", "", "", OP_LOG_VISANET_OPCOES_PAGTO_INCLUSAO, s_log
								end if
						else
							log_via_vetor_carrega_do_recordset r, vLog2, campos_a_omitir
							s_log = log_via_vetor_monta_alteracao(vLog1, vLog2)
							if s_log <> "" then 
								s_log="tipo=" & Trim("" & r("tipo")) & "; " & s_log
								grava_log usuario, "", "", "", OP_LOG_VISANET_OPCOES_PAGTO_ALTERACAO, s_log
								end if
							end if
					else
						erro_fatal=True
						alerta = "FALHA AO GRAVAR OS DADOS (" & Cstr(Err) & ": " & Err.Description & ")."
						end if
					end if
				end if
			next
		
		if alerta = "" then
		'	~~~~~~~~~~~~~~
			cn.CommitTrans
		'	~~~~~~~~~~~~~~
			if Err<>0 then 
				alerta=Cstr(Err) & ": " & Err.Description
				end if
		else
		'	~~~~~~~~~~~~~~~~
			cn.RollbackTrans
		'	~~~~~~~~~~~~~~~~
			end if
		end if
	
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
	<title>CENTRAL ADMINISTRATIVA</title>
	</head>



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


<body onload="bVOLTAR.focus();">
<center>
<br>

<!--  T E L A  -->

<p class="T">A V I S O</p>

<% 
	s = ""
	s_aux="'MtAviso'"
	if alerta <> "" then
		s = "<p style='margin:5px 2px 5px 2px;'>" & alerta & "</p>"
		s_aux="'MtAlerta'"
	else
		s = "DADOS GRAVADOS COM SUCESSO."
		if s <> "" then s="<P style='margin:5px 2px 5px 2px;'>" & s & "</P>"
		end if
%>
<div class=<%=s_aux%> style="width:400px;font-weight:bold;" align="center"><%=s%></div>
<br><br>

<p class="TracoBottom"></p>

<table cellSpacing="0">
<tr>
<%
	s="resumo.asp" & "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))
	if erro_consistencia And (Not erro_fatal) then s="javascript:history.back()"
%>
	<td align="center"><a name="bVOLTAR" id="bVOLTAR" href="<%=s%>"><img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
</tr>
</table>

</center>
</body>
</html>


<%
'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>
