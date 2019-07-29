<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp"        -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ===========================================================================
'	  CadPercMaxDescCadLojaConfirma.asp
'     ===========================================================================
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

	if Not operacao_permitida(OP_CEN_CAD_PARAMETROS_GLOBAIS, s_lista_operacoes_permitidas) then
		Response.Redirect("aviso.asp?id=" & ERR_NIVEL_ACESSO_INSUFICIENTE)
		end if

	Dim s_log
	s_log = ""
	
'	OBTÉM DADOS DO FORMULÁRIO ANTERIOR
	class cl_CadPercMaxDescCadLoja
		dim num_loja
		dim nome_loja
		dim percentual
		end class

	dim vTabela
	redim vTabela(0)
	set vTabela(Ubound(vTabela)) = new cl_CadPercMaxDescCadLoja
	
	dim i, n
	n = Request.Form("c_loja").Count
	for i = 1 to n
		s = Trim(Request.Form("c_loja")(i))
		if s <> "" then
			if Trim(vTabela(Ubound(vTabela)).num_loja) <> "" then
				redim preserve vTabela(Ubound(vTabela)+1)
				set vTabela(Ubound(vTabela)) = new cl_CadPercMaxDescCadLoja
				end if
			vTabela(Ubound(vTabela)).num_loja = Trim(Request.Form("c_loja")(i))
			vTabela(Ubound(vTabela)).nome_loja = Trim(Request.Form("c_nome_loja")(i))
			vTabela(Ubound(vTabela)).percentual = converte_numero(Trim(Request.Form("c_percentual")(i)))
			end if
		next
	
	dim erro_consistencia, erro_fatal
	
	erro_consistencia=false
	erro_fatal=false
	
	alerta = ""
	for i = Lbound(vTabela) to Ubound(vTabela)
		if Trim("" & vTabela(i).num_loja) <> "" then
			if vTabela(i).percentual < 0 then
				alerta=texto_add_br(alerta)
				alerta=alerta & "PERCENTUAL NEGATIVO NÃO É VÁLIDO (LOJA " & vTabela(i).num_loja & ")"
			elseif vTabela(i).percentual > 100 then
				alerta=texto_add_br(alerta)
				alerta=alerta & "PERCENTUAL NÃO PODE EXCEDER 100% (LOJA " & vTabela(i).num_loja & ")"
				end if
			end if
		next
	
	if alerta <> "" then erro_consistencia=True	
		
	Err.Clear
	
	dim msg_erro
	if Not cria_recordset_otimista(r, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

'	GRAVA OS DADOS!!
	if alerta = "" then
	'	~~~~~~~~~~~~~
		cn.BeginTrans
	'	~~~~~~~~~~~~~
		for i=Lbound(vTabela) to Ubound(vTabela)
			if Trim("" & vTabela(i).num_loja) <> "" then
				s = "SELECT * FROM t_LOJA WHERE (loja = '" & vTabela(i).num_loja & "')"
				if r.State <> 0 then r.Close
				r.Open s, cn
				if r.Eof then
					alerta = texto_add_br(alerta)
					alerta = alerta & "LOJA '" & vTabela(i).num_loja & "' NÃO FOI ENCONTRADA NO BANCO DE DADOS."
					end if
				
				if alerta = "" then
					if converte_numero(r("PercMaxSenhaDesconto")) <> converte_numero(vTabela(i).percentual) then
						if s_log <> "" then s_log = s_log & "; "
						s_log = s_log & vTabela(i).num_loja & ": " & formata_perc(r("PercMaxSenhaDesconto")) & " => " & formata_perc(vTabela(i).percentual)
						r("PercMaxSenhaDesconto") = vTabela(i).percentual
						r.Update
						
						if Err <> 0 then
							erro_fatal=True
							alerta = "FALHA AO GRAVAR OS DADOS (" & Cstr(Err) & ": " & Err.Description & ")."
							end if
						end if
					end if
				end if
			next

		if r.State <> 0 then r.Close
		set r = nothing
		
		if alerta = "" then
			if s_log <> "" then
				s_log = "Alteração do percentual máximo da senha de desconto para cadastramento na loja: " & s_log
				grava_log usuario, "", "", "", OP_LOG_PERC_MAX_SENHA_DESC_CADASTRADO_NA_LOJA, s_log
				end if
			end if
		
		if alerta = "" then
		'	~~~~~~~~~~~~~~
			cn.CommitTrans
		'	~~~~~~~~~~~~~~
			if Err <> 0 then alerta=Cstr(Err) & ": " & Err.Description
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
		s = "<P style='margin:5px 2px 5px 2px;'>" & alerta & "</P>"
		s_aux="'MtAlerta'"
	else
		s = "DADOS ALTERADOS COM SUCESSO."
		if s <> "" then s="<P style='margin:5px 2px 5px 2px;'>" & s & "</P>"
		end if
%>
<div class=<%=s_aux%> style="width:400px;FONT-WEIGHT:bold;" align="CENTER"><%=s%></div>
<BR><BR>

<p class="TracoBottom"></p>

<table cellSpacing="0">
<tr>
<%
	s="resumo.asp" & "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))
	if erro_consistencia And (Not erro_fatal) then s="javascript:history.back()"
%>
	<td align="CENTER"><a name="bVOLTAR" id="bVOLTAR" href="<%=s%>"><img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
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
