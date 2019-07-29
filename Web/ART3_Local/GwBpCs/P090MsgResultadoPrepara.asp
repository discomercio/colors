<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =============================================
'	  P090MsgResultadoPrepara.asp
'     =============================================
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

'	OBS: A P�GINA QUE EXIBE A MENSAGEM SOBRE O RESULTADO DA TRANSA��O � ACIONADA
'	~~~~ ATRAV�S DOS SEGUINTES PASSOS:
'	1) A P�GINA QUE EXECUTA A TRANSA��O C/ A BRASPAG VIA WEB SERVICE ARMAZENA OS DADOS
'		RECEBIDOS NO BD E, EM SEGUIDA, ENCAMINHA P/ ESTA P�GINA INTERMEDI�RIA INFORMANDO
'		O ID DO REGISTRO.
'	2) A P�GINA INTERMEDI�RIA PREPARA OS DADOS EM CAMPOS HIDDEN DE UM FORM, L� E APAGA OS
'		DADOS ARMAZENADOS ATRAV�S DA SESSION E, POR FIM, FAZ UM SUBMIT() P/ A P�GINA
'		FINAL DE EXIBI��O.
'	3) COM ESTE MECANISMO, SE O USU�RIO ACIONAR O REFRESH NA P�GINA DE EXIBI��O, EVITAM-SE
'		OS SEGUINTES PROBLEMAS:
'		A) REEXECUTAR O PROCESSAMENTO DA TRANSA��O.
'		B) PARA OS DADOS ARMAZENADOS NA SESSION, A PARTIR DA 2� EXECU��O OS DADOS J� TERIAM
'			SIDO APAGADOS.

	On Error GoTo 0
	Err.Clear

	dim alerta
	alerta = ""

	dim pedido_selecionado, id_pedido_base
	pedido_selecionado = Trim(Request("pedido"))
	if pedido_selecionado = "" then Response.Redirect("aviso.asp?id=" & ERR_PEDIDO_NAO_ESPECIFICADO)
	id_pedido_base = retorna_num_pedido_base(pedido_selecionado)
	
	dim strIdPagtoGwPag
	strIdPagtoGwPag = Trim(Request("idPagtoGwPag"))
	if strIdPagtoGwPag <> "" then strIdPagtoGwPag = decriptografa(strIdPagtoGwPag)
	if Trim("" & strIdPagtoGwPag) = "" then Response.Redirect("aviso.asp?id=" & ERR_ID_NAO_INFORMADO)

	dim cn, msg_erro
	dim t_PAG
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(t_PAG, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

	dim s
	if alerta = "" then
		s = "SELECT * FROM t_PAGTO_GW_PAG WHERE (id = " & strIdPagtoGwPag & ")"
		if t_PAG.State <> 0 then t_PAG.Close
		t_PAG.Open s, cn
		if Not t_PAG.Eof then
			alerta = Trim("" & t_PAG("msg_alerta_tela"))
			end if
		if t_PAG.State <> 0 then t_PAG.Close
		end if

'	FILTRAGEM DE ASPAS P/ N�O CAUSAR ERRO AO CARREGAR NO CAMPO VALUE DE ELEMENTO INPUT TEXT
	if alerta <> "" then alerta = substitui_caracteres(alerta, chr(34), "'")
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
	<title>LOJA</title>
	</head>


<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
	setTimeout('fBraspag.submit()', 100);
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


<body>
<center>

<form id="fBraspag" name="fBraspag" method="post" action="P100MsgResultadoExibe.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="pedido_selecionado" value="<%=pedido_selecionado%>" />
<input type="hidden" name="idPagtoGwPag" value="<%=strIdPagtoGwPag%>" />
<input type="hidden" name="alerta" value="<%=alerta%>" />
</form>

<b>Aguarde, redirecionando para exibir mensagem...</b>
<br />
<a name="bREDIRECIONA" id="bREDIRECIONA" href="javascript:fBraspag.submit();"><b>Se o redirecionamento n�o ocorrer automaticamente, clique aqui.</b></a>

</center>
</body>

</html>

<%
	if t_PAG.State <> 0 then t_PAG.Close
	set t_PAG=nothing

'	FECHA CONEX�O
	cn.Close
	set cn = nothing
%>