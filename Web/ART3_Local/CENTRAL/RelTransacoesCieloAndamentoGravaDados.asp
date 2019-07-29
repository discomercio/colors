<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =================================================================
'	  RelTransacoesCieloAndamentoGravaDados.asp
'     =================================================================
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

	class cl_TIPO_GRAVA_REL_CIELO_ANDAMENTO
		dim id_registro
		dim pedido
		dim tid
		dim obs
		end class

	Const COD_MANUAL_NAO_TRATADO = 0
	Const COD_MANUAL_TRATADO = 1

	dim s, usuario, msg_erro, s_log

	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO)

	dim alerta
	alerta=""

'	OBTÉM FILTROS
	dim c_dt_inicio, c_dt_termino, c_resultado_transacao, c_bandeira, c_pedido, c_cliente_cnpj_cpf, c_loja, rb_ordenacao_saida

	c_dt_inicio = Trim(Request("c_dt_inicio"))
	c_dt_termino = Trim(Request("c_dt_termino"))
	c_resultado_transacao = Trim(Request("c_resultado_transacao"))
	c_bandeira = Trim(Request("c_bandeira"))
	c_pedido = Trim(Request("c_pedido"))
	c_cliente_cnpj_cpf = retorna_so_digitos(Trim(Request("c_cliente_cnpj_cpf")))
	c_loja = retorna_so_digitos(Trim(Request("c_loja")))
	rb_ordenacao_saida = Trim(Request("rb_ordenacao_saida"))
	
	s = normaliza_num_pedido(c_pedido)
	if s <> "" then c_pedido = s
	
'	OBTÉM DADOS DO FORMULÁRIO
	dim i, n, s_dados, s_id_registro, vAux, s_pedido, s_tid, s_texto_obs
	dim intNsu

'	CHECK BOX P/ INDICAR A GRAVAÇÃO DA INFORMAÇÃO DE QUE A TRANSAÇÃO FOI TRATADA
'	CAMPO TEXTO P/ INCLUIR OBSERVAÇÕES DA TRANSAÇÃO TRATADA
	dim v_transacao_tratada, qtde_transacao_tratada
	redim v_transacao_tratada(0)
	set v_transacao_tratada(Ubound(v_transacao_tratada)) = new cl_TIPO_GRAVA_REL_CIELO_ANDAMENTO
	v_transacao_tratada(Ubound(v_transacao_tratada)).id_registro=""
	qtde_transacao_tratada=0
	
	n = Request.Form("ckb_tratado").Count
	
	for i = 1 to n
		s_dados = Trim(Request.Form("ckb_tratado")(i))
		
		if s_dados <> "" then
			vAux=Split(s_dados, "|")
			s_id_registro = Trim(vAux(LBound(vAux)))
			s_pedido = Trim(vAux(LBound(vAux)+1))
			s_tid = Trim(vAux(LBound(vAux)+2))
			s_texto_obs = Trim(Request.Form("c_tratado_manual_obs_" & s_id_registro))
			if v_transacao_tratada(Ubound(v_transacao_tratada)).id_registro <> "" then
				redim preserve v_transacao_tratada(Ubound(v_transacao_tratada)+1)
				set v_transacao_tratada(Ubound(v_transacao_tratada)) = new cl_TIPO_GRAVA_REL_CIELO_ANDAMENTO
				end if
			v_transacao_tratada(Ubound(v_transacao_tratada)).id_registro = s_id_registro
			v_transacao_tratada(Ubound(v_transacao_tratada)).pedido = s_pedido
			v_transacao_tratada(Ubound(v_transacao_tratada)).tid = s_tid
			v_transacao_tratada(Ubound(v_transacao_tratada)).obs = s_texto_obs
			qtde_transacao_tratada = qtde_transacao_tratada + 1
			end if
		next

	if alerta = "" then
		if (qtde_transacao_tratada = 0) then
			alerta = "Nenhuma transação foi assinalada para ser marcada como já tratada."
			end if
		end if
	
	if alerta = "" then
		for i=Lbound(v_transacao_tratada) to UBound(v_transacao_tratada)
			if v_transacao_tratada(i).id_registro <> "" then
				if Len(v_transacao_tratada(i).obs) > MAX_TAM_T_PEDIDO_PAGTO_CIELO__TRATADO_MANUAL_OBS then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Texto com observações para a transação do pedido " & v_transacao_tratada(i).pedido & " excede o tamanho máximo."
					end if
				end if
			next
		end if


'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	
	if alerta = "" then
	'	~~~~~~~~~~~~~
		cn.BeginTrans
	'	~~~~~~~~~~~~~
		If Not cria_recordset_pessimista(rs, msg_erro) then
		'	~~~~~~~~~~~~~~~~
			cn.RollbackTrans
		'	~~~~~~~~~~~~~~~~
			Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
			end if

	'	GRAVA DADOS
	'	===========
		for i=Lbound(v_transacao_tratada) to Ubound(v_transacao_tratada)
			if (v_transacao_tratada(i).id_registro <> "") then
				s = "SELECT * FROM t_PEDIDO_PAGTO_CIELO WHERE (id = " & v_transacao_tratada(i).id_registro & ")"
				if rs.State <> 0 then rs.Close
				rs.Open s, cn
				if rs.Eof then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Registro da transação não foi encontrado (ID=" & v_transacao_tratada(i).id_registro & ")."
				else
					rs("tratado_manual_status") = COD_MANUAL_TRATADO
					rs("tratado_manual_usuario") = usuario
					rs("tratado_manual_data_hora") = Now
					rs("tratado_manual_obs") = Trim(v_transacao_tratada(i).obs)
					
					rs.Update
					
					if alerta = "" then
					'	INFORMAÇÕES PARA O LOG
						if s_log <> "" then s_log = s_log & ", "
						s_log = s_log & v_transacao_tratada(i).pedido & " (TID=" & v_transacao_tratada(i).tid & "; t_PEDIDO_PAGTO_CIELO.id=" & v_transacao_tratada(i).id_registro & ")"
						end if
					end if
				end if
			
		'	SE HOUVE ERRO, CANCELA O LAÇO
			if alerta <> "" then exit for
			next

		if alerta = "" then
			if s_log <> "" then
				s_log = "Transações Cielo em andamento marcadas como já tratadas: " & s_log
				grava_log usuario, "", "", "", OP_LOG_PAGTO_CIELO_EM_ANDAMENTO_MARCAR_COMO_TRATADO, s_log
				end if
			end if


	'	FINALIZA TRANSAÇÃO
	'	==================
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
	<title>CENTRAL</title>
	</head>



<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
function fRetornar(f) {
	f.action = "RelTransacoesCieloAndamentoExec.asp?url_back=X";
	f.submit();
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

<link href="<%=URL_FILE__E_CSS%>" rel="stylesheet" type="text/css">

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
<body onload="window.status='Concluído';bVOLTAR.focus();" link=#000000 alink=#000000 vlink=#000000>

<center>

<form id="f" name="f" method="post">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<!-- FILTROS -->
<input type="hidden" name="c_dt_inicio" id="c_dt_inicio" value="<%=c_dt_inicio%>" />
<input type="hidden" name="c_dt_termino" id="c_dt_termino" value="<%=c_dt_termino%>" />
<input type="hidden" name="c_resultado_transacao" id="c_resultado_transacao" value="<%=c_resultado_transacao%>" />
<input type="hidden" name="c_bandeira" id="c_bandeira" value="<%=c_bandeira%>" />
<input type="hidden" name="c_pedido" id="c_pedido" value="<%=c_pedido%>" />
<input type="hidden" name="c_cliente_cnpj_cpf" id="c_cliente_cnpj_cpf" value="<%=c_cliente_cnpj_cpf%>" />
<input type="hidden" name="c_loja" id="c_loja" value="<%=c_loja%>" />
<input type="hidden" name="rb_ordenacao_saida" id="rb_ordenacao_saida" value="<%=rb_ordenacao_saida%>" />

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black;">

<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Transações Cielo em Andamento<span class="C">&nbsp;</span></span></td>
</tr>
</table>
<br>
<br>

<%if qtde_transacao_tratada > 0 then %>
<!-- ************   MENSAGEM  ************ -->
<% 
	s = ""
	for i=Lbound(v_transacao_tratada) to Ubound(v_transacao_tratada)
		if v_transacao_tratada(i).pedido <> "" then
			if s <> "" then s = s & "<br />"
			s = s & v_transacao_tratada(i).pedido & " (TID: " & v_transacao_tratada(i).tid & ")"
			end if
		next
	
	if s = "" then s = "nenhuma transação selecionada"
%>
<div class='MtAviso' style="width:649px;font-weight:bold;border:1pt solid black;" align="center"><p style='margin:5px 2px 5px 2px;'>Transações Cielo em andamento marcadas como já tratadas:<br /> <%=s%></p></div>
<br>
<% end if %>

<!-- ************   SEPARADOR   ************ -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black;">
<tr><td class="Rc" align="left">&nbsp;</td></tr>
</table>

<!-- ************   LINKS: PÁGINA INICIAL / ENCERRA SESSÃO   ************ -->
<table width="649" cellpadding="0" cellspacing="0">
<tr><td align="right"><span class="Rc">
	<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
	<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
	</span>
</td></tr>
</table>

<!-- ************   BOTÕES   ************ -->
<table class="notPrint" width="649" cellspacing="0">
<tr>
	<td align="center"><a name="bVOLTAR" id="bVOLTAR" href="javascript:fRetornar(f)" title="Retornar para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
</tr>
</table>
</form>

</center>
</body>

<% end if %>

</html>


<%
'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>