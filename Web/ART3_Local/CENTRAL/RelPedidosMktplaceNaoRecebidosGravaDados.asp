<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =================================================================
'	  RelPedidosMktplaceNaoRecebidosGravaDados.asp
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

	dim s, usuario, msg_erro, s_log, s_log_aux
	s_log = ""

	usuario = Trim(Session("usuario_atual"))
	if (usuario = "") then usuario = Trim(Request("c_usuario_sessao"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO)

	dim alerta
	alerta=""

'	OBTÉM DADOS DO FORMULÁRIO
	dim i, j, n, c, vAux, s_pedido, s_dt_recebimento, dt_recebimento, s_dados_aux, s_dt_pedido
	dim c_dt_entregue_inicio, c_dt_entregue_termino
	dim c_transportadora, c_loja


	dim v_pedidos, v_dt_pedido, qtde_pedidos
	redim v_pedidos(0)
	v_pedidos(Ubound(v_pedidos))=""
    redim v_dt_pedido(0)
    v_dt_pedido(Ubound(v_dt_pedido))="XXXXX"
	qtde_pedidos=0

	c_dt_entregue_inicio = Trim(Request.Form("c_dt_entregue_inicio"))
	c_dt_entregue_termino = Trim(Request.Form("c_dt_entregue_termino"))
	c_transportadora = Trim(Request.Form("c_transportadora"))
	c_loja = retorna_so_digitos(Trim(Request.Form("c_loja")))

	s_dt_recebimento = Request.Form("c_dt_recebimento")
	dt_recebimento = StrToDate(s_dt_recebimento)

	n = Request.Form("ckb_recebido").Count
	for i = 1 to n
		if v_pedidos(Ubound(v_pedidos)) <> "" then
			redim preserve v_pedidos(Ubound(v_pedidos)+1)
			end if
        if v_dt_pedido(Ubound(v_dt_pedido)) <> "XXXXX" then
            redim preserve v_dt_pedido(UBound(v_dt_pedido)+1)
            end if
		s_pedido = Trim(Request.Form("ckb_recebido")(i))
        s_dt_pedido = Trim(Request.Form("c_dt_recebimento_pedido")(i))
		v_pedidos(Ubound(v_pedidos)) = s_pedido
        v_dt_pedido(UBound(v_dt_pedido)) = s_dt_pedido
		qtde_pedidos = qtde_pedidos + 1
		next

	if alerta = "" then
		if (qtde_pedidos = 0) then
			alerta = "Nenhum pedido foi selecionado"
			end if
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
		for i=Lbound(v_pedidos) to Ubound(v_pedidos)
			if (v_pedidos(i) <> "") then
				s = "SELECT * FROM t_PEDIDO WHERE (pedido = '" & v_pedidos(i) & "')"
				if rs.State <> 0 then rs.Close
				rs.Open s, cn
				if rs.Eof then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Pedido não foi encontrado (" & v_pedidos(i) & ")."
				else
					if Trim("" & rs("st_entrega"))<>ST_ENTREGA_ENTREGUE then
						alerta=texto_add_br(alerta)
						alerta=alerta & "Pedido " & v_pedidos(i) & " possui status inválido para a operação 'Recebido': " & Ucase(x_status_entrega(Trim("" & rs("st_entrega"))))
					else
						if CLng(rs("MarketplacePedidoRecebidoRegistrarStatus")) = CLng(COD_ST_PEDIDO_RECEBIDO_SIM) then
							alerta=texto_add_br(alerta)
							alerta=alerta & "Pedido " & v_pedidos(i) & " já está assinalado como pedido Marketplace recebido pelo cliente (" & formata_data(rs("MarketplacePedidoRecebidoRegistrarDataRecebido")) & ")"
						else
							rs("MarketplacePedidoRecebidoRegistrarStatus") = CLng(COD_ST_PEDIDO_RECEBIDO_SIM)
							if v_dt_pedido(i) <> "" then
                                rs("MarketplacePedidoRecebidoRegistrarDataRecebido") = StrToDate(v_dt_pedido(i))
                            else
							    rs("MarketplacePedidoRecebidoRegistrarDataRecebido") = dt_recebimento
                            end if
							rs("MarketplacePedidoRecebidoRegistrarDataHora") = Now
							rs("MarketplacePedidoRecebidoRegistrarUsuario") = usuario
							rs.Update
							if Err <> 0 then
								alerta=texto_add_br(alerta)
								alerta=alerta & Cstr(Err) & ": " & Err.Description
								end if
							end if
						if rs.State <> 0 then rs.Close
						end if
					end if
				end if

		'	SE HOUVE ERRO, CANCELA O LAÇO
			if alerta <> "" then exit for
			next

		if alerta = "" then
			s_log = "Operação=Recebido"
            if s_dt_recebimento <> "" then s_log = s_log & " (Data Padrão=" & s_dt_recebimento & ")"
			
			s_log_aux = ""
			for i=Lbound(v_pedidos) to Ubound(v_pedidos)
				if v_pedidos(i) <> "" then
				'	INFORMAÇÕES PARA O LOG
					if s_log_aux <> "" then s_log_aux = s_log_aux & ", "
					s_log_aux = s_log_aux & v_pedidos(i)
                    if v_dt_pedido(i) <> "" then s_log_aux = s_log_aux & " (" & v_dt_pedido(i) & ")"
					end if
				next
			
			if s_log_aux <> "" then
				if s_log <> "" then s_log = s_log & "; "
				s_log = s_log & "Pedidos=" & s_log_aux
				end if
			
			grava_log usuario, "", "", "", "MKT PED RECEBIDO", s_log
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
    f.action = "RelPedidosMktplaceNaoRecebidos.asp?url_back=X";
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
<input type="hidden" name="pedido_selecionado" id="pedido_selecionado" value=''>
<input type="hidden" name="c_dt_entregue_inicio" id="c_dt_entregue_inicio" value="<%=c_dt_entregue_inicio%>">
<input type="hidden" name="c_dt_entregue_termino" id="c_dt_entregue_termino" value="<%=c_dt_entregue_termino%>">
<input type="hidden" name="c_transportadora" id="c_transportadora" value="<%=c_transportadora%>">
<input type="hidden" name="c_loja" id="c_loja" value="<%=c_loja%>">


<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black;">

<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Pedidos de Marketplace Não Recebidos Pelo Cliente<span class="C">&nbsp;</span></span>
		<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span>
	</td>
</tr>
</table>
<br>
<br>

<%if qtde_pedidos > 0 then %>
<!-- ************   MENSAGEM  ************ -->
<% 
	s = ""
	for i=Lbound(v_pedidos) to Ubound(v_pedidos)
		if v_pedidos(i) <> "" then
			if s <> "" then s = s & ", "
			s = s & v_pedidos(i)
			end if
		next
	
	if s = "" then s = "nenhum pedido selecionado"
%>
<div class='MtAviso' style="width:649px;font-weight:bold;border:1pt solid black;" align="center"><span style='margin:5px 2px 5px 2px;'>Pedidos Marcados como Recebidos:</span><br /><span><%=s%></span></div>
<br>
<% end if %>

<!-- ************   SEPARADOR   ************ -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black;">
<tr><td class="Rc" align="left">&nbsp;</td></tr>
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