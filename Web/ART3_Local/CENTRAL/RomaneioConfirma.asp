<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/estoque.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =================================================================
'	  R O M A N E I O C O N F I R M A . A S P
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

	dim s, strScript, usuario, msg_erro, s_log, s_log_aux, s_log_transp, s_transportadora, lista_pedidos, v_pedido, i, achou
	dim c_nsu_romaneio, c_num_coleta, c_dt_entrega, c_transportadora_contato, c_conferente, c_motorista, c_placa_veiculo, c_nfe_emitente

	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	OBTÉM DADOS DO FORMULÁRIO
	s_transportadora = Trim(request("c_transportadora"))
	lista_pedidos = ucase(Trim(request("pedidos_selecionados")))
	if (lista_pedidos = "") then Response.Redirect("aviso.asp?id=" & ERR_PEDIDO_NAO_ESPECIFICADO)

	dim lngNsuWmsRomaneioN1
	c_nsu_romaneio = Trim(Request.Form("c_nsu_romaneio"))
	lngNsuWmsRomaneioN1 = converte_numero(c_nsu_romaneio)
	
	c_num_coleta = Trim(Request.Form("c_num_coleta"))
	c_dt_entrega = Trim(Request.Form("c_dt_entrega"))
	c_transportadora_contato = Trim(Request.Form("c_transportadora_contato"))
	c_conferente = Trim(Request.Form("c_conferente"))
	c_motorista = Trim(Request.Form("c_motorista"))
	c_placa_veiculo = Trim(Request.Form("c_placa_veiculo"))
	c_nfe_emitente = Trim(Request.Form("c_nfe_emitente"))
	
	lista_pedidos=substitui_caracteres(lista_pedidos,chr(10),"")
	v_pedido = split(lista_pedidos,chr(13),-1)
	achou=False
	for i=Lbound(v_pedido) to Ubound(v_pedido)
		if Trim(v_pedido(i))<>"" then
			achou = True
			s = normaliza_num_pedido(v_pedido(i))
			if s <> "" then v_pedido(i) = s
			end if
		next

	if Not achou then Response.Redirect("aviso.asp?id=" & ERR_PEDIDO_NAO_ESPECIFICADO)

	dim alerta
	alerta=""

'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	
	s_log = ""
	s_log_transp = ""
	
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

		for i=Lbound(v_pedido) to Ubound(v_pedido)
			if v_pedido(i) <> "" then
				s = "SELECT * FROM t_PEDIDO WHERE (pedido = '" & v_pedido(i) & "')"
				if rs.State <> 0 then rs.Close
				rs.Open s, cn
				if rs.Eof then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Pedido " & v_pedido(i) & " não foi encontrado."
				else
					if Not IsPedidoRomaneioPossivel(Trim("" & rs("st_entrega"))) then
						alerta=texto_add_br(alerta)
						alerta=alerta & "Pedido " & v_pedido(i) & " possui status inválido para esta operação: " & Ucase(x_status_entrega(Trim("" & rs("st_entrega"))))
					else
						'ROMANEIO GERADO
						rs("romaneio_status")=CInt(COD_ROMANEIO_STATUS__OK)
						rs("romaneio_data")=Date
						rs("romaneio_data_hora")=Now
						rs("romaneio_usuario")=usuario
						
					'	LIMPA DADOS DA SELEÇÃO AUTOMÁTICA DE TRANSPORTADORA BASEADO NO CEP
					'	MANTÉM OS DADOS ANTERIORES (SE HOUVER) P/ FINS DE HISTÓRICO/LOG DOS SEGUINTES CAMPOS:
					'	'transportadora_selecao_auto_cep', 'transportadora_selecao_auto_tipo_endereco' e 'transportadora_selecao_auto_transportadora'
						if Ucase(Trim("" & rs("transportadora_id"))) <> Ucase(Trim(s_transportadora)) then
							if s_log_transp <> "" then s_log_transp = s_log_transp & "; "
							s_log_transp = s_log_transp & v_pedido(i) & ": '" & Ucase(Trim("" & rs("transportadora_id"))) & "' -> '" & Ucase(Trim(s_transportadora)) & "'"
							rs("transportadora_selecao_auto_status") = TRANSPORTADORA_SELECAO_AUTO_STATUS_FLAG_N
							rs("transportadora_selecao_auto_data_hora") = Now
							end if
						
						'TRANSPORTADORA
						rs("transportadora_id")=s_transportadora
						rs("transportadora_data")=Now
						rs("transportadora_usuario")=usuario
						rs("transportadora_num_coleta")=c_num_coleta
						rs("transportadora_contato")=c_transportadora_contato
						rs("transportadora_conferente")=c_conferente
						rs("transportadora_motorista")=c_motorista
						rs("transportadora_placa_veiculo")=c_placa_veiculo
						'DATA DE COLETA (RÓTULO ANTIGO: ENTREGA MARCADA PARA)
						rs("a_entregar_status")=1
						rs("a_entregar_data_marcada")=StrToDate(c_dt_entrega)
						rs("a_entregar_data")=Date
						rs("a_entregar_hora")=retorna_so_digitos(formata_hora(Now))
						rs("a_entregar_usuario")=usuario
						rs.Update
						if Err <> 0 then 
							alerta=texto_add_br(alerta)
							alerta=alerta & Cstr(Err) & ": " & Err.Description
							end if
					'	INFORMAÇÕES PARA O LOG
						if s_log <> "" then s_log = s_log & ", "
						s_log = s_log & v_pedido(i)
						if rs.State <> 0 then rs.Close
						end if
					end if
				end if
				
		'	SE HOUVE ERRO, CANCELA O LAÇO
			if alerta <> "" then exit for
			next
		
		if alerta = "" then
			s_log_aux = ""
			if c_num_coleta <> "" then s_log_aux = s_log_aux & " Nº Coleta: " & c_num_coleta & ";"
			if c_dt_entrega <> "" then s_log_aux = s_log_aux & " Data de coleta: " & c_dt_entrega & ";"
			if c_transportadora_contato <> "" then s_log_aux = s_log_aux & " Contato: " & c_transportadora_contato & ";"
			if c_conferente <> "" then s_log_aux = s_log_aux & " Conferente: " & c_conferente & ";"
			if c_motorista <> "" then s_log_aux = s_log_aux & " Motorista: " & c_motorista & ";"
			if c_placa_veiculo <> "" then s_log_aux = s_log_aux & " Placa do veículo: " & c_placa_veiculo & ";"
			
			s_log = "Romaneio de entrega (NSU=" & normaliza_a_esq(Cstr(lngNsuWmsRomaneioN1), 3) & "; CD=" & obtem_apelido_empresa_NFe_emitente(c_nfe_emitente) & ") p/ a transportadora " & s_transportadora & ";" & s_log_aux & " Pedido(s) = " & s_log & "; Log de alteração da transportadora: " & s_log_transp
			grava_log usuario, "", "", "", OP_LOG_PEDIDO_ALTERACAO, s_log
		'	~~~~~~~~~~~~~~
			cn.CommitTrans
		'	~~~~~~~~~~~~~~
			if Err<>0 then 
				alerta=Cstr(Err) & ": " & Err.Description
				end if
			
			if alerta = "" then
			'	EXIBE A MENSAGEM DE OPERAÇÃO BEM SUCEDIDA EM UMA PÁGINA APENAS DE EXIBIÇÃO
			'	ISSO É FEITO P/ EVITAR QUE O USUÁRIO EXECUTE UM COMANDO DE ATUALIZAR PÁGINA (F5) QUE CAUSARIA UMA NOVA GRAVAÇÃO DOS DADOS
				Session(SESSION_CLIPBOARD) = "Pedidos atualizados com sucesso"
				Response.Redirect("mensagem.asp?url_back=" & server.URLEncode("RomaneioPreFiltro.asp") & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")))
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


<%=DOCTYPE_LEGADO%>


<html>


<head>
	<title>CENTRAL</title>
	</head>



<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>




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
<body onload="window.status='Concluído';bVOLTAR.focus();" link=#000000 alink=#000000 vlink=#000000>

<center>

<form id="f" name="f" method="post" action="RomaneioConfirma.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="pedidos_selecionados" id="pedidos_selecionados" value="<%=lista_pedidos%>">
<input type="hidden" name="c_transportadora" id="c_transportadora" value="<%=s_transportadora%>">
<input type="hidden" name="c_nfe_emitente" id="c_nfe_emitente" value="<%=c_nfe_emitente%>" />


<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">

<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Romaneio de Entrega<span class="C">&nbsp;</span></span></td>
</tr>
</table>
<br>
<br>


<!-- ************   MENSAGEM  ************ -->
<div class='MtAviso' style="width:649px;font-weight:bold;border:1pt solid black;" align="center"><span style='margin:5px 2px 5px 2px;'>Pedidos atualizados com sucesso</span></div>
<br>

	
<!-- ************   SEPARADOR   ************ -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc" align="left">&nbsp;</td></tr>
</table>

<!-- ************   LINKS: PÁGINA INICIAL / ENCERRA SESSÃO   ************ -->
<table width="649" cellPadding="0" CellSpacing="0">
<tr><td align="right"><span class="Rc">
	<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
	<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
	</span>
</td></tr>
</table>

<!-- ************   BOTÕES   ************ -->
<table class="notPrint" width="649" cellSpacing="0">
<tr>
	<td align="center"><a name="bVOLTAR" id="bVOLTAR" href="RomaneioPreFiltro.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="Retornar">
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