<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =====================================================
'	  PedidoRecebidoConsiste.asp
'     =====================================================
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

	dim s, usuario, lista_pedidos, v_aux, i, j, n, achou_pedido, achou_obs2, msg_erro
	dim lista_obs2, s_filtro_obs2
	dim ckb_entrega, ckb_recebido, c_dt_recebido, c_transportadora
	dim strTransportadoraPedido, strTransportadoraFiltro, strAux
	dim vConsiste, strObs2Digitado
	dim lista_pedidos_selecionados, vPedidosSelecionados
	dim blnFlagOk
	
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

	dim alerta
	alerta = ""

	dim observacoes
	observacoes = ""
	
'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	
	dim blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos
	blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos = isActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos

	ckb_entrega = Trim(Request.Form("ckb_entrega"))
	ckb_recebido = Trim(Request.Form("ckb_recebido"))
	c_dt_recebido = Trim(Request.Form("c_dt_recebido"))
	if ckb_recebido = "" then 
		c_dt_recebido = ""
	else
		if c_dt_recebido = "" then c_dt_recebido = formata_data(Date)
		end if
	
	c_transportadora = Trim(Request.Form("c_transportadora"))
	lista_obs2 = Trim(request("c_obs2"))
	lista_pedidos = ucase(Trim(request("c_pedidos")))

	dim c_nfe_emitente
	c_nfe_emitente = Trim(Request.Form("c_nfe_emitente"))

	if (lista_pedidos = "") And (lista_obs2 = "") then Response.Redirect("aviso.asp?id=" & ERR_PEDIDO_NAO_ESPECIFICADO)
	
	if alerta = "" then
		if c_nfe_emitente = "" then
			alerta=texto_add_br(alerta)
			alerta = alerta & "Não foi informado o CD"
		elseif converte_numero(c_nfe_emitente) = 0 then
			alerta=texto_add_br(alerta)
			alerta = alerta & "É necessário definir um CD válido"
			end if
		end if

	dim rNfeEmitente
	set rNfeEmitente = le_nfe_emitente(c_nfe_emitente)
	
'	LISTA COM Nº PEDIDOS
	lista_pedidos=substitui_caracteres(lista_pedidos,chr(10),"")
	v_aux = split(lista_pedidos,chr(13),-1)
	achou_pedido=False
	for i=Lbound(v_aux) to Ubound(v_aux)
		if Trim(v_aux(i))<>"" then
			achou_pedido = True
			s = normaliza_num_pedido(v_aux(i))
			if s <> "" then v_aux(i) = s
			end if
		next

	redim vConsiste(0)
	set vConsiste(Ubound(vConsiste)) = New cl_ANOTA_PEDIDO_RECEBIDO
	vConsiste(Ubound(vConsiste)).pedido_digitado=""
	vConsiste(Ubound(vConsiste)).obs2_digitado=""
	for i = Lbound(v_aux) to Ubound(v_aux)
		if Trim(v_aux(i))<>"" then
			if (Trim(vConsiste(Ubound(vConsiste)).pedido_digitado)<>"") Or _
			   (Trim(vConsiste(Ubound(vConsiste)).obs2_digitado)<>"") then
				redim preserve vConsiste(Ubound(vConsiste)+1)
				set vConsiste(Ubound(vConsiste)) = New cl_ANOTA_PEDIDO_RECEBIDO
				end if
			with vConsiste(Ubound(vConsiste))
				.pedido = Trim(v_aux(i))
				.pedido_digitado = Trim(v_aux(i))
				.obs2_digitado = ""
				end with
			end if
		next

	for i=Lbound(v_aux) to Ubound(v_aux)
		if v_aux(i)<>"" then
			for j=Lbound(v_aux) to (i-1)
				if v_aux(i) = v_aux(j) then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Pedido " & v_aux(i) & ": linha " & renumera_com_base1(Lbound(v_aux),i) & " repete o mesmo pedido da linha " & renumera_com_base1(Lbound(v_aux),j) & "."
					exit for
					end if
				next
			end if
		next
	
'	LISTA COM OBS2
	lista_obs2=substitui_caracteres(lista_obs2,chr(10),"")
	v_aux = split(lista_obs2,chr(13),-1)
	achou_obs2=False
	for i=Lbound(v_aux) to Ubound(v_aux)
		if Trim(v_aux(i))<>"" then
			achou_obs2 = True
			s = retorna_so_digitos(v_aux(i))
			if s <> "" then v_aux(i) = s
			end if
		next

	for i = Lbound(v_aux) to Ubound(v_aux)
		if Trim(v_aux(i))<>"" then
			if (Trim(vConsiste(Ubound(vConsiste)).pedido_digitado)<>"") Or _
			   (Trim(vConsiste(Ubound(vConsiste)).obs2_digitado)<>"") then
				redim preserve vConsiste(Ubound(vConsiste)+1)
				set vConsiste(Ubound(vConsiste)) = New cl_ANOTA_PEDIDO_RECEBIDO
				end if
			with vConsiste(Ubound(vConsiste))
				.pedido = ""
				.pedido_digitado = ""
				.obs2_digitado = Trim(v_aux(i))
				end with
			end if
		next

	for i=Lbound(v_aux) to Ubound(v_aux)
		if v_aux(i)<>"" then
			for j=Lbound(v_aux) to (i-1)
				if v_aux(i) = v_aux(j) then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Obs II = " & v_aux(i) & ": linha " & renumera_com_base1(Lbound(v_aux),i) & " repete o mesmo número da linha " & renumera_com_base1(Lbound(v_aux),j) & "."
					exit for
					end if
				next
			end if
		next
	
'	NÃO HÁ Nº PEDIDO E NEM DE OBS2 VÁLIDOS
	if (Not achou_pedido) And (Not achou_obs2) then Response.Redirect("aviso.asp?id=" & ERR_PEDIDO_NAO_ESPECIFICADO)

'	PESQUISA PELO Nº OBS2 E OBTÉM O Nº PEDIDO
	if alerta = "" then
		for i = Lbound(vConsiste) to Ubound(vConsiste)
			if (Trim(vConsiste(i).obs2_digitado) <> "") And ((Trim(vConsiste(i).pedido_digitado) = "")) then
				strObs2Digitado = Trim(vConsiste(i).obs2_digitado)
				'Como o campo 'obs_2' armazena números em formato texto, tenta
				'maximizar a capacidade de pesquisa p/ os casos em que foram
				'cadastrados zeros à esquerda
				s_filtro_obs2 = "'" & strObs2Digitado & "'"
				for j = (Len(strObs2Digitado)+1) to MAX_OBS_2 'Tamanho do campo no BD
					s_filtro_obs2 = s_filtro_obs2 & ","
					n = MAX_OBS_2 - j + 1
					s_filtro_obs2 = s_filtro_obs2 & "'" & String(n,"0") & strObs2Digitado & "'"
					next
				s = "SELECT" & _
						" pedido," & _
						" data,"

				if blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos then
					s = s & _
						" t_PEDIDO.endereco_nome AS nome"
				else
					s = s & _
						" nome"
					end if

				s = s & _
					" FROM t_PEDIDO INNER JOIN t_CLIENTE" & _
						" ON (t_PEDIDO.id_cliente=t_CLIENTE.id)" & _
					" WHERE" & _
						" (obs_2 IN (" & s_filtro_obs2 & "))" & _
						" AND (data >= " & bd_monta_data(Date-365) & ")" & _
					" ORDER BY" & _
						" data," & _
						" pedido"
				if rs.State <> 0 then rs.Close
				rs.open s, cn
				if rs.Eof then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Nenhum pedido encontrado com Obs II = " & strObs2Digitado
				else
					if rs.RecordCount > 1 then
						alerta=texto_add_br(alerta)
						alerta=alerta & Cstr(rs.RecordCount) & " pedidos encontrados com Obs II = " & strObs2Digitado
					else
						vConsiste(i).pedido = Trim("" & rs("pedido"))
						end if
					end if
				end if
			next
		end if

	if alerta = "" then
		for i=Lbound(vConsiste) to Ubound(vConsiste)
			if vConsiste(i).pedido<>"" then
				for j=Lbound(vConsiste) to (i-1)
					if vConsiste(i).pedido = vConsiste(j).pedido then
						alerta=texto_add_br(alerta)
						s = ""
						if vConsiste(i).pedido_digitado <> "" then s = s & "Nº pedido=" & vConsiste(i).pedido_digitado
						if vConsiste(i).obs2_digitado <> "" then s = s & "Obs II=" & vConsiste(i).obs2_digitado
						s = s & " e "
						if vConsiste(j).pedido_digitado <> "" then s = s & "Nº pedido=" & vConsiste(j).pedido_digitado
						if vConsiste(j).obs2_digitado <> "" then s = s & "Obs II=" & vConsiste(j).obs2_digitado
						alerta=alerta & "Pedido " & vConsiste(i).pedido & ": as seguintes informações fornecidas especificam o mesmo pedido (" & s & ")"
						exit for
						end if
					next
				end if
			next
		end if
	
	if alerta = "" then
		for i = Lbound(vConsiste) to Ubound(vConsiste)
			if vConsiste(i).pedido <> "" then
				s = "SELECT" & _
						" pedido," & _
						" numero_loja," & _
						" id_nfe_emitente," & _
						" obs_2," & _
						" st_entrega," & _
						" entregue_data," & _
						" transportadora_id," & _
						" PedidoRecebidoStatus," & _
						" PedidoRecebidoData," & _
						" data,"

				if blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos then
					s = s & _
						" t_PEDIDO.endereco_nome AS nome"
				else
					s = s & _
						" nome"
					end if

				s = s & _
					" FROM t_PEDIDO INNER JOIN t_CLIENTE" & _
						" ON (t_PEDIDO.id_cliente=t_CLIENTE.id)" & _
					" WHERE" & _
						" (pedido='" & vConsiste(i).pedido & "')"
				if rs.State <> 0 then rs.Close
				rs.open s, cn
				if rs.Eof then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Pedido " & vConsiste(i).pedido & " não está cadastrado."
				else
					blnFlagOk = True
					
				'	SERÁ ENTREGUE, PORTANTO, VERIFICA SE O PEDIDO ESTÁ NO STATUS 'A ENTREGAR'
					if ckb_entrega <> "" then
						if Trim("" & rs("st_entrega")) <> ST_ENTREGA_A_ENTREGAR then
							blnFlagOk = False
							alerta=texto_add_br(alerta)
							alerta=alerta & "Pedido " & vConsiste(i).pedido & " possui status inválido para a operação 'Entrega': " & Ucase(x_status_entrega(Trim("" & rs("st_entrega"))))
							end if
						end if
					
					if ckb_recebido <> "" then
						if CLng(rs("PedidoRecebidoStatus")) = CLng(COD_ST_PEDIDO_RECEBIDO_SIM) then
							blnFlagOk = False
							alerta=texto_add_br(alerta)
							alerta=alerta & "Pedido " & vConsiste(i).pedido & " já consta como Recebido em " & formata_data(rs("PedidoRecebidoData"))
							end if
						end if
						
				'	SE O PEDIDO SERÁ MARCADO COMO RECEBIDO, ENTÃO DEVE ESTAR NO STATUS 'ENTREGUE'
				'	(A NÃO SER QUE ESTEJA SENDO MARCADO COMO 'ENTREGUE' AGORA, NESTA MESMA OPERAÇÃO)
					if (ckb_recebido <> "") And (ckb_entrega = "") then
						if Trim("" & rs("st_entrega")) <> ST_ENTREGA_ENTREGUE then
							blnFlagOk = False
							alerta=texto_add_br(alerta)
							alerta=alerta & "Pedido " & vConsiste(i).pedido & " possui status inválido para a operação 'Recebido': " & Ucase(x_status_entrega(Trim("" & rs("st_entrega"))))
							end if
						end if
					
				'	TRANSPORTADORA CONFERE?
					if UCase(Trim("" & rs("transportadora_id"))) <> UCase(c_transportadora) then
						strTransportadoraPedido = Trim("" & rs("transportadora_id"))
						if strTransportadoraPedido = "" then strTransportadoraPedido = "nenhuma"
						strTransportadoraFiltro = c_transportadora
						if strTransportadoraFiltro = "" then strTransportadoraFiltro = "nenhuma"
						blnFlagOk = False
						alerta=texto_add_br(alerta)
						alerta=alerta & "Pedido " & vConsiste(i).pedido & " indica a transportadora (" & strTransportadoraPedido & ") e não coincide com a transportadora informada na checagem (" & strTransportadoraFiltro & ")"
						end if
					
				'	CONSISTE A DATA DE RECEBIMENTO PELO CLIENTE
					if blnFlagOk then
						if ckb_recebido <> "" then
							if ckb_entrega <> "" then
							'	O PEDIDO ESTÁ SENDO MARCADO COMO 'ENTREGUE' AGORA
								if StrToDate(c_dt_recebido) < Date then
									blnFlagOk = False
									alerta=texto_add_br(alerta)
									alerta=alerta & "Pedido " & vConsiste(i).pedido & ": a data de recebimento informada (" & c_dt_recebido & ") é anterior à data de hoje!"
									end if
							else
							'	O PEDIDO JÁ DEVE ESTAR NO STATUS 'ENTREGUE'
								if StrToDate(c_dt_recebido) < StrToDate(formata_data(rs("entregue_data"))) then
									blnFlagOk = False
									alerta=texto_add_br(alerta)
									alerta=alerta & "Pedido " & vConsiste(i).pedido & ": a data de recebimento informada (" & c_dt_recebido & ") é anterior à data do pedido entregue (" & formata_data(rs("entregue_data")) & ")"
									end if
								end if
							end if
						end if
					
				'	VERIFICA SE O CD DO USUÁRIO ESTÁ COERENTE COM O PEDIDO
					if CLng(rNfeEmitente.id) <> CLng(rs("id_nfe_emitente")) then
					'	ERRO: PEDIDO PERTENCE A OUTRO CD
						blnFlagOk = False
						alerta=texto_add_br(alerta)
						alerta=alerta & "Pedido " & vConsiste(i).pedido & " pertence a outro CD"
						end if

				'	CONSISTÊNCIA OK
					if blnFlagOk then
						with vConsiste(i)
							.data = rs("data")
							.cliente = iniciais_em_maiusculas(Trim("" & rs("nome")))
							.obs2 = Trim("" & rs("obs_2"))
							end with
						end if
					end if
				end if
			next
		end if
		
	lista_pedidos_selecionados = lista_pedidos
	
	if alerta = "" then
		redim vPedidosSelecionados(0)
		vPedidosSelecionados(Ubound(vPedidosSelecionados)) = ""
		for i = Lbound(vConsiste) to Ubound(vConsiste)
			if vConsiste(i).pedido <> "" then
				if vPedidosSelecionados(Ubound(vPedidosSelecionados)) <> "" then
					redim preserve vPedidosSelecionados(Ubound(vPedidosSelecionados)+1)
					vPedidosSelecionados(Ubound(vPedidosSelecionados)) = ""
					end if
				vPedidosSelecionados(Ubound(vPedidosSelecionados)) = vConsiste(i).pedido
				end if
			next
			
		lista_pedidos_selecionados = join(vPedidosSelecionados,chr(13))
		end if
	
	dim dt_limite_inferior
	dt_limite_inferior = DateAdd("m", -3, Date)
	if alerta = "" then
		 if (ckb_recebido <> "") And (c_dt_recebido <> "") then
			if StrToDate(c_dt_recebido) > Date then
				alerta=texto_add_br(alerta)
				alerta=alerta & "Data de recebimento informada é inválida por ser uma data futura (" & c_dt_recebido & ")"
			elseif StrToDate(c_dt_recebido) < dt_limite_inferior then
				alerta=texto_add_br(alerta)
				alerta=alerta & "Data de recebimento informada é inválida por ser muito antiga!<br>Data informada: " & c_dt_recebido & "<br>Data limite: " & formata_data(dt_limite_inferior)
				end if
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



<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
function fPedRecConfirma( f ) {
	dCONFIRMA.style.visibility="hidden";
	window.status = "Aguarde ...";
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
<div class="MtAlerta" style="width:600px;font-weight:bold;" align="left"><p style='margin:5px 20px 5px 20px;'><%=alerta%></p></div>
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
<!-- ***************************************************************** -->
<!-- **********  PÁGINA PARA EXIBIR DADOS PARA CONFIRMAÇÃO  ********** -->
<!-- ***************************************************************** -->
<body onload="focus();">
<center>

<form id="fPedRec" name="fPedRec" method="post" action="PedidoRecebidoConfirma.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="pedidos_selecionados" id="pedidos_selecionados" value="<%=lista_pedidos_selecionados%>">
<input type="hidden" name="ckb_entrega" id="ckb_entrega" value="<%=ckb_entrega%>">
<input type="hidden" name="ckb_recebido" id="ckb_recebido" value="<%=ckb_recebido%>">
<input type="hidden" name="c_dt_recebido" id="c_dt_recebido" value="<%=c_dt_recebido%>">
<input type="hidden" name="c_transportadora" id="c_transportadora" value="<%=c_transportadora%>">

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">

<tr>
	<td align="right" valign="bottom"><p class="PEDIDO">Pedido Recebido Pelo Cliente<span class="C">&nbsp;</span></p></td>
</tr>
</table>
<br>


<!-- ************   HÁ OBSERVAÇÕES?  ************ -->
<% if observacoes <> "" then %>
		<span class="Lbl">OBSERVAÇÕES</span>
		<div class='MtAviso' style="width:649px;font-weight:bold;border:1pt solid black;" align="center"><p style='margin:5px 2px 5px 2px;'><%=observacoes%></p></div>
		<br><br>
<% end if %>

<!-- OPERAÇÃO -->
<table class="Qx" cellSpacing="0" cellPadding="4">
	<tr bgColor="#FFFFFF">
		<td class="MT" NOWRAP colspan="2" style='background:azure;'><span class="PLTe">Operação&nbsp;</span></td>
	</tr>
	<tr>
	<tr bgColor="#FFFFFF">
		<td class="MDBE" NOWRAP>
			<input type="checkbox" tabindex="-1" id="ckb_entrega_aux" name="ckb_entrega_aux" disabled tabindex="-1" value="ON" <%if ckb_entrega <> "" then Response.Write " checked"%>><span class="C" style="cursor:default">Entrega</span>
		</td>
	</tr>
	<tr bgColor="#FFFFFF">
		<td class="MDBE" NOWRAP>
			<input type="checkbox" tabindex="-1" id="ckb_recebido_aux" name="ckb_recebido_aux" disabled tabindex="-1" value="ON" <%if ckb_recebido <> "" then Response.Write " checked"%>><span class="C" style="cursor:default">Recebido em</span>
			&nbsp;
			<input class="Cc" maxlength="10" style="width:90px;" name="c_dt_recebido_aux" id="c_dt_recebido_aux" readonly tabindex=-1 value="<%=c_dt_recebido%>">
		</td>
	</tr>
	<tr bgColor="#FFFFFF">
		<td class="MDBE" NOWRAP valign="bottom">
			<%	strAux = c_transportadora 
				if strAux = "" then 
					strAux = "N.I."
				else
					strAux = strAux & " - " & x_transportadora(strAux)
					end if
			%>
			<span class="PLTe" style="vertical-align:bottom;">Transportadora:&nbsp;</span><span class="C" style="vertical-align:bottom;"><%=strAux%></span>
		</td>
	</tr>
</table>
<br>
<br>

<!--  PEDIDOS  -->
<table class="Qx" cellSpacing="0">
	<tr bgColor="#FFFFFF">
		<td class="MT" NOWRAP style='background:azure;'><span class="PLTe">Pedido&nbsp;</span></td>
		<td class="MC MB MD" NOWRAP style='background:azure;'><span class="PLTe">Obs II&nbsp;</span></td>
		<td class="MC MB MD" NOWRAP style='background:azure;'><span class="PLTe">Cliente&nbsp;</span></td>
	</tr>
<% for i=Lbound(vConsiste) to Ubound(vConsiste) 
		if vConsiste(i).pedido <> "" then %>
			<tr bgColor="#FFFFFF">
				<td class="MDBE" NOWRAP><input name="c_pedido" id="c_pedido" readonly tabindex=-1 class="PLLe" style="width:60px;margin-left:2pt;" 
					value="<%=vConsiste(i).pedido%>"></td>
				<td class="MDB" NOWRAP><input name="c_obs2" id="c_obs2" readonly tabindex=-1 class="PLLe" style="width:80px;margin-left:2pt;" 
					value="<%=vConsiste(i).obs2%>"></td>
				<td class="MDB" NOWRAP><input name="c_cliente" id="c_cliente" readonly tabindex=-1 class="PLLe" style="width:280px;margin-left:2pt;" 
					value="<%=vConsiste(i).cliente%>"></td>
			</tr>
<%			end if
		next		%>
</table>

<!-- ************   SEPARADOR   ************ -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>
<br>


<table class="notPrint" width="649" cellSpacing="0">
<tr>
	<td><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()" title="volta para página anterior">
		<img src="../botao/anterior.gif" width="176" height="55" border="0"></a></td>
	<td align="right"><div name="dCONFIRMA" id="dCONFIRMA">
	<a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fPedRecConfirma(fPedRec)" title="confirma o recebimento do pedido">
		<img src="../botao/confirmar.gif" width="176" height="55" border="0"></a></div>
	</td>
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