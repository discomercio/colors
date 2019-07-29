<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =================================================================
'	  R O M A N E I O C O N S I S T E . A S P
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

	dim s, usuario, s_transportadora, s_nome_transportadora, lista_pedidos, v_pedido, v_aux, i, j, achou, msg_erro
	dim c_dt_entrega, c_num_coleta, c_transportadora_contato, c_conferente, c_motorista, c_placa_veiculo, c_nfe_emitente
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 
	
'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	
	s_transportadora = Ucase(Trim(request("c_transportadora")))
	c_num_coleta = Trim(Request.Form("c_num_coleta"))
	c_dt_entrega = Trim(Request.Form("c_dt_entrega"))
	c_transportadora_contato = Trim(Request.Form("c_transportadora_contato"))
	c_conferente = Trim(Request.Form("c_conferente"))
	c_motorista = Trim(Request.Form("c_motorista"))
	c_placa_veiculo = Trim(Request.Form("c_placa_veiculo"))
	c_nfe_emitente = Trim(Request.Form("c_nfe_emitente"))
	
	lista_pedidos = ucase(Trim(request("c_pedidos")))
	if (lista_pedidos = "") then Response.Redirect("aviso.asp?id=" & ERR_PEDIDO_NAO_ESPECIFICADO)
	
	lista_pedidos=substitui_caracteres(lista_pedidos,chr(10),"")
	v_aux = split(lista_pedidos,chr(13),-1)
	achou=False
	for i=Lbound(v_aux) to Ubound(v_aux)
		if Trim(v_aux(i))<>"" then
			achou = True
			s = normaliza_num_pedido(v_aux(i))
			if s <> "" then v_aux(i) = s
			end if
		next

	if Not achou then Response.Redirect("aviso.asp?id=" & ERR_PEDIDO_NAO_ESPECIFICADO)

	redim v_pedido(0)
	v_pedido(Ubound(v_pedido))=""
	for i = Lbound(v_aux) to Ubound(v_aux)
		if Trim(v_aux(i))<>"" then
			if Trim(v_pedido(Ubound(v_pedido)))<>"" then
				redim preserve v_pedido(Ubound(v_pedido)+1)
				end if
			v_pedido(Ubound(v_pedido)) = Trim(v_aux(i))
			end if
		next
	
	lista_pedidos = join(v_pedido,chr(13))
	
	dim alerta
	alerta = ""

	dim observacoes
	observacoes = ""
	
	for i=Lbound(v_pedido) to Ubound(v_pedido)
		if v_pedido(i)<>"" then
			for j=Lbound(v_pedido) to (i-1)
				if v_pedido(i) = v_pedido(j) then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Pedido " & v_pedido(i) & ": linha " & renumera_com_base1(Lbound(v_pedido),i) & " repete o mesmo pedido da linha " & renumera_com_base1(Lbound(v_pedido),j) & "."
					exit for
					end if
				next
			end if
		next

	if alerta = "" then
		if s_transportadora = "" then
			alerta = "Informe a transportadora"
		else
			s = "SELECT id, nome FROM t_TRANSPORTADORA WHERE (id='" & s_transportadora & "')"
			if rs.State <> 0 then rs.Close
			rs.open s, cn
			if rs.Eof then
				alerta=texto_add_br(alerta)
				alerta=alerta & "Transportadora '" & s_transportadora & "' não está cadastrada."
			else
				s_nome_transportadora = Trim("" & rs("nome"))
				end if
			end if
		end if
	
	if alerta = "" then
		if c_dt_entrega = "" then
			alerta = "Informe a data para entrega"
		elseif Not IsDate(c_dt_entrega) then
			alerta = "Data de coleta é inválida"
		else
			if StrToDate(c_dt_entrega) < Date then alerta = "Data de coleta não pode ser uma data passada"
			end if
		end if
	
	if alerta = "" then
		if c_conferente = "" then
			alerta = "Informe o nome do conferente."
		elseif c_motorista = "" then
			alerta = "Informe o nome do motorista."
		elseif c_placa_veiculo = "" then
			alerta = "Informe a placa do veículo."
		elseif Not isPlacaVeiculoOk(c_placa_veiculo) then
			alerta = "Placa de veículo inválida."
			end if
		end if
		
	dim vDadosPedido()
	redim vDadosPedido(Ubound(v_pedido))
	
	if alerta = "" then
		for i = Lbound(v_pedido) to Ubound(v_pedido)
		'	COPIA O Nº DO PEDIDO P/ O VETOR QUE ARMAZENA OS DADOS DO PEDIDO P/ EXIBIÇÃO NA TELA
			set vDadosPedido(i) = new cl_QUATRO_COLUNAS
			vDadosPedido(i).c1 = v_pedido(i)
			
			if v_pedido(i) <> "" then
				s = "SELECT" & _
						" pedido, st_entrega, transportadora_id, a_entregar_data_marcada," & _
						" danfe_impressa_status, romaneio_status, romaneio_data_hora, romaneio_usuario, " & _
						" obs_1, forma_pagto, tipo_parcelamento, av_forma_pagto," & _
						" pu_forma_pagto, pu_valor, pu_vencto_apos," & _
						" pc_qtde_parcelas, pc_valor_parcela, pc_maquineta_qtde_parcelas, pc_maquineta_valor_parcela," & _
						" pce_forma_pagto_entrada, pce_forma_pagto_prestacao, pce_entrada_valor, pce_prestacao_qtde," & _
						" pce_prestacao_valor, pce_prestacao_periodo, pse_forma_pagto_prim_prest," & _
						" pse_forma_pagto_demais_prest, pse_prim_prest_valor, pse_prim_prest_apos," & _
						" pse_demais_prest_qtde, pse_demais_prest_valor, pse_demais_prest_periodo" & _
					" FROM t_PEDIDO" & _
					" WHERE" & _
						" (pedido='" & v_pedido(i) & "')"
				if rs.State <> 0 then rs.Close
				rs.open s, cn
				if rs.Eof then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Pedido " & v_pedido(i) & " não está cadastrado."
				else
					vDadosPedido(i).c2 = Trim("" & rs("obs_1"))
					vDadosPedido(i).c3 = monta_descricao_forma_pagto(rs)
					vDadosPedido(i).c4 = Trim("" & rs("forma_pagto"))
					
				'	ALERTA DE ERRO
					if Not IsPedidoRomaneioPossivel(Trim("" & rs("st_entrega"))) then
						alerta=texto_add_br(alerta)
						alerta=alerta & "Pedido " & v_pedido(i) & " possui status inválido para esta operação: " & Ucase(x_status_entrega(Trim("" & rs("st_entrega"))))
					elseif (CInt(rs("danfe_impressa_status")) <> CInt(COD_DANFE_IMPRESSA_STATUS__OK)) And _
						   (CInt(rs("danfe_impressa_status")) <> CInt(COD_DANFE_IMPRESSA_STATUS__NAO_DEFINIDO)) then
						alerta=texto_add_br(alerta)
						alerta=alerta & "Pedido " & v_pedido(i) & ": ainda não consta a impressão da DANFE."
						end if
					
				'	OBSERVAÇÕES P/ O USUÁRIO FICAR ATENTO A POSSÍVEIS ENGANOS
					if Trim("" & rs("a_entregar_data_marcada")) <> "" then
						if CDate(rs("a_entregar_data_marcada")) <> StrToDate(c_dt_entrega) then
							observacoes=texto_add_br(observacoes)
							observacoes=observacoes & "Pedido " & v_pedido(i) & " possui data de coleta para " & formata_data(rs("a_entregar_data_marcada"))
							end if
						end if
					
					if Trim("" & rs("transportadora_id")) <> "" then
						if UCase(Trim("" & rs("transportadora_id"))) <> UCase(s_transportadora) then
							observacoes=texto_add_br(observacoes)
							observacoes=observacoes & "Pedido " & v_pedido(i) & " possui a seguinte transportadora anotada: " & Trim("" & rs("transportadora_id"))
							end if
						end if
						
					if CInt(rs("romaneio_status")) = CInt(COD_ROMANEIO_STATUS__OK) then
						observacoes=texto_add_br(observacoes)
						observacoes=observacoes & "Pedido " & v_pedido(i) & ": consta que já foi incluído em um romaneio em " & formata_data_e_talvez_hora_hhmm(rs("romaneio_data_hora")) & " por '" & Trim("" & rs("romaneio_usuario")) & "'"
						end if
					end if
				end if
			next
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

<script language="JavaScript" type="text/javascript">
var MAX_TAM_CAMPO_OBS = 250;
var MAX_TAM_CAMPO_PAGTO = 250;
	
function fConfirma( f ) {
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

<style type="text/css">
table#tabPedido
{
	padding: 1px;
}
.colPedido 
{
	width: 60px;
	vertical-align: top;
	text-align: center;
}
.colObs1
{
	width:230px;
	vertical-align: top;
}
.colFormaPagto
{
	width:230px;
	vertical-align: top;
}
.colObsPlanilha
{
	width:230px;
	vertical-align: top;
}
.colPagtoPlanilha
{
	width:230px;
	vertical-align: top;
}
</style>


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
<!-- ***************************************************************** -->
<!-- **********  PÁGINA PARA EXIBIR DADOS PARA CONFIRMAÇÃO  ********** -->
<!-- ***************************************************************** -->
<body onload="focus();">
<center>

<form id="f" name="f" method="post" action="RomaneioGeraPlanilha.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="pedidos_selecionados" id="pedidos_selecionados" value="<%=lista_pedidos%>">
<input type="hidden" name="c_transportadora" id="c_transportadora" value="<%=s_transportadora%>">
<input type="hidden" name="c_num_coleta" id="c_num_coleta" value="<%=c_num_coleta%>">
<input type="hidden" name="c_dt_entrega" id="c_dt_entrega" value="<%=c_dt_entrega%>">
<input type="hidden" name="c_transportadora_contato" id="c_transportadora_contato" value="<%=c_transportadora_contato%>">
<input type="hidden" name="c_conferente" id="c_conferente" value="<%=c_conferente%>" />
<input type="hidden" name="c_motorista" id="c_motorista" value="<%=c_motorista%>" />
<input type="hidden" name="c_placa_veiculo" id="c_placa_veiculo" value="<%=c_placa_veiculo%>" />
<input type="hidden" name="c_nfe_emitente" id="c_nfe_emitente" value="<%=c_nfe_emitente%>" />

<!-- FORÇA A CRIAÇÃO DE UM ARRAY MESMO QUANDO EXISTE SOMENTE 1 ITEM -->
<input type="hidden" name="c_pedido" id="c_pedido" value="">
<input type="hidden" name="c_obs" id="c_obs" value="">
<input type="hidden" name="c_pagto" id="c_pagto" value="">


<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="1000" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">

<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Romaneio de Entrega<span class="C">&nbsp;</span></span></td>
</tr>
</table>
<br>


<!-- ************   HÁ OBSERVAÇÕES?  ************ -->
<% if observacoes <> "" then %>
		<span class="Lbl">OBSERVAÇÕES</span>
		<div class='MtAviso' style="width:800px;font-weight:bold;border:1pt solid black;" align="center"><span style='margin:5px 2px 5px 2px;'><%=observacoes%></span></div>
		<br><br>
<% end if %>

<!--  PEDIDOS  -->
<table class="Qx" cellSpacing="0">
	<tr bgColor="#FFFFFF">
		<td class="MT" align="right" nowrap style='background:azure;'><span class="PLTe">Transportadora&nbsp;</span></td>
		<td class="MTBD" align="left" nowrap><span class="PLLe" style="margin-left:2pt;margin-right:5pt;"><%=s_transportadora & " - " & s_nome_transportadora%></span></td>
	</tr>
	<tr bgColor="#FFFFFF">
		<td class="MDBE" align="right" nowrap style='background:azure;'><span class="PLTe">Nº Coleta&nbsp;</span></td>
		<td class="MDB" align="left" nowrap><span class="PLLe" style="margin-left:2pt;margin-right:5pt;"><%=c_num_coleta%></span></td>
	</tr>
	<tr bgColor="#FFFFFF">
		<td class="MDBE" align="right" nowrap style='background:azure;'><span class="PLTe">Data de Coleta&nbsp;</span></td>
		<td class="MDB" align="left" nowrap><span class="PLLe" style="margin-left:2pt;margin-right:5pt;"><%=c_dt_entrega%></span></td>
	</tr>
	<tr bgColor="#FFFFFF">
		<td class="MDBE" align="right" nowrap style='background:azure;'><span class="PLTe">Contato na Transportadora&nbsp;</span></td>
		<td class="MDB" align="left" nowrap><span class="PLLe" style="margin-left:2pt;margin-right:5pt;"><%=c_transportadora_contato%></span></td>
	</tr>
	<tr bgColor="#FFFFFF">
		<td class="MDBE" align="right" nowrap style='background:azure;'><span class="PLTe">Conferente&nbsp;</span></td>
		<td class="MDB" align="left" nowrap><span class="PLLe" style="margin-left:2pt;margin-right:5pt;"><%=c_conferente%></span></td>
	</tr>
	<tr bgColor="#FFFFFF">
		<td class="MDBE" align="right" nowrap style='background:azure;'><span class="PLTe">Motorista&nbsp;</span></td>
		<td class="MDB" align="left" nowrap><span class="PLLe" style="margin-left:2pt;margin-right:5pt;"><%=c_motorista%></span></td>
	</tr>
	<tr bgColor="#FFFFFF">
		<td class="MDBE" align="right" nowrap style='background:azure;'><span class="PLTe">Placa do Veículo&nbsp;</span></td>
		<td class="MDB" align="left" nowrap><span class="PLLe" style="margin-left:2pt;margin-right:5pt;"><%=c_placa_veiculo%></span></td>
	</tr>
</table>
<br>
<table id="tabPedido" class="Qx" cellSpacing="0">
	<tr style='background:azure;'>
		<td class="MT colPedido" align="left"><span class="PLTc">Pedido</span></td>
		<td class="MTBD colObs1" align="left"><span class="PLTe">Observações I</span></td>
		<td class="MTBD colObsPlanilha" align="left"><span class="PLTe">OBS (planilha)</span></td>
		<td class="MTBD colFormaPagto" align="left"><span class="PLTe">Forma de Pagamento</span></td>
		<td class="MTBD colPagtoPlanilha" align="left"><span class="PLTe">&nbsp;</span></td>
	</tr>
	
<% for i=Lbound(vDadosPedido) to Ubound(vDadosPedido) 
		if vDadosPedido(i).c1 <> "" then %>
			<tr bgcolor="#FFFFFF">
				<td class="MDBE colPedido" style="height:80px;" align="left" valign="top" nowrap><input name="c_pedido" id="c_pedido" readonly tabindex=-1 class="PLLc" style="width:60px;margin-left:2pt;" 
					value="<%=vDadosPedido(i).c1%>"></td>
				<%	s = vDadosPedido(i).c2
					if s = "" then s = "&nbsp;" %>
				<td class="MDB colObs1" align="left"><span class="Cn"><%=s%></span></td>
				<td class="MDB colObsPlanilha" align="left"><textarea name='c_obs' id="c_obs" class='PLLe' rows='5' style='width:100%;height:100%;margin-left:0pt;' onkeypress='limita_tamanho(this,MAX_TAM_CAMPO_OBS);' onblur='this.value=trim(this.value);'><%=s%></textarea></td>
				<%	s = vDadosPedido(i).c3
					if s = "" then s = "&nbsp;" %>
				<td class="MDB colFormaPagto" align="left"><span class="Cn"><%=s%></span></td>
				<td class="MDB colPagtoPlanilha" align="left"><textarea name='c_pagto' id='c_pagto' class='PLLe' rows='5' style='width:100%;height:100%;margin-left:0pt;' onkeypress='limita_tamanho(this,MAX_TAM_CAMPO_PAGTO);' onblur='this.value=trim(this.value);'></textarea><span class="Cn"></span></td>
			</tr>
<%			end if
		next		%>
</table>

<!-- ************   SEPARADOR   ************ -->
<table width="1000" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc" align="left">&nbsp;</td></tr>
</table>
<br>


<table class="notPrint" width="649" cellSpacing="0">
<tr>
	<td align="left"><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()" title="volta para página anterior">
		<img src="../botao/anterior.gif" width="176" height="55" border="0"></a></td>
	<td align="right"><div name="dCONFIRMA" id="dCONFIRMA">
	<a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fConfirma(f)" title="confirma o romaneio de entrega">
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