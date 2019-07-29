<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =================================================================
'	  RelClientesNegativadosGravaDados.asp
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

	class cl_TIPO_GRAVA_CLIENTE_NEGATIVADO
		dim id_cliente
		dim nome_cliente
		dim cnpj_cliente
		dim status_negativado
		dim descricao_status
		end class

	Const COD_MANUAL_NAO_TRATADO = 0
	Const COD_MANUAL_TRATADO = 1

	dim s, usuario, msg_erro, s_log

	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO)

	dim alerta
	alerta=""

'	OBTÉM FILTROS
	'dim c_dt_inicio, c_dt_termino, c_resultado_transacao, c_bandeira, c_pedido, c_cliente_cnpj_cpf, c_loja, rb_ordenacao_saida
	dim c_dt_inicio, c_dt_termino
	dim c_cliente_cnpj_cpf, rb_negativado, rb_ordenacao_saida
	dim s_nome_cliente
	dim c_uf

	c_uf = Trim(Request("c_uf"))
	c_cliente_cnpj_cpf = Trim(Request("c_cliente_cnpj_cpf"))
	rb_negativado = Trim(Request("rb_negativado"))
	rb_ordenacao_saida = Trim(Request("rb_ordenacao_saida"))

'	OBTÉM DADOS DO FORMULÁRIO
	dim i, j, n, c, vAux, s_dados, s_dados_anteriores, s_dados_aux, s_id_cliente, s_status_anterior, s_cnpj, s_nome
	dim intNsuHistorico, s_id_historico

'	CHECK BOX P/ INDICAR A GRAVAÇÃO DA INFORMAÇÃO DE QUE O CLIENTE FOI TRATADO
'	CAMPO TEXTO P/ INCLUIR OBSERVAÇÕES DA TRANSAÇÃO TRATADA
	dim v_clientes, qtde_clientes
	redim v_clientes(0)
	set v_clientes(Ubound(v_clientes)) = new cl_TIPO_GRAVA_CLIENTE_NEGATIVADO
	v_clientes(Ubound(v_clientes)).id_cliente=""
	qtde_clientes=0
	
	n = Request.Form("ckb_inicial").Count
	c = Request.Form("ckb_negativado").Count
	for i = 1 to n
	
		'obtemos cada linha da coluna oculta da página anterior
		s_dados_anteriores = Trim(Request.Form("ckb_inicial")(i))

		'nem sempre teremos o retorno da coluna do checkbox (quando o mesmo está desmarcado, retorna valor vazio);
		'portanto, varreremos as linhas desta coluna, comparando com a linha da coluna oculta;
		'se encontrar, significa que o checkbox foi marcado para negativação (ou já estava marcado, portanto, não será alterado);
		'se não encontrar, significa que o checkbox foi desmarcado (ou não estava desde o início, portanto, não será alterado)
		j = 1
		s_dados = ""
		do while (j <= c) and (s_dados = "")
			s_dados_aux = Trim(Request.Form("ckb_negativado")(j))
			if (s_dados_aux <> "") and (InStr(s_dados_anteriores, s_dados_aux) > 0) then
				s_dados = s_dados_aux
				end if
			j = j + 1
			loop
		
		if s_dados_anteriores <> "" then
			vAux = Split(s_dados_anteriores, "|")
			s_status_anterior = Trim(vAux(LBound(vAux)))
			s_id_cliente = Trim(vAux(LBound(vAux)+1))
			s_nome = Trim(vAux(LBound(vAux)+2))
			s_cnpj = Trim(vAux(LBound(vAux)+3))
			
			'primeiro caso: clientes não negativados que foram marcados
			if ((s_dados <> "") and (s_status_anterior = "false")) then
				if v_clientes(Ubound(v_clientes)).id_cliente <> "" then
					redim preserve v_clientes(Ubound(v_clientes)+1)
					set v_clientes(Ubound(v_clientes)) = new cl_TIPO_GRAVA_CLIENTE_NEGATIVADO
					end if
				v_clientes(Ubound(v_clientes)).id_cliente = s_id_cliente
				v_clientes(Ubound(v_clientes)).nome_cliente = s_nome
				v_clientes(Ubound(v_clientes)).cnpj_cliente = s_cnpj
				v_clientes(Ubound(v_clientes)).status_negativado = 1
				v_clientes(Ubound(v_clientes)).descricao_status = "negativação efetuada"
				qtde_clientes = qtde_clientes + 1
			'segundo caso: clientes negativados que foram desmarcados
			elseif ((s_dados = "") and (s_status_anterior = "true")) then
				if v_clientes(Ubound(v_clientes)).id_cliente <> "" then
					redim preserve v_clientes(Ubound(v_clientes)+1)
					set v_clientes(Ubound(v_clientes)) = new cl_TIPO_GRAVA_CLIENTE_NEGATIVADO
					end if
				v_clientes(Ubound(v_clientes)).id_cliente = s_id_cliente
				v_clientes(Ubound(v_clientes)).nome_cliente = s_nome
				v_clientes(Ubound(v_clientes)).cnpj_cliente = s_cnpj
				v_clientes(Ubound(v_clientes)).status_negativado = 0
				v_clientes(Ubound(v_clientes)).descricao_status = "negativação removida"
				qtde_clientes = qtde_clientes + 1
				end if
				
		end if
		
		next

	if alerta = "" then
		if (qtde_clientes = 0) then
			alerta = "Nenhum cliente foi alterado."
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
		for i=Lbound(v_clientes) to Ubound(v_clientes)
			if (v_clientes(i).id_cliente <> "") then
				s = "SELECT * FROM t_CLIENTE WHERE (id = '" & v_clientes(i).id_cliente & "')"
				if rs.State <> 0 then rs.Close
				rs.Open s, cn
				if rs.Eof then
					alerta=texto_add_br(alerta)
					alerta=alerta & "Registro de cliente não foi encontrado (ID=" & v_clientes(i).id_cliente & ")."
				else
					rs("spc_negativado_status") = v_clientes(i).status_negativado
					rs("spc_negativado_data") = Date
					rs("spc_negativado_data_hora") = Now
					rs("spc_negativado_usuario") = usuario
					
					rs.Update
					
					if alerta = "" then
					'	INFORMAÇÕES PARA O LOG
						if s_log <> "" then s_log = s_log & ", "
						s_log = s_log & v_clientes(i).id_cliente & " (CNPJ/CPF=" & v_clientes(i).cnpj_cliente & "; status=" & v_clientes(i).status_negativado & ")"
						end if
					end if
				end if

		'	GERA O NSU PARA O NOVO REGISTRO DO HISTÓRICO
			if Not fin_gera_nsu(T_CLIENTE_SPC_HISTORICO, intNsuHistorico, msg_erro) then
				alerta = "FALHA AO GERAR NSU PARA O NOVO REGISTRO DO HISTÓRICO (" & msg_erro & ")"
			else
				if intNsuHistorico <= 0 then
					alerta = "NSU GERADO É INVÁLIDO (" & intNsuHistorico & ")"
				else
					s_id_historico = Cstr(intNsuHistorico)
					end if
				end if

		'	GRAVA NO HISTÓRICO DE NEGATIVAÇÃO
			if alerta = "" then
				s = "INSERT INTO t_CLIENTE_SPC_HISTORICO " & _
						"(" & _
						"id, " & _
						"id_cliente, " & _
						"data, " & _
						"data_hora, " & _
						"usuario, " & _
						"spc_negativado_status," & _
						"cnpj_cpf" & _
						")" & _
					" VALUES (" & _
						s_id_historico & ", " & _
						"'" & v_clientes(i).id_cliente & "', " & _
						bd_formata_data(Now) & ", " & _
						bd_formata_data_hora(Now) & ", " & _
						"'" & usuario & "', " & _
						CStr(v_clientes(i).status_negativado) & ", "  & _
						"'" & retorna_so_digitos(v_clientes(i).cnpj_cliente) & "'" & _
						")"
				cn.Execute(s)
				If Err <> 0 then
					alerta = "FALHA AO TENTAR GRAVAR O HISTÓRICO (" & Cstr(Err) & ": " & Err.Description & ")."
					end if
				end if

			
		'	SE HOUVE ERRO, CANCELA O LAÇO
			if alerta <> "" then exit for
			next

		if alerta = "" then
			if s_log <> "" then
				s_log = "Clientes alterados: " & s_log
				grava_log usuario, "", "", "", OP_LOG_SPC_CLIENTE_NEGATIVADO_ALTERACAO, s_log
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
	f.action = "RelClientesNegativadosExec.asp?url_back=X";
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
<input type="hidden" name="rb_negativado" id="rb_negativado" value="<%=rb_negativado%>" />
<input type="hidden" name="c_cliente_cnpj_cpf" id="c_cliente_cnpj_cpf" value="<%=c_cliente_cnpj_cpf%>" />
<input type="hidden" name="rb_ordenacao_saida" id="rb_ordenacao_saida" value="<%=rb_ordenacao_saida%>" />
<input type="hidden" name="c_uf" id="c_uf" value="<%=c_uf%>" />

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black;">

<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Clientes Negativados (SPC)<span class="C">&nbsp;</span></span></td>
</tr>
</table>
<br>
<br>

<%if qtde_clientes > 0 then %>
<!-- ************   MENSAGEM  ************ -->
<% 
	s = ""
	for i=Lbound(v_clientes) to Ubound(v_clientes)
		if v_clientes(i).id_cliente <> "" then
			if s <> "" then s = s & "<br />"
			s = s & v_clientes(i).cnpj_cliente & " - " & v_clientes(i).nome_cliente & " - (" & v_clientes(i).descricao_status & ")"
			end if
		next
	
	if s = "" then s = "nenhum cliente selecionado"
%>
<div class='MtAviso' style="width:649px;font-weight:bold;border:1pt solid black;" align="center"><p style='margin:5px 2px 5px 2px;'>Clientes alterados:<br /><br /> <%=s%></p></div>
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