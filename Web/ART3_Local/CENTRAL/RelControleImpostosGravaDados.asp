<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =================================================================
'	  RelControleImpostosGravaDados.asp
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

	dim blnExecutaUpdate
	dim s, usuario, msg_erro, s_log, s_log_aux

	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

	dim alerta
	alerta=""

'	OBTÉM FILTROS
	dim c_transportadora, c_dt_coleta, c_dt_coleta_inicio, c_dt_coleta_termino, rb_tipo_consulta, c_nfe_emitente, c_uf

	c_transportadora = Trim(Request.Form("c_transportadora"))
	c_dt_coleta = Trim(Request.Form("c_dt_coleta"))
	c_dt_coleta_inicio = Trim(Request.Form("c_dt_coleta_inicio"))
	c_dt_coleta_termino = Trim(Request.Form("c_dt_coleta_termino"))
	rb_tipo_consulta = Trim(Request.Form("rb_tipo_consulta"))
	c_nfe_emitente = Trim(Request.Form("c_nfe_emitente"))
    c_uf = Trim(Request.Form("c_uf"))

'	OBTÉM DADOS DO FORMULÁRIO
	dim i, n, s_nfe, s_id_nfe, s_num_nfe, s_pedido, vAux, strFlagNFeImpostosOK
	dim intNsu

'	TIPO DE RELATÓRIO: SOLICITAÇÃO DE COLETA
'	========================================
'	CHECK BOX P/ INDICAR A GRAVAÇÃO DA DATA DE COLETA + TRANSPORTADORA
'	CHECK BOX P/ SOLICITAR A EMISSÃO DA NFe
	dim v_nota_verificada_id, v_nota_verificada_num, v_nota_verificada_pedido, qtde_nota_verificada
	redim v_nota_verificada_id(0)
	v_nota_verificada_id(Ubound(v_nota_verificada_id))=""
	redim v_nota_verificada_num(0)
	v_nota_verificada_num(Ubound(v_nota_verificada_num))=""
	redim v_nota_verificada_pedido(0)
	v_nota_verificada_pedido(Ubound(v_nota_verificada_pedido))=""
	qtde_nota_verificada=0
	
	n = Request.Form("ckb_controle_impostos").Count
	for i = 1 to n
		s_nfe = Trim(Request.Form("ckb_controle_impostos")(i))
		if s_nfe <> "" then
			vAux=Split(s_nfe,"|")
			s_id_nfe = Trim(vAux(LBound(vAux)))
			s_num_nfe = Trim(vAux(LBound(vAux)+1))
			s_pedido = Trim(vAux(LBound(vAux)+2))
			strFlagNFeImpostosOK = UCase(Trim(vAux(UBound(vAux))))
			if strFlagNFeImpostosOK = "N" then
				if Trim(v_nota_verificada_id(Ubound(v_nota_verificada_id)))<>"" then
					redim preserve v_nota_verificada_id(Ubound(v_nota_verificada_id)+1)
					end if
				v_nota_verificada_id(Ubound(v_nota_verificada_id)) = s_id_nfe
				if Trim(v_nota_verificada_num(Ubound(v_nota_verificada_num)))<>"" then
					redim preserve v_nota_verificada_num(Ubound(v_nota_verificada_num)+1)
					end if
				v_nota_verificada_num(Ubound(v_nota_verificada_num)) = s_num_nfe
				if Trim(v_nota_verificada_pedido(Ubound(v_nota_verificada_pedido)))<>"" then
					redim preserve v_nota_verificada_pedido(Ubound(v_nota_verificada_pedido)+1)
					end if
				v_nota_verificada_pedido(Ubound(v_nota_verificada_pedido)) = s_pedido
				qtde_nota_verificada = qtde_nota_verificada + 1
				end if
			end if
		next

	if alerta = "" then
		if (qtde_nota_verificada = 0) then
			alerta = "Não foi especificada nenhuma NFe para indicar a verificação dos impostos."
			end if
		end if

	if alerta = "" then
		if (c_dt_coleta = "") and (c_dt_coleta_inicio = "") and (c_dt_coleta_termino = "") then
			alerta = "Data de coleta não foi informada."
		elseif (c_dt_coleta <> "") and (Not isDate(c_dt_coleta)) then
			alerta = "Data de coleta é inválida."
		elseif (c_dt_coleta <> "") and (StrToDate(c_dt_coleta) > Date + 5) then
			alerta = "Data de coleta não pode ser superior a cinco dias."
		elseif (c_dt_coleta_inicio <> "") and (Not isDate(c_dt_coleta_inicio)) then
			alerta = "Data de início é inválida."
		elseif (c_dt_coleta_termino <> "") and (Not isDate(c_dt_coleta_termino)) then
			alerta = "Data de início é inválida."
		elseif (StrToDate(c_dt_coleta_inicio) > Date + 5) or (StrToDate(c_dt_coleta_termino) > Date + 5) then
			alerta = "Período de coleta informado não pode conter uma data futura superior a 05 dias."
			end if
		end if

'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	
	s_log = ""
	
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

		for i=Lbound(v_nota_verificada_id) to Ubound(v_nota_verificada_id)
			if v_nota_verificada_id(i) <> "" then
				s = "SELECT * FROM t_NFE_EMISSAO WHERE (id = '" & v_nota_verificada_id(i) & "')"
				if rs.State <> 0 then rs.Close
				rs.Open s, cn
				if rs.Eof then
					alerta=texto_add_br(alerta)
					alerta=alerta & "NFe n° " & v_nota_verificada_num(i) & " não foi encontrada."
				else
                '   VERIFICA SE OUTRO USUÁRIO ALTEROU O STATUS
                    if rs("controle_impostos_status")=CInt(COD_CONTROLE_IMPOSTOS_STATUS__OK) then
                        alerta=texto_add_br(alerta)
                        alerta=alerta & "Status do controle de impostos do pedido " & v_nota_verificada_pedido(i) & " foi alterado por outro usuário (" & Trim("" & rs("controle_impostos_usuario")) & " às " & formata_data_hora_sem_seg(rs("controle_impostos_data_hora")) & ")"
                    else
					    rs("controle_impostos_status")=CInt(COD_CONTROLE_IMPOSTOS_STATUS__OK)
					    rs("controle_impostos_data")=Date
					    rs("controle_impostos_data_hora")=Now
					    rs("controle_impostos_usuario")=usuario
					    rs.Update
					    if Err <> 0 then 
						    alerta=texto_add_br(alerta)
						    alerta=alerta & Cstr(Err) & ": " & Err.Description
						    end if
                        end if
						
					s_log = ""
					if alerta = "" then
					'	INFORMAÇÕES PARA O LOG
						s_log = s_log & v_nota_verificada_num(i)
						end if

					if alerta = "" then
						if s_log <> "" then
							s_log = "Verificação de impostos realizada para a NFe " & s_log
							grava_log usuario, "", v_nota_verificada_pedido(i), "", OP_LOG_NFE_CTRL_IMPOSTOS, s_log
							end if
						end if

					end if
				end if
			next

'		if alerta = "" then
'			if s_log <> "" then
'				s_log = "Verificação de impostos realizada para a(s) NFe(s): " & s_log
'				grava_log usuario, "", "", "", OP_LOG_NFE_CTRL_IMPOSTOS, s_log
'				end if
'			end if
		
    '   LIMPA LOCKS
        s = "UPDATE t_CTRL_RELATORIO_USUARIO_X_PEDIDO SET" & _
                " locked = 0," & _
                " cod_motivo_lock_released = " & CTRL_RELATORIO_CodMotivoLockReleased_OperacaoFinalizada & "," & _
                " dt_hr_lock_released = getdate()" & _
            " WHERE" & _
                " (usuario = '" & QuotedStr(usuario) & "')" & _
                " AND (id_relatorio = " & ID_CTRL_RELATORIO_RelControleImpostos & ")" & _
                " AND (locked = 1)"
        cn.Execute(s)

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

        if alerta <> "" then
            grava_log usuario, "", "", "", OP_LOG_NFE_CTRL_IMPOSTOS, alerta
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
function fRetornar(f) {
	f.action = "RelControleImpostosExec.asp?url_back=X";
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

<form id="f" name="f" method="post">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<!-- FILTROS -->
<input type="hidden" name="c_transportadora" id="c_transportadora" value="<%=c_transportadora%>" />
<input type="hidden" name="c_dt_coleta" id="c_dt_coleta" value="<%=c_dt_coleta%>" />
<input type="hidden" name="c_dt_coleta_inicio" id="c_dt_coleta_inicio" value="<%=c_dt_coleta_inicio%>" />
<input type="hidden" name="c_dt_coleta_termino" id="c_dt_coleta_termino" value="<%=c_dt_coleta_termino%>" />
<input type="hidden" name="rb_tipo_consulta" id="rb_tipo_consulta" value="<%=rb_tipo_consulta%>" />
<input type="hidden" name="c_nfe_emitente" id="c_nfe_emitente" value="<%=c_nfe_emitente%>" />
<input type="hidden" name="c_uf" id="c_uf" value="<%=c_uf%>" />

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">

<tr>
	<td align="right" valign="bottom"><p class="PEDIDO">Controle de Impostos<span class="C">&nbsp;</span></p></td>
</tr>
</table>
<br>
<br>

<% if qtde_nota_verificada > 0 then %>
<!-- ************   MENSAGEM  ************ -->
<% 
	s = ""
	for i=Lbound(v_nota_verificada_num) to Ubound(v_nota_verificada_num)
		if v_nota_verificada_num(i) <> "" then
			if s <> "" then s = s & ", "
			s = s & v_nota_verificada_num(i)
			end if
		next
	
	if s = "" then s = "nenhuma NFe"
%>
<div class='MtAviso' style="width:649px;font-weight:bold;border:1pt solid black;" align="center"><p style='margin:5px 2px 5px 2px;'>Anotação de verificação de impostos<br />NFe(s): <%=s%></p></div>
<br>
<% end if %>

<!-- ************   SEPARADOR   ************ -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
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