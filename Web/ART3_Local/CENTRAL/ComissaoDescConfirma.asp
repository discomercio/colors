<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =================================================================
'	  C O M I S S A O D E S C C O N F I R M A . A S P
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

	dim s, usuario, msg_erro, s_log, rb_comissao_descontada, i
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	if Not operacao_permitida(OP_CEN_FLAG_COMISSAO_PAGA, s_lista_operacoes_permitidas) then 
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	dim alerta
	alerta=""

'	OBTÉM DADOS DO FORMULÁRIO
	rb_comissao_descontada = Trim(Request.Form("rb_comissao_descontada"))

	dim n, qtde_registro
	dim v_registro, v_aux
	redim v_registro(0)
	qtde_registro = 0
	n = Request.Form("ckb_alterar").Count
	for i = 1 to n
		s = Trim(Request.Form("ckb_alterar")(i))
		if s <> "" then
			if Trim(v_registro(ubound(v_registro))) <> "" then
				redim preserve v_registro(ubound(v_registro)+1)
				end if
			qtde_registro = qtde_registro + 1
			v_registro(ubound(v_registro)) = s
			end if
		next
		
	if qtde_registro = 0 then
		alerta = "Nenhum item foi selecionado"
		end if
		
	
'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	const c_OP_DEVOLUCAO = "DEVOLUCAO"
	const c_OP_PERDA = "PERDA"
	Dim s_nome_tabela, s_operacao, s_pedido, s_id_registro, s_data, s_valor
	
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

		for i=Lbound(v_registro) to Ubound(v_registro)
			if Trim(v_registro(i)) <> "" then
			'	PEDIDO / ID REGISTRO / OPERAÇÃO
				v_aux = split(v_registro(i), ";", -1)
				s_pedido = v_aux(LBound(v_aux))
				s_id_registro = v_aux(LBound(v_aux)+1)
				s_operacao = v_aux(LBound(v_aux)+2)

				if s_operacao = c_OP_DEVOLUCAO then
					s_nome_tabela = "t_PEDIDO_ITEM_DEVOLVIDO"
				elseif s_operacao = c_OP_PERDA then
					s_nome_tabela = "t_PEDIDO_PERDA"
				else
					alerta = "OPERAÇÃO DESCONHECIDA (" & s_operacao & ")"
					end if
					
				if alerta = "" then
					s = "SELECT * FROM " & s_nome_tabela & " WHERE (id = '" & s_id_registro & "')"
					if rs.State <> 0 then rs.Close
					rs.Open s, cn
					if rs.Eof then
						alerta=texto_add_br(alerta)
						alerta=alerta & "Registro id=" & s_id_registro & " da tabela " & s_nome_tabela & " referente ao pedido " & s_pedido & " não foi encontrado."
					else
						if rb_comissao_descontada = "S" then
							if rs("comissao_descontada") = CLng(COD_COMISSAO_DESCONTADA) then
								alerta=texto_add_br(alerta)
								alerta=alerta & "Registro id=" & s_id_registro & " da tabela " & s_nome_tabela & " referente ao pedido " & s_pedido & " já está assinalado com comissão descontada."
								end if
						elseif rb_comissao_descontada = "N" then
							if rs("comissao_descontada") = CLng(COD_COMISSAO_NAO_DESCONTADA) then
								alerta=texto_add_br(alerta)
								alerta=alerta & "Registro id=" & s_id_registro & " da tabela " & s_nome_tabela & " referente ao pedido " & s_pedido & " já está assinalado com comissão não-descontada."
								end if
						else
							alerta=texto_add_br(alerta)
							alerta=alerta & "Opção desconhecida (" & rb_comissao_descontada & ")."
							end if

						if rb_comissao_descontada = "S" then
							rs("comissao_descontada")=CLng(COD_COMISSAO_DESCONTADA)
						elseif rb_comissao_descontada = "N" then
							rs("comissao_descontada")=CLng(COD_COMISSAO_NAO_DESCONTADA)
							end if
							
						if s_operacao = c_OP_DEVOLUCAO then
							s_data = formata_data(rs("devolucao_data"))
							s_valor = formata_moeda(rs("qtde")*rs("preco_venda"))
						else
							s_data = formata_data(rs("data"))
							s_valor = formata_moeda(rs("valor"))
							end if
							
						rs("comissao_descontada_ult_op")=rb_comissao_descontada
						rs("comissao_descontada_data")=Date
						rs("comissao_descontada_usuario")=usuario
						rs.Update
						if Err <> 0 then 
							alerta=texto_add_br(alerta)
							alerta=alerta & Cstr(Err) & ": " & Err.Description
							end if
					'	INFORMAÇÕES PARA O LOG
						if s_log <> "" then s_log = s_log & ", "
						s_log = s_log & s_pedido & "=(" & s_valor & " em " & s_data & ", " & s_nome_tabela & "->" & s_id_registro & ")"
						if rs.State <> 0 then rs.Close
						end if
					end if
				end if
				
		'	SE HOUVE ERRO, CANCELA O LAÇO
			if alerta <> "" then exit for
			next

		if alerta = "" then
			if rb_comissao_descontada = "S" then
				s_log = "Assinala comissão descontada; Pedido(s): " & s_log
			elseif rb_comissao_descontada = "N" then
				s_log = "Assinala comissão não-descontada; Pedido(s): " & s_log
				end if
			
			grava_log usuario, "", "", "", OP_LOG_PEDIDO_ALTERACAO, s_log
		'	~~~~~~~~~~~~~~
			cn.CommitTrans
		'	~~~~~~~~~~~~~~
			if Err=0 then 
				Response.Redirect("resumo.asp" & "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")))
			else
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


<%=DOCTYPE_LEGADO%>

<html>


<head>
	<title>CENTRAL</title>
	</head>



<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>



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
<% end if %>

</html>


<%
	if rs.State <> 0 then rs.Close
	set rs = nothing

'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>