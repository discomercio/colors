<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ===============================================================================
'	  RelEstoqueDevolucaoGravaDados.asp
'     ===============================================================================
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

	class cl_TIPO_GRAVA_REL_BLOCO_NOTAS_ITEM_DEVOLVIDO
		dim id_item_devolvido
		dim pedido
		dim mensagem
		end class
		
	dim s, msg_erro
	dim usuario
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO)

	dim alerta
	alerta=""
	
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

	if Not operacao_permitida(OP_CEN_REL_PRODUTOS_ESTOQUE_DEVOLUCAO, s_lista_operacoes_permitidas) then
		Response.Redirect("aviso.asp?id=" & ERR_NIVEL_ACESSO_INSUFICIENTE)
		end if

	dim c_qtde_registros, intQtdeRegistros, vBlocoNotas
	c_qtde_registros=Trim(Request("c_qtde_registros"))
	intQtdeRegistros=CInt(c_qtde_registros)
	
	redim vBlocoNotas(0)
	set vBlocoNotas(Ubound(vBlocoNotas)) = new cl_TIPO_GRAVA_REL_BLOCO_NOTAS_ITEM_DEVOLVIDO
	vBlocoNotas(Ubound(vBlocoNotas)).id_item_devolvido = ""
	
	dim i
	dim c_id_item_devolvido, c_pedido_item_devolvido, c_nova_msg
	for i = 1 to intQtdeRegistros
		c_id_item_devolvido = Trim(Request.Form("c_id_item_devolvido_" & Cstr(i)))
		c_pedido_item_devolvido = Trim(Request.Form("c_pedido_" & Cstr(i)))
		c_nova_msg = Trim(Request.Form("c_nova_msg_" & Cstr(i)))
		if (c_id_item_devolvido<>"") And (c_nova_msg<>"") then
			if vBlocoNotas(Ubound(vBlocoNotas)).id_item_devolvido <> "" then
				redim preserve vBlocoNotas(Ubound(vBlocoNotas)+1)
				set vBlocoNotas(Ubound(vBlocoNotas)) = new cl_TIPO_GRAVA_REL_BLOCO_NOTAS_ITEM_DEVOLVIDO
				end if
			vBlocoNotas(Ubound(vBlocoNotas)).id_item_devolvido = c_id_item_devolvido
			vBlocoNotas(Ubound(vBlocoNotas)).pedido = c_pedido_item_devolvido
			vBlocoNotas(Ubound(vBlocoNotas)).mensagem = c_nova_msg
			end if
		next

	for i=Lbound(vBlocoNotas) to Ubound(vBlocoNotas)
		if Trim(vBlocoNotas(i).id_item_devolvido)<>"" then
			if len(vBlocoNotas(i).mensagem) > MAX_TAM_MENSAGEM_BLOCO_NOTAS_EM_ITEM_DEVOLVIDO then
				alerta=texto_add_br(alerta)
				alerta=alerta & "O tamanho do texto da anotação (" & Cstr(len(vBlocoNotas(i).mensagem)) & ") da devolução do pedido " & vBlocoNotas(i).pedido & " excede o tamanho máximo permitido de " & Cstr(MAX_TAM_MENSAGEM_BLOCO_NOTAS_EM_ITEM_DEVOLVIDO) & " caracteres."
				end if
			end if
		next
	
'	RECUPERA OS FILTROS USADOS NA CONSULTA P/ QUE O RELATÓRIO POSSA SER REEXECUTADO AUTOMATICAMENTE
	dim s_url
	dim c_fabricante, c_produto, c_pedido
	dim c_vendedor, c_indicador, c_captador
	dim c_lista_loja
	dim c_empresa, c_uf, c_transportadora
	c_fabricante = retorna_so_digitos(Trim(Request.Form("c_fabricante")))
	c_produto = Ucase(Trim(Request.Form("c_produto")))
	c_pedido = Ucase(Trim(Request.Form("c_pedido")))
	c_vendedor = Ucase(Trim(Request.Form("c_vendedor")))
	c_indicador = Ucase(Trim(Request.Form("c_indicador")))
	c_captador = Ucase(Trim(Request.Form("c_captador")))
	c_lista_loja = Trim(Request.Form("c_lista_loja"))
	c_uf = Trim(Request.Form("c_uf"))
	c_transportadora = Trim(Request.Form("c_transportadora"))
	c_empresa = Trim(Request.Form("c_empresa"))
	
	dim intNsuNovoRegistro
	dim campos_a_omitir
	dim vLog()
	dim s_log
	s_log = ""
	campos_a_omitir = "|dt_cadastro|dt_hr_cadastro|anulado_status|anulado_usuario|anulado_data|anulado_data_hora|"


'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

'	GRAVA A MENSAGEM NO BLOCO DE NOTAS DO ITEM DEVOLVIDO
	if alerta = "" then
	'	INICIA A TRANSAÇÃO
	'	~~~~~~~~~~~~~
		cn.BeginTrans
	'	~~~~~~~~~~~~~
		if Not cria_recordset_pessimista(rs, msg_erro) then
		'	~~~~~~~~~~~~~~~~
			cn.RollbackTrans
		'	~~~~~~~~~~~~~~~~
			Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
			end if

		for i=Lbound(vBlocoNotas) to Ubound(vBlocoNotas)
			if Trim(vBlocoNotas(i).id_item_devolvido)<>"" then
			'	TEM MENSAGEM NOVA P/ GRAVAR?
				if Trim(vBlocoNotas(i).mensagem)<>"" then
				'	GERA O NSU PARA GRAVAR A NOVA MENSAGEM
					if Not fin_gera_nsu(T_PEDIDO_ITEM_DEVOLVIDO_BLOCO_NOTAS, intNsuNovoRegistro, msg_erro) then
						alerta = "FALHA AO GERAR NSU PARA O NOVO REGISTRO (" & msg_erro & ")"
					else
						if intNsuNovoRegistro <= 0 then
							alerta = "NSU GERADO É INVÁLIDO (" & intNsuNovoRegistro & ")"
							end if
						end if
					
					if alerta = "" then
						s = "SELECT * FROM t_PEDIDO_ITEM_DEVOLVIDO_BLOCO_NOTAS WHERE (id = -1)"
						rs.Open s, cn
						rs.AddNew 
						rs("id")=intNsuNovoRegistro
						rs("id_item_devolvido")=vBlocoNotas(i).id_item_devolvido
						rs("usuario")=usuario
						rs("mensagem")=Trim(vBlocoNotas(i).mensagem)
						rs.Update 
						if Err <> 0 then
						'	~~~~~~~~~~~~~~~~
							cn.RollbackTrans
						'	~~~~~~~~~~~~~~~~
							Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
							end if

						log_via_vetor_carrega_do_recordset rs, vLog, campos_a_omitir
						s_log = log_via_vetor_monta_inclusao(vLog)
						
						if rs.State <> 0 then rs.Close
						
						if s_log <> "" then grava_log usuario, "", vBlocoNotas(i).pedido, "", OP_LOG_PEDIDO_ITEM_DEVOLVIDO_BLOCO_NOTAS_INCLUSAO, s_log
						end if
					end if  'if Trim(vBlocoNotas(i).mensagem)<>""
				end if  'if Trim(vBlocoNotas(i).id_item_devolvido)<>""
			next
			
		if alerta = "" then
		'	~~~~~~~~~~~~~~
			cn.CommitTrans
		'	~~~~~~~~~~~~~~
			if Err=0 then 
			'	FILTROS USADOS NA CONSULTA
				s_url = "RelEstoqueDevolucaoExec.asp?origem=A" & _
						"&c_fabricante=" & c_fabricante & _
						"&c_produto=" & c_produto & _
						"&c_pedido=" & c_pedido & _
						"&c_vendedor=" & c_vendedor & _
						"&c_indicador=" & c_indicador & _
						"&c_captador=" & c_captador & _
						"&c_empresa=" & c_empresa & _
						"&c_uf=" & c_uf & _
						"&c_transportadora=" & c_transportadora & _
						"&c_lista_loja=" & c_lista_loja & _
						"&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))
				Response.Redirect(s_url)
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
<% end if %>

</html>


<%
'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>