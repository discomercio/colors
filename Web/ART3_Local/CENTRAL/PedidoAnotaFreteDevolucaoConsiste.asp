<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =============================================================
'	  PedidoAnotaFreteDevolucaoConsiste.asp
'     =============================================================
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

	dim s, usuario
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 
	
'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs, msg_erro
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	
' FUNÇÕES
'  obtem_empresa_NFe_emitente
function obtem_empresa_NFe_emitente(byVal id_nfe_emitente)
dim s, empresa
dim tNE

    if id_nfe_emitente = "-1" then
        obtem_empresa_NFe_emitente = "Cliente"
        exit function
    end if

	s = "SELECT" & _
			" apelido" & _
		" FROM t_NFe_EMITENTE" & _
		" WHERE" & _
			" (id = " & id_nfe_emitente & ")"
	set tNE = cn.Execute(s)
    empresa = tNE("apelido")
	obtem_empresa_NFe_emitente = empresa
	tNE.Close
	set tNE = nothing
end function

'	OBTÉM DADOS DO FORMULÁRIO
	dim c_transportadora, s_transportadora_header
	dim intCounter, intCounterAux, intQtdeItens,c_valor_nf
	dim v_item
	redim v_item(0)
	set v_item(0) = New cl_PEDIDO_ANOTA_FRETE
	c_transportadora = Trim(Request.Form("c_transportadora"))
	intQtdeItens = Request.Form("c_pedido").Count
	for intCounter = 1 to intQtdeItens
		if (Trim(Request.Form("c_pedido")(intCounter)) <> "") Or (Trim(Request.Form("c_NF")(intCounter)) <> "") then
			if (Trim(v_item(Ubound(v_item)).pedido) <> "") Or (Trim(v_item(Ubound(v_item)).num_NF) <> "") then
				redim preserve v_item(Ubound(v_item)+1)
				set v_item(Ubound(v_item)) = New cl_PEDIDO_ANOTA_FRETE
				end if
			with v_item(Ubound(v_item))
			'	Nº NF
				.num_NF = retorna_so_digitos(Trim(Request.Form("c_NF")(intCounter)))
            '	SÉRIE NF
				.serie_NF = retorna_so_digitos(Trim(Request.Form("c_serie_NF")(intCounter)))
			'	Nº PEDIDO
				.pedido = UCase(Trim(Request.Form("c_pedido")(intCounter)))
				s = normaliza_num_pedido(.pedido)
				if s <> "" then .pedido = s
			'	VALOR DO FRETE
				s = Trim(Request.Form("c_valor_frete")(intCounter))
				.vl_frete = converte_numero(s)
            '   TIPO DE FRETE
                .tipo_frete = Request.Form("c_tipo_frete")(intCounter)
            '   EMITENTE NF
                .emitente_NF = Request.Form("c_emitente")(intCounter)
            '   VALOR DA NF
                .vl_NF_devolucao = Request.Form("c_valor_nf")(intCounter)
				end with
			end if
		next
	
'	CONSISTE DADOS DIGITADOS
	dim strCampo, strColor
	dim strPedido, strListaPedidos
	dim n_pedido, id_transportadora
	dim n_linha
	dim blnNfCadastrada
	dim blnErroFatal
	blnErroFatal = False

	dim alerta
	alerta=""

	if Trim(c_transportadora) = "" then
		alerta=texto_add_br(alerta)
		alerta = alerta & "Não foi especificada a transportadora!"
		end if
	
	for intCounter = Lbound(v_item) to Ubound(v_item)
		if alerta = "" then
			with v_item(intCounter)
				if (Trim(.pedido) <> "") Or (Trim(.num_NF) <> "") then
					if (.vl_frete < 0) then
						blnErroFatal = True
						if Trim(.pedido) <> "" then
							.msg_erro = texto_add_br(.msg_erro)
							.msg_erro = .msg_erro & "Valor de frete inválido para o pedido " & Trim(.pedido) & " (" & formata_moeda(.vl_frete) & ")"
						else
							.msg_erro = texto_add_br(.msg_erro)
							.msg_erro = .msg_erro & "Valor de frete inválido para a NF " & Trim(.num_NF) & "/" & Trim(.serie_NF) & " (" & formata_moeda(.vl_frete) & ")"
							end if
						end if
					
				'	SE FOI INFORMADO O Nº DA NF:
				'		1) SE NÃO FOI INFORMADO O Nº DO PEDIDO, TENTA LOCALIZAR ATRAVÉS DO Nº DA NF
				'		2) SE FOI INFORMADO O Nº DO PEDIDO, FAZ A CONSISTÊNCIA
                if .emitente_NF <> "-1" then
					if Trim(.num_NF) <> "" then
						s = "SELECT DISTINCT" & _
								" pedido" & _
							" FROM t_NFe_EMISSAO" & _
							" WHERE" & _
								" (NFe_numero_NF = " & Trim(.num_NF) & ")" & _
                                " AND (NFe_serie_NF = " & Trim(.serie_NF) & ")" & _
                                " AND (id_nfe_emitente = " & Trim(.emitente_NF) & ")"
								
                    if .tipo_frete = "002" then
                        s = s & " AND (tipo_NF = '0')"
                    else 
                        s = s & " AND (tipo_NF = '1')"
                    end if
							s = s &	" AND (st_anulado = 0)" & _
								" AND (codigo_retorno_NFe_T1 = 1)" & _
							" ORDER BY" & _
								" pedido"
						if rs.State <> 0 then rs.Close
						rs.open s, cn
						blnNfCadastrada = False
						n_pedido = 0
						strPedido = ""
						strListaPedidos = ""
						do while Not rs.Eof
							blnNfCadastrada = True
							if Trim("" & rs("pedido")) <> "" then
								n_pedido = n_pedido + 1
								if strPedido = "" then strPedido = Trim("" & rs("pedido"))
								if strListaPedidos <> "" then strListaPedidos = strListaPedidos & ", "
								strListaPedidos = strListaPedidos & .pedido
								end if
							rs.MoveNext
							loop
						
						if n_pedido = 0 then

							if Not blnNfCadastrada then
								blnErroFatal = True                                
								.msg_erro = texto_add_br(.msg_erro)
								.msg_erro = .msg_erro & "A NF nº " & Trim(.num_NF) & "/" & Trim(.serie_NF) & " do emitente " & obtem_empresa_NFe_emitente(.emitente_NF) & " não está cadastrada!"
							
								end if
							end if
						
						if n_pedido = 1 then
							if Trim(.pedido) = "" then
								.pedido = strPedido
							else
								if Trim(.pedido) <> strPedido then
									blnErroFatal = True
									.msg_erro = texto_add_br(.msg_erro)
									.msg_erro = .msg_erro & "A NF " & Trim(.num_NF) & "/" & Trim(.serie_NF) & " do emitente " & obtem_empresa_NFe_emitente(.emitente_NF) & " refere-se a outro pedido (" & strPedido & ")"
									end if
								end if
							end if
						
						if n_pedido > 1 then
							blnErroFatal = True
							.msg_erro = texto_add_br(.msg_erro)
							.msg_erro = .msg_erro & "A NF " & Trim(.num_NF) & "/" & Trim(.serie_NF) & " está associada a mais do que um pedido (" & strListaPedidos & ")"
							end if
						end if 'if Trim(.num_NF) <> ""
					
					.id_cliente = ""
					.transportadora_id = ""
					
				'	SE NÃO FOI POSSÍVEL DETERMINAR O Nº PEDIDO ATRAVÉS DA NF, ASSEGURA QUE HAVERÁ UM MENSAGEM DE ERRO
					if Trim(.pedido) = "" then
                        blnErroFatal = True
						if Trim(.msg_erro) = "" then .msg_erro = "Falha ao tentar determinar o nº do pedido! Informe o número do pedido!"
						end if
				  end if ' if id_nfe_emitente <> "-1"

					if (alerta = "") And (.pedido <> "") then
					'	VERIFICA SE O PEDIDO ESTÁ CADASTRADO E REALIZA CONSISTÊNCIAS
						s = "SELECT " & _
								"id_cliente, " & _
								"st_entrega, " & _
								"id_cliente, " & _
								"transportadora_id, " & _
								"frete_status, " & _
								"frete_valor, " & _
								"frete_data, " & _
								"frete_usuario, " & _
								"id_nfe_emitente, " & _
								"obs_2, " & _            
								"(SELECT Sum(qtde*preco_NF) FROM t_PEDIDO_ITEM WHERE t_PEDIDO_ITEM.pedido=t_PEDIDO.pedido) AS vl_total_NF" & _
							" FROM t_PEDIDO" & _
							" WHERE" & _
								" (pedido = '" & Trim(.pedido) & "')"
						if rs.State <> 0 then rs.Close
						rs.open s, cn
						if rs.Eof then
							blnErroFatal = True
							.msg_erro = texto_add_br(.msg_erro)
							.msg_erro = .msg_erro & "Pedido " & Trim(.pedido) & " NÃO está cadastrado."
						else
                            if .num_NF = "" then .num_NF = rs("obs_2")
							.id_cliente = Trim("" & rs("id_cliente"))
                            .transportadora_id = Trim("" & rs("transportadora_id"))
							.vl_total_NF = rs("vl_total_NF")
							if .vl_total_NF = 0 then
								.perc_frete_x_total_NF = 0
							else
								.perc_frete_x_total_NF = 100 * (.vl_frete / .vl_total_NF)
                                
                                ' COMPARA O VALOR TOTAL DA NOTA FISCAL PARA VERIFICAR SE É UMA DEVOLUÇÃO TOTAL OU PARCIAL
                                if .vl_NF_devolucao <> "" then
                                    if CCur(.vl_total_NF) > CCur(.vl_NF_devolucao) then
                                        .msg_alerta = texto_add_br(.msg_alerta)
                                        .msg_alerta = .msg_alerta & "Valor da NF informado difere do valor total da NF cadastrada com o pedido " & .pedido & "."
                                    end if
                                    if CCur(.vl_total_NF) < CCur(.vl_NF_devolucao) then
                                        blnErroFatal = True
                                        .msg_erro = texto_add_br(.msg_erro)
                                        .msg_erro = .msg_erro & "Valor da NF informado é maior do que o valor da NF cadastrada com o pedido " & .pedido & "."
                                    end if
                                end if
						    end if
                            if .emitente_NF <> "-1" then
							if Trim("" & rs("st_entrega")) = ST_ENTREGA_CANCELADO then
								blnErroFatal = True
								.msg_erro = texto_add_br(.msg_erro)
								.msg_erro = .msg_erro & "Pedido " & Trim(.pedido) & " possui status inválido para esta operação (" & x_status_entrega(Trim("" & rs("st_entrega"))) & ")"
								end if
							if Trim("" & rs("transportadora_id")) = "" then
								blnErroFatal = True
								.msg_erro = texto_add_br(.msg_erro)
								.msg_erro = .msg_erro & "Pedido " & Trim(.pedido) & " NÃO está alocado para nenhuma transportadora."
								end if
							
							if rs.State <> 0 then rs.Close    
                            s = "SELECT codigo_tipo_frete " & _
	                            ",vl_frete " & _
	                            ",transportadora_id " & _
	                            ",numero_NF " & _
                                ",dt_cadastro " & _
                            "FROM t_PEDIDO_FRETE " & _
                            "WHERE pedido = '" & Trim(.pedido) & "'"
							
                            rs.open s, cn            
                            if rs.Eof then
                                if (Ucase(.transportadora_id) <> Ucase(c_transportadora)) And (.transportadora_id <> "") then
								    blnErroFatal = True
							        .msg_erro = texto_add_br(.msg_erro)
								    .msg_erro = .msg_erro & "Pedido " & Trim(.pedido) & " está alocado para outra transportadora (" & .transportadora_id & ")"
                                end if
                            else               
                                do while Not rs.Eof
                                    if (.vl_frete = rs("vl_frete")) And (.tipo_frete = Trim("" & rs("codigo_tipo_frete"))) And ((Ucase(c_transportadora) = Trim("" & rs("transportadora_id")) And Trim("" & rs("transportadora_id"))<>"")) then
                                        blnErroFatal = True
                                        .msg_erro = texto_add_br(.msg_erro)
								        .msg_erro = .msg_erro & "Pedido " & Trim(.pedido) & " já possui o valor de frete de " & SIMBOLO_MONETARIO & " " & formata_moeda(rs("vl_frete")) & " (" & rs("dt_cadastro") & ") como " & obtem_descricao_tabela_t_codigo_descricao(GRUPO_T_CODIGO_DESCRICAO__PEDIDO_TIPO_FRETE, rs("codigo_tipo_frete")) & " e transportadora " & Trim("" & rs("transportadora_id"))
                                    end if
                                    if (.tipo_frete = Trim("" & rs("codigo_tipo_frete"))) And ((Ucase(c_transportadora) = Trim("" & rs("transportadora_id")) And Trim("" & rs("transportadora_id"))<>"")) And (.msg_alerta="") then
                                        .msg_alerta = texto_add_br(.msg_alerta)								        
    
                                        if .vl_NF_devolucao <> "" then
								            .msg_alerta = .msg_alerta & "Pedido " & Trim(.pedido) & " já possui frete de " & SIMBOLO_MONETARIO & " " & formata_moeda(rs("vl_frete")) & " e valor NF de " & SIMBOLO_MONETARIO & " " & formata_moeda(.vl_total_NF) & " (" & rs("dt_cadastro") & ") cadastrado como " & obtem_descricao_tabela_t_codigo_descricao(GRUPO_T_CODIGO_DESCRICAO__PEDIDO_TIPO_FRETE, rs("codigo_tipo_frete")) & " e transportadora " & Trim("" & rs("transportadora_id"))
                                        else
                                            .msg_alerta = .msg_alerta & "Pedido " & Trim(.pedido) & " já possui frete de " & SIMBOLO_MONETARIO & " " & formata_moeda(rs("vl_frete")) & " (" & rs("dt_cadastro") & ") cadastrado como " & obtem_descricao_tabela_t_codigo_descricao(GRUPO_T_CODIGO_DESCRICAO__PEDIDO_TIPO_FRETE, rs("codigo_tipo_frete")) & " e transportadora " & Trim("" & rs("transportadora_id"))
                                        end if

                                    end if
                                    if UCase(c_transportadora) <> Trim("" & rs("transportadora_id")) And (.msg_alerta="") then 
                                        if .msg_alerta<>"" then 
                                            .msg_alerta = texto_add_br(.msg_alerta)
                                            .msg_alerta = .msg_alerta & "\n(" & Trim("" & rs("transportadora_id")) & ")"
                                        else
                                        .msg_alerta = texto_add_br(.msg_alerta)
                                        if .vl_NF_devolucao <> "" then
                                                .msg_alerta = .msg_alerta & "Pedido " & Trim(.pedido) & " já possui frete de " & SIMBOLO_MONETARIO & " " & formata_moeda(rs("vl_frete")) & " e valor NF de " & SIMBOLO_MONETARIO & " " & formata_moeda(.vl_total_NF) & " (" & rs("dt_cadastro") & ") cadastrado para outra transportadora (" & Trim("" & rs("transportadora_id")) & ")"
                                            else
                                                .msg_alerta = .msg_alerta & "Pedido " & Trim(.pedido) & " já possui frete de " & SIMBOLO_MONETARIO & " " & formata_moeda(rs("vl_frete")) & " (" & rs("dt_cadastro") & ") cadastrado para outra transportadora (" & Trim("" & rs("transportadora_id")) & ")"
                                            end if
                                        end if
                                    end if
                                    rs.MoveNext
                                loop
                            end if 'if .emitente_NF <> "-1"
                            end if
						end if
						if .emitente_NF <> "-1" then
						if .id_cliente <> "" then
							s = "SELECT " & _
									"nome" & _
								" FROM t_CLIENTE" & _
								" WHERE" & _
									" (id = '" & .id_cliente & "')"
							if rs.State <> 0 then rs.Close
							rs.open s, cn
							if rs.Eof then
								blnErroFatal = True
								.msg_erro = texto_add_br(.msg_erro)
								.msg_erro = .msg_erro & "Falha ao tentar localizar os dados do cliente do pedido " & Trim(.pedido)
							else
								.nome_cliente = Trim("" & rs("nome"))
								end if
							end if
						
						if .transportadora_id <> "" then
							s = "SELECT " & _
									"nome" & _
                                    ",cnpj" & _
								" FROM t_TRANSPORTADORA" & _
								" WHERE" & _
									" (id = '" & c_transportadora & "')"
							if rs.State <> 0 then rs.Close
							rs.open s, cn
							if rs.Eof then
								blnErroFatal = True
								.msg_erro = texto_add_br(.msg_erro)
								.msg_erro = .msg_erro & "Falha ao tentar localizar os dados da transportadora do pedido " & Trim(.pedido)
							else
								.nome_transportadora = Trim("" & rs("nome"))
                                .transportadora_cnpj = Trim("" & rs("cnpj"))
								end if
							end if
						end if 'if alerta = ""
                    end if 'if .emitente_NF <> "-1" 
					end if  'if (tem item)
				end with
			end if  'if (alerta)
		next
	
'	VERIFICA SE HÁ PEDIDOS REPETIDOS
	if alerta = "" then
		for intCounter = Lbound(v_item) to Ubound(v_item)
			with v_item(intCounter)
				if .pedido <> "" then
					for intCounterAux=Lbound(v_item) to (intCounter-1)
						if (.pedido = v_item(intCounterAux).pedido) then
							alerta=texto_add_br(alerta)
							alerta=alerta & "Pedido " & .pedido & ": linha " & renumera_com_base1(Lbound(v_item),intCounter) & " repete o mesmo pedido da linha " & renumera_com_base1(Lbound(v_item),intCounterAux) & "."
							exit for
							end if
						next
					end if
				end with
			next
		end if
	
'	VERIFICA SE HÁ NF's REPETIDAS
	if alerta = "" then
		for intCounter = Lbound(v_item) to Ubound(v_item)
			with v_item(intCounter)
				if .num_NF <> "" then
					for intCounterAux=Lbound(v_item) to (intCounter-1)
						if (.num_NF = v_item(intCounterAux).num_NF) then
							alerta=texto_add_br(alerta)
							alerta=alerta & "NF Nº " & .num_NF & "/" & Trim(.serie_NF) & ": linha " & renumera_com_base1(Lbound(v_item),intCounter) & " repete a mesma NF da linha " & renumera_com_base1(Lbound(v_item),intCounterAux) & "."
							exit for
							end if
						next
					end if
				end with
			next
		end if
	
	if alerta = "" then
		s_transportadora_header = Ucase(c_transportadora)
		s = Ucase(x_transportadora(c_transportadora))
		if (s_transportadora_header <> "") And (Ucase(s_transportadora_header) <> Ucase(s)) then s_transportadora_header = s_transportadora_header & " - " & s
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
function fOPConfirma( f ) {
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

<style TYPE="text/css">
.CelNumLinha
{
	vertical-align: top;
}
.CelNumNF {
	width: 65px;
	vertical-align: top;
}
.CelPedido {
	width: 65px;
	vertical-align: top;
	}
.CelFrete {
	width: 75px;
	vertical-align: top;
	}
.CelVlTotalNf {
	width: 75px;
	vertical-align: top;
	}
.CelPercFreteXTotalNf {
	width: 50px;
	vertical-align: top;
	}
.CelTransp {
	width: 120px;
	vertical-align: top;
	}
.CelCliente {
	width: 150px;
	vertical-align: top;
	}
.CelMsgAlerta {
	width: 223px;
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
<!-- *************************************************************** -->
<!-- **********  PÁGINA PARA EXIBIR DADOS DE CONFIRMAÇÃO  ********** -->
<!-- *************************************************************** -->
<body onload="focus();">
<center>

<form id="fOP" name="fOP" method="post" action="PedidoAnotaFreteConfirma.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="c_transportadora" id="c_transportadora" value="<%=c_transportadora%>">
<!-- FORÇA A CRIAÇÃO DE UM ARRAY MESMO QUANDO EXISTE SOMENTE 1 ITEM -->
<input type="hidden" name="c_NF" id="c_NF" value="">
<input type="hidden" name="c_serie_NF" id="c_serie_NF" value="" />
<input type="hidden" name="c_emitente_NF" id="c_emitente_NF" value="" />
<input type="hidden" name="c_pedido" id="c_pedido" value="">
<input type="hidden" name="c_valor_frete" id="c_valor_frete" value="">
<input type="hidden" name="c_tipo_frete" id="c_tipo_frete" value="" />
<input type="hidden" name="c_transportadora_id" id="c_transportadora_id" value="" />
<input type="hidden" name="c_transportadora_cnpj" id="c_transportadora_cnpj" value="" />
<input type="hidden" name="c_valor_nf" id="c_valor_nf" value="" />
<input type="hidden" name="c_valor_total_nf" id="c_valor_total_nf" value="" />
<input type="hidden" name="c_url_origem" id="c_url_origem" value="PedidoAnotaFreteDevolucao.asp" />

<%
	for intCounter=Lbound(v_item) to Ubound(v_item)
		with v_item(intCounter)
%>	
		<input type="hidden" name="c_NF" id="c_NF" value="<%=.num_NF%>">
        <input type="hidden" name="c_serie_NF" id="c_serie_NF" value="<%=.serie_NF%>" />
        <input type="hidden" name="c_emitente_NF" id="c_emitente_NF" value="<%=.emitente_NF %>" />
		<input type="hidden" name="c_pedido" id="c_pedido" value="<%=.pedido%>">
		<input type="hidden" name="c_valor_frete" id="c_valor_frete" value="<%=formata_moeda(.vl_frete)%>">
        <input type="hidden" name="c_tipo_frete" id="c_tipo_frete" value="<%=.tipo_frete%>" />
        <input type="hidden" name="c_transportadora_id" id="c_transportadora_id" value="<%=c_transportadora%>" />
        <input type="hidden" name="c_transportadora_cnpj" id="c_transportadora_cnpj" value="<%=.transportadora_cnpj%>" />
        <input type="hidden" name="c_valor_nf" id="c_valor_nf" value="<%=.vl_NF_devolucao%>" />
        <input type="hidden" name="c_valor_total_nf" id="c_valor_total_nf" value="<%=.vl_total_NF%>" />

<%		end with
	next 
%>



<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="879" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">

<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Anotar Frete no Pedido<span class="C">&nbsp;</span></span></td>
</tr>
</table>
<br>


<table class="Qx" cellSpacing="0">
	<tr>
		<td>&nbsp;</td>
		<td colspan="8" class="MT" align="left"><span class="PLTe">Transportadora: </span><span class="C"><%=s_transportadora_header%></span></td>
	</tr>
	<tr>
		<td colspan="9">&nbsp;</td>
	</tr>
	<tr>
		<td colspan="9">&nbsp;</td>
	</tr>
	<tr bgColor="#FFFFFF">
	<td align="left">&nbsp;</td>
	<td class="MB CelNumNF" align="left" style="vertical-align:bottom"><span class="PLTe">NF</span></td>
	<td class="MB CelPedido" align="left" style="vertical-align:bottom"><span class="PLTe">Pedido</span></td>
	<td class="MB CelFrete" align="right" style="vertical-align:bottom"><span class="PLTd">Frete (<%=SIMBOLO_MONETARIO%>)</span></td>
	<td class="MB CelVlTotalNf" align="right" style="vertical-align:bottom"><span class="PLTd">Valor NF (<%=SIMBOLO_MONETARIO%>)</span></td>
	<td class="MB CelPercFreteXTotalNf" align="right" style="vertical-align:bottom"><span class="PLTd">%</span></td>
	<td class="MB CelTransp" align="left" style="vertical-align:bottom"><span class="PLTe">Transportadora</span></td>
	<td class="MB CelCliente" align="left" style="vertical-align:bottom"><span class="PLTe">Cliente</span></td>
	<td class="MB CelMsgAlerta" align="left" style="vertical-align:bottom"><span class="PLTe">Alerta do Sistema</span></td>
	</tr>

<%
	n_linha = 0
	for intCounter=Lbound(v_item) to Ubound(v_item)
		n_linha = n_linha + 1
		with v_item(intCounter)
%>	
	<tr bgColor="#FFFFFF">
	<!--  Nº LINHA  -->
	<td class="CelNumLinha" align="right" valign="bottom">
		<span class="PLLd" style="margin-bottom:3px;"><%=Cstr(n_linha)%>.</span>
	</td>
	<!--  NF  -->
	<td class="MDBE CelNumNF" align="left">
		<span class="C"><%if .num_NF <> "" then Response.Write .num_NF & "/" & .serie_NF%></span>
		</td>
	<!--  PEDIDO  -->
	<td class="MDB CelPedido" align="left">
		<span class="C"><%=.pedido%></span>
		</td>

<!--  FRETE  -->
	<td class="MDB CelFrete" align="right">
		<span class="Cd"><%=formata_moeda(.vl_frete)%></span>
		</td>

<!--  VALOR TOTAL NF  -->
	<td class="MDB CelVlTotalNf" align="right">
		<span class="Cd"><%  if .vl_NF_devolucao <> "" then
                                Response.Write formata_moeda(.vl_NF_devolucao)
                             else 
                                Response.Write formata_moeda(.vl_total_NF)
                             end if%></span>
		</td>

<!--  PERC FRETE / VL TOTAL NF  -->
	<td class="MDB CelPercFreteXTotalNf" align="right">
		<span class="Cd"><%=formata_perc(.perc_frete_x_total_NF)%>%</span>
		</td>

<!--  TRANSPORTADORA  -->
	<td class="MDB CelTransp" align="left">
		<% strCampo = iniciais_em_maiusculas(Trim(.nome_transportadora))
			if strCampo = "" then strCampo = "&nbsp;"
		%>
		<span class="C"><%=strCampo%></span>
		</td>

<!--  CLIENTE  -->
	<td class="MDB CelCliente" align="left">
		<% strCampo = iniciais_em_maiusculas(Trim(.nome_cliente))
			if strCampo = "" then strCampo = "&nbsp;"
		%>
		<span class="C"><%=strCampo%></span>
		</td>

<!--  ALERTA DO SISTEMA  -->
	<td class="MDB CelMsgAlerta" align="left">
		<% 	if Trim(.msg_erro) <> "" then
				strCampo = Trim(.msg_erro)
				strColor = "red"
			elseif Trim(.msg_alerta) <> "" then
				strCampo = Trim(.msg_alerta)
				strColor = "blue"
			else
				strCampo = ""
				strColor = "black"
				end if
			if strCampo = "" then strCampo = "&nbsp;"
		%>
		<span class="C" style="color:<%=strColor%>;"><%=strCampo%></span>
		</td>

	</tr>
	
<%		end with
	next
%>
</table>


<!-- ************   SEPARADOR   ************ -->
<table width="879" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc" align="left">&nbsp;</td></tr>
</table>
<br>


<table class="notPrint" width="879" cellSpacing="0">
<tr>
	<% if Not blnErroFatal then %>
	<td align="left"><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()" title="volta para página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a>
	</td>
	<td align="right"><div name="dCONFIRMA" id="dCONFIRMA"><a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fOPConfirma(fOP)" title="confirma o cadastramento da senha">
		<img src="../botao/confirmar.gif" width="176" height="55" border="0"></a></div>
	</td>
	<% else %>
	<td align="center"><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()" title="volta para página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a>
	</td>
	<% end if %>
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