<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/Global.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =================================================================
'	  R E L C O N T R O L E I M P O S T O S E X E C . A S P
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

	Server.ScriptTimeout = MAX_SERVER_SCRIPT_TIMEOUT_EM_SEG
	
	dim usuario
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	CONECTA COM O BANCO DE DADOS
	dim cn, tLock, msg_erro
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(tLock, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	if Not operacao_permitida(OP_CEN_REL_CONTROLE_IMPOSTOS, s_lista_operacoes_permitidas) then 
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

'   NÃO USA O history.back() PARA QUE A PÁGINA SEJA ACIONADA NOVAMENTE E LIMPE OS LOCKS
	dim strUrlBotaoVoltar
	strUrlBotaoVoltar = "javascript:fRELVoltarTelaFiltro(fREL)"
	
	dim alerta
	dim s, s_aux, s_filtro
	dim c_transportadora, s_nome_transportadora, c_dt_coleta, c_dt_coleta_inicio, c_dt_coleta_termino, ckb_exibir_verificados, c_nfe_emitente, c_uf

	alerta = ""

	c_dt_coleta = Trim(Request.Form("c_dt_coleta"))
	c_dt_coleta_inicio = Trim(Request.Form("c_dt_coleta_inicio"))
	c_dt_coleta_termino = Trim(Request.Form("c_dt_coleta_termino"))
	c_transportadora = Trim(Request.Form("c_transportadora"))
	ckb_exibir_verificados = Trim(Request.Form("ckb_exibir_verificados"))
	c_nfe_emitente = Trim(Request.Form("c_nfe_emitente"))
	c_uf = Trim(Request.Form("c_uf"))

	s_nome_transportadora = ""
	if c_transportadora <> "" then s_nome_transportadora = x_transportadora(c_transportadora)

	dim qtde_notas
	qtde_notas = 0

	if alerta = "" then
		if c_nfe_emitente = "" then
			alerta=texto_add_br(alerta)
			alerta = alerta & "Não foi informado o CD"
		elseif converte_numero(c_nfe_emitente) = 0 then
			alerta=texto_add_br(alerta)
			alerta = alerta & "É necessário definir um CD válido"
			end if
		end if

'   LIMPA EVENTUAIS LOCKS REMANESCENTES
    s = "UPDATE t_CTRL_RELATORIO_USUARIO_X_PEDIDO SET" & _
            " locked = 0," & _
            " cod_motivo_lock_released = " & CTRL_RELATORIO_CodMotivoLockReleased_AcessadaTelaResultado & "," & _
            " dt_hr_lock_released = getdate()" & _
        " WHERE" & _
            " (usuario = '" & QuotedStr(usuario) & "')" & _
            " AND (id_relatorio = " & ID_CTRL_RELATORIO_RelControleImpostos & ")" & _
            " AND (locked = 1)"
    cn.Execute(s)

'	OBTÉM PARÂMETROS
    dim rPTimeout, rPQtdeMax
    set rPTimeout = get_registro_t_parametro(ID_PARAMETRO_CtrlRelatorio_RelControleImpostos_TimeoutLockEmMinutos)
    set rPQtdeMax = get_registro_t_parametro(ID_PARAMETRO_CtrlRelatorio_RelControleImpostos_MaxQtdeResultadoPorConsulta)

    dim blnErroAcessoConcorrente, s_msg_erro_acesso_concorrente
    blnErroAcessoConcorrente = False
    s_msg_erro_acesso_concorrente = ""





' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________

' _____________________________________
' CONSULTA EXECUTA
'
sub consulta_executa
dim r
dim blnDisabled
dim s, s_sql, cab_table, cab, s_num_nfe, s_serie_nfe, s_link_nfe, s_row, s_html_color, s_link_habilita_nfe
dim s_pedido, s_transportadora, s_data_entrega_yyyymmdd, s_data_coleta, s_cnpj_cpf
dim s_where, s_where_data, s_from
dim i, intOrdenacaoCor
dim blnRegistroOk
dim rNfeEmitente
dim x
dim qtde_produto,total_fcp,total_icms_destino, valor
dim ChaveAcesso
dim qtdePedidosExibida, maxQtdePedidosResultado, tempoTimeoutEmMin, blnPedidoBloqueado
dim qtdeTotalPedidosATratar, intRecordsAffected
dim vLocked, idxLocked, qtdePedidoBloqueado

    total_fcp = 0
    total_icms_destino = 0
    qtdePedidosExibida = 0
    qtdeTotalPedidosATratar = 0
    qtdePedidoBloqueado = 0

    maxQtdePedidosResultado = converte_numero(rPQtdeMax.campo_inteiro)
    if maxQtdePedidosResultado <= 0 then maxQtdePedidosResultado = 10

    tempoTimeoutEmMin = converte_numero(rPTimeout.campo_inteiro)
    if tempoTimeoutEmMin <= 0 then tempoTimeoutEmMin = 20

    redim vLocked(0)
    set vLocked(UBound(vLocked)) = new cl_DUAS_COLUNAS
    vLocked(UBound(vLocked)).c1 = ""
    vLocked(UBound(vLocked)).c2 = ""

'	MONTA CLÁUSULA WHERE
	s_where = " AND (t_NFE_IMAGEM.ide__idDest = '2') " & _
			" AND (t_NFE_IMAGEM.ide__tpNF = '1') " & _
			" AND (t_NFE_EMISSAO.st_anulado = 0) " & _
			" AND (t_NFE_EMISSAO.codigo_retorno_NFe_T1 = 1) " & _
			" AND (st_entrega <> 'CAN') "

	
'	CRITÉRIO: TRANSPORTADORA
	if c_transportadora <> "" then
		s = " (transportadora_id = '" & c_transportadora & "')"
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & s
		end if

'	CRITÉRIO: UF
	if c_uf <> "" then
		s = " (dest__UF = '" & c_uf & "')"
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & s
		end if
	
'	CRITÉRIO: DATA DE COLETA
	s_where_data = ""
	if c_dt_coleta <> "" then
		s = " (a_entregar_data_marcada = " & bd_formata_data(StrToDate(c_dt_coleta)) & ")"
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & s
		s_where_data = s
		end if

'	CRITÉRIO: PERÍODO DE COLETA
	if (c_dt_coleta = "") and ((c_dt_coleta_inicio <> "") and (c_dt_coleta_termino <> "")) then
		s = " (a_entregar_data_marcada >= " & bd_formata_data(StrToDate(c_dt_coleta_inicio)) & ")"
		s = s & " AND (a_entregar_data_marcada <= " & bd_formata_data(StrToDate(c_dt_coleta_termino)) & ")"
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & s
		s_where_data = s
		end if

'	CRITÉRIO: EXIBIR VERIFICADOS
	if ckb_exibir_verificados = "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_NFE_EMISSAO.controle_impostos_status = " & COD_CONTROLE_IMPOSTOS_STATUS__INICIAL & ")"
	else
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (t_NFE_EMISSAO.controle_impostos_status IN (" & COD_CONTROLE_IMPOSTOS_STATUS__INICIAL & "," & COD_CONTROLE_IMPOSTOS_STATUS__OK & "))"
		end if
	
'	OWNER DO PEDIDO
	set rNfeEmitente = le_nfe_emitente(c_nfe_emitente)
	if s_where <> "" then s_where = s_where & " AND"
	s_where = s_where & " (t_NFE_EMISSAO.id_nfe_emitente = " & rNfeEmitente.id & ")"

'	Primeiro grupo selecionado: NFes interestaduais emitidas automaticamente
	s_sql = "SELECT" & _
				" t_NFE_EMISSAO.id," & _
				" t_NFE_EMISSAO.NFe_serie_NF," & _
				" t_NFE_EMISSAO.NFe_numero_NF," & _
				" t_NFE_EMISSAO.controle_impostos_status," & _
				" t_PEDIDO.pedido," & _
				" t_PEDIDO.a_entregar_data_marcada," & _
				" t_PEDIDO.st_entrega," & _
				" t_PEDIDO.transportadora_id," & _
				" t_NFE_IMAGEM.dest__UF," & _
				" t_NFE_IMAGEM.dest__xMun," & _
				" t_NFE_IMAGEM.dest__xNome," & _
				" t_NFE_IMAGEM.dest__CNPJ," & _
				" t_NFE_IMAGEM.dest__CPF," & _
				" t_NFE_IMAGEM.total__vFCPUFDest," & _
				" t_NFE_IMAGEM.total__vICMSUFDest," & _
				" t_NFE_IMAGEM.total__vICMSUFRemet" & _
			" FROM t_NFE_EMISSAO" & _
				" INNER JOIN t_NFE_IMAGEM ON (t_NFE_EMISSAO.NFe_numero_NF=t_NFE_IMAGEM.NFe_numero_NF AND t_NFE_EMISSAO.NFe_serie_NF=t_NFE_IMAGEM.NFe_serie_NF AND t_NFE_EMISSAO.id_nfe_emitente=t_NFE_IMAGEM.id_nfe_emitente)" & _
				" INNER JOIN (SELECT id_nfe_emitente, NFe_serie_NF, NFe_numero_NF, max(id) AS id FROM t_NFE_IMAGEM GROUP BY id_nfe_emitente, NFe_serie_NF, NFe_numero_NF) img_max_id ON (t_NFE_IMAGEM.id = img_max_id.id AND t_NFE_IMAGEM.id_nfe_emitente=img_max_id.id_nfe_emitente)" & _
				" INNER JOIN (SELECT id_nfe_emitente, NFe_serie_NF, NFe_numero_NF, max(id) AS id FROM t_NFE_EMISSAO GROUP BY id_nfe_emitente, NFe_serie_NF, NFe_numero_NF) emi_max_id ON (t_NFE_EMISSAO.id = emi_max_id.id AND t_NFE_EMISSAO.id_nfe_emitente=emi_max_id.id_nfe_emitente)" & _
				" INNER JOIN t_PEDIDO ON (t_NFE_EMISSAO.pedido=t_PEDIDO.pedido AND t_NFE_EMISSAO.id_nfe_emitente=t_PEDIDO.id_nfe_emitente)" & _
			" WHERE t_NFE_EMISSAO.pedido IS NOT NULL " & _
			s_where

'	Segundo grupo selecionado: NFes interestaduais emitidas manualmente, com um conjunto específico de CFOPs relacionados
			s_sql = s_sql & _
			"UNION " & _
			"SELECT" & _
				" t_NFE_EMISSAO.id," & _
				" t_NFE_EMISSAO.NFe_serie_NF," & _
				" t_NFE_EMISSAO.NFe_numero_NF," & _
				" t_NFE_EMISSAO.controle_impostos_status," & _
				" pedidos_nf_ok.pedido," & _
				" pedidos_nf_ok.a_entregar_data_marcada," & _
				" pedidos_nf_ok.st_entrega," & _
				" pedidos_nf_ok.transportadora_id," & _
				" t_NFE_IMAGEM.dest__UF," & _
				" t_NFE_IMAGEM.dest__xMun," & _
				" t_NFE_IMAGEM.dest__xNome," & _
				" t_NFE_IMAGEM.dest__CNPJ," & _
				" t_NFE_IMAGEM.dest__CPF," & _
				" t_NFE_IMAGEM.total__vFCPUFDest," & _
				" t_NFE_IMAGEM.total__vICMSUFDest," & _
				" t_NFE_IMAGEM.total__vICMSUFRemet" & _
			" FROM t_NFE_EMISSAO" & _
				" INNER JOIN t_NFE_IMAGEM ON (t_NFE_EMISSAO.NFe_numero_NF=t_NFE_IMAGEM.NFe_numero_NF AND t_NFE_EMISSAO.NFe_serie_NF=t_NFE_IMAGEM.NFe_serie_NF AND t_NFE_EMISSAO.id_nfe_emitente=t_NFE_IMAGEM.id_nfe_emitente)" & _
				" INNER JOIN (SELECT id_nfe_emitente, NFe_serie_NF, NFe_numero_NF, max(id) AS id FROM t_NFE_IMAGEM GROUP BY id_nfe_emitente, NFe_serie_NF, NFe_numero_NF) img_max_id ON (t_NFE_IMAGEM.id = img_max_id.id AND t_NFE_IMAGEM.id_nfe_emitente=img_max_id.id_nfe_emitente)" & _
				" INNER JOIN (SELECT id_nfe_emitente, NFe_serie_NF, NFe_numero_NF, max(id) AS id FROM t_NFE_EMISSAO GROUP BY id_nfe_emitente, NFe_serie_NF, NFe_numero_NF) emi_max_id ON (t_NFE_EMISSAO.id = emi_max_id.id AND t_NFE_EMISSAO.id_nfe_emitente=emi_max_id.id_nfe_emitente)" & _
				" INNER JOIN (SELECT * FROM t_PEDIDO WHERE (ISNUMERIC(t_PEDIDO.obs_2) = 1) AND (LEN(t_PEDIDO.obs_2) < 10) AND " & s_where_data &") pedidos_nf_ok" & _
				"		ON t_NFE_EMISSAO.NFe_numero_NF=CONVERT(INT, pedidos_nf_ok.obs_2) AND t_NFE_EMISSAO.id_nfe_emitente = pedidos_nf_ok.id_nfe_emitente" & _
			" WHERE t_NFE_EMISSAO.pedido IS NULL " & _
				" AND " & _
				" (EXISTS (SELECT 1  " & _
							" FROM t_NFe_IMAGEM_ITEM  " & _
							" WHERE t_NFE_IMAGEM.id = t_NFe_IMAGEM_ITEM.id_nfe_imagem " & _
							" AND t_NFe_IMAGEM_ITEM.det__CFOP IN ('5102','6102','6108','5119','6119','5910','6910')))" & _
			s_where
	
	s_sql = s_sql & " ORDER BY pedido, t_NFE_EMISSAO.NFe_serie_NF, t_NFE_EMISSAO.NFe_numero_NF"
	
  ' CABEÇALHO
	cab_table = "<table cellspacing=0 id='tabelaRelatorio'>" & chr(13)
	cab = "	<tr style='background:azure'>" & chr(13) & _
		  "		<td class='ME MC MD' style='width:50px' align='center' valign='bottom' nowrap><span class='R'>Pedido</span></td>" & chr(13) & _
		  "		<td class='MC MD' style='width:35px' align='center' valign='bottom' nowrap><span class='R'>Nº NF</span></td>" & chr(13) & _
		  "		<td class='MC MD' style='width:35px' align='center' valign='bottom' nowrap><span class='R'>Data Coleta</span></td>" & chr(13) & _
		  "		<td class='MC MD' style='width:140px' align='center' valign='bottom' nowrap><span class='R'>Chave de Acesso</span></td>" & chr(13) & _
		  "		<td class='MC MD' style='width:80px' align='center' valign='bottom' nowrap><span class='R'>CNPJ/CPF</span></td>" & chr(13) & _
		  "		<td class='MC MD' style='width:140px' align='left' valign='bottom' nowrap><span class='R'>Cliente</span></td>" & chr(13) & _
		  "		<td class='MC MD' style='width:100px' align='left' valign='bottom' nowrap><span class='R'>Transportadora</span></td>" & chr(13) & _
		  "		<td class='MC MD' style='width:100px' align='left' valign='bottom' nowrap><span class='R'>Cidade</span></td>" & chr(13) & _
        "		<td class='MC MD' style='width:25px' align='center' valign='bottom' nowrap><span class='R'>UF</span></td>" & chr(13) & _
		  "		<td class='MC MD' style='width:30px' align='right' valign='bottom' nowrap><span class='R'>FCP</span></td>" & chr(13) & _
		  "		<td class='MC MD' style='width:55px' align='right' valign='bottom' nowrap><span class='R'>ICMS UF </br> Destino</span></td>" & chr(13) & _
		  "		<td class='MC MD' style='width:30px' align='center' valign='bottom' nowrap><span class='R'>Guia </br> OK</span></td>" & chr(13) & _
		  "	</tr>" & chr(13)

	x = cab_table & cab

	set r = cn.execute(s_sql)
	do while Not r.Eof
        blnPedidoBloqueado = False
		
	'	SE A NOTA NÃO FOI COMPLETAMENTE EMITIDA, PULAR
		
		s_num_nfe = NFeFormataNumeroNF(Trim("" & r("NFe_numero_NF")))
		s_serie_nfe = NFeFormataSerieNF(Trim("" & r("NFe_serie_NF")))

		if IsNFeCompletamenteEmitida(rNfeEmitente.id, s_serie_nfe, s_num_nfe, ChaveAcesso) then
            qtdeTotalPedidosATratar = qtdeTotalPedidosATratar + 1

            s = "SELECT " & _
                    "*" & _
                " FROM t_CTRL_RELATORIO_USUARIO_X_PEDIDO" & _
                " WHERE" & _
                    "(id_relatorio = " & ID_CTRL_RELATORIO_RelControleImpostos & ")" & _
                    " AND (usuario <> '" & QuotedStr(usuario) & "')" & _
                    " AND (pedido = '" & Trim("" & r("pedido")) & "')" & _
                    " AND (locked = 1)" & _
                    " AND (data_hora >= DATEADD(minute, -" & tempoTimeoutEmMin & ", getdate()))" & _
                " ORDER BY" & _
                    " data_hora DESC"
			if tLock.State <> 0 then tLock.Close
			tLock.Open s, cn
            if Not tLock.Eof then
                blnPedidoBloqueado = True
                qtdePedidoBloqueado = qtdePedidoBloqueado + 1

            '   COLETA DADOS P/ EXIBIR PEDIDOS BLOQUEADOS POR OUTROS USUÁRIOS
                idxLocked = -1
                for i=LBound(vLocked) to UBound(vLocked)
                    if Trim("" & vLocked(i).c1) = Trim("" & tLock("usuario")) then
                        idxLocked = i
                        exit for
                        end if
                    next

            '   JÁ EXISTE NO VETOR UMA POSIÇÃO P/ O USUÁRIO QUE ESTÁ CAUSANDO O BLOQUEIO?
                if idxLocked = -1 then
                    if vLocked(UBound(vLocked)).c1 <> "" then
                        redim preserve vLocked(UBound(vLocked)+1)
                        set vLocked(UBound(vLocked)) = new cl_DUAS_COLUNAS
                        vLocked(UBound(vLocked)).c1 = ""
                        vLocked(UBound(vLocked)).c2 = ""
                        end if
                    vLocked(UBound(vLocked)).c1 = Trim("" & tLock("usuario"))
                    idxLocked = UBound(vLocked)
                    end if

                if vLocked(idxLocked).c2 <> "" then vLocked(idxLocked).c2 = vLocked(idxLocked).c2 & ", "
                vLocked(idxLocked).c2 = vLocked(idxLocked).c2 & Trim("" & tLock("pedido"))
                end if 'if Not tLock.Eof

            if (Not blnPedidoBloqueado) And (qtdePedidosExibida < maxQtdePedidosResultado) then
            '   CRIA O REGISTRO P/ BLOQUEAR O PEDIDO NESTE RELATÓRIO SENDO EXECUTADO POR OUTROS USUÁRIOS
                s = "INSERT INTO t_CTRL_RELATORIO_USUARIO_X_PEDIDO (" & _
                        "id_relatorio" & _
                        ", usuario" & _
                        ", pedido" & _
                        ", locked" & _
                    ") VALUES (" & _
                        ID_CTRL_RELATORIO_RelControleImpostos & _
                        ", '" & QuotedStr(usuario) & "'" & _
                        ", '" & Trim("" & r("pedido")) & "'" & _
                        ", 1" & _
                    ")"
                cn.Execute s, intRecordsAffected
                if intRecordsAffected <> 1 then
                    blnErroAcessoConcorrente = True
                    s_msg_erro_acesso_concorrente = "FALHA AO TENTAR REGISTRAR BLOQUEIO DO USUÁRIO NO PEDIDO " & Trim("" & r("pedido"))
                    exit sub
                    end if

                'DEVIDO À POSSIBILIDADE DE ACESSO CONCORRENTE, VERIFICA SE OUTRO USUÁRIO TAMBÉM INSERIU REGISTRO DE BLOQUEIO
                s = "SELECT " & _
                        "*" & _
                    " FROM t_CTRL_RELATORIO_USUARIO_X_PEDIDO" & _
                    " WHERE" & _
                        "(id_relatorio = " & ID_CTRL_RELATORIO_RelControleImpostos & ")" & _
                        " AND (usuario <> '" & QuotedStr(usuario) & "')" & _
                        " AND (pedido = '" & Trim("" & r("pedido")) & "')" & _
                        " AND (locked = 1)" & _
                        " AND (data_hora >= DATEADD(minute, -" & tempoTimeoutEmMin & ", getdate()))" & _
                    " ORDER BY" & _
                        " data_hora DESC"
			    if tLock.State <> 0 then tLock.Close
			    tLock.Open s, cn
                if Not tLock.Eof then
                    'OCORREU CONFLITO DE ACESSO CONCORRENTE, POIS OUTRO USUÁRIO TAMBÉM INSERIU REGISTRO DE BLOQUEIO
                    blnErroAcessoConcorrente = True
                    s_msg_erro_acesso_concorrente = "FALHA DEVIDO A ACESSO CONCORRENTE: O PEDIDO " & Trim("" & tLock("pedido")) & " FOI BLOQUEADO SIMULTANEAMENTE POR '" & usuario & "' E '" & Trim("" & tLock("usuario")) & "'"
                    'INTERROMPE GERAÇÃO DO RELATÓRIO (A PÁGINA DEVE EXIBIR A MENSAGEM DE ERRO, LIMPAR OS LOCKS E EXIBIR O BOTÃO P/ RETORNAR)
                    exit sub
                    end if

		        '	CONTAGEM
                qtdePedidosExibida = qtdePedidosExibida + 1

			    intOrdenacaoCor = 0
			    s_html_color = "black"
			
			    s_html_color = " style='color:" & s_html_color & ";'"
			
		    '	MONTA O HTML DA LINHA DA TABELA
		    '	===============================
			    s_row = "	<tr onmouseover='realca_cor_mouse_over(this);' onmouseout='realca_cor_mouse_out(this);'>" & chr(13)

		    '> Nº PEDIDO
			    s_pedido = Trim("" & r("pedido"))
			    s = s_pedido
			    if s = "" then s = "&nbsp;"
			    s_row = s_row & _
					    "		<td align='center' valign='top' class='ME MC MD'>" & chr(13) & _
					    "			<span class='C'" & s_html_color & ">&nbsp;<a href='javascript:fRELConcluir(" & chr(34) & Trim("" & r("pedido")) & chr(34) & ")' tabindex=-1 title='clique para consultar o pedido'" & s_html_color & ">" & Trim("" & r("pedido")) & "</a></span>" & chr(13) & _
					    "			<input type='hidden' name='c_numero_pedido' value='" & s & "' />" & chr(13) & _
					    "		</td>" & chr(13)

		    '	Nº NFe
			    s_num_nfe = NFeFormataNumeroNF(Trim("" & r("NFe_numero_NF")))
			    if s_num_nfe <> "" then
				    s = "<span class='C'" & s_html_color & ">" & NFeFormataNumeroNF(s_num_nfe) & "</span>"
				    s_link_nfe = monta_link_para_DANFE(s_pedido, MAX_PERIODO_LINK_DANFE_DISPONIVEL_NO_PEDIDO_EM_DIAS, s)
				    s_link_habilita_nfe = s_link_nfe
				    if s_link_nfe = "" then s_link_nfe = "<span class='C' style='color:gray;'>" & s_num_nfe & "</span>"
			    else
				    s_link_nfe = "&nbsp;"
				    end if
				
			    s_row = s_row & _
					    "		<td align='center' valign='top' class='MC MD'>" & chr(13) & _
					    "			" & s_link_nfe & chr(13) & _
					    "		</td>" & chr(13)

            '	DATA DE COLETA
			    s_data_coleta = Trim("" & r("a_entregar_data_marcada"))
			    s = s_data_coleta
			    if s = "" then s = "&nbsp;"
			    s_row = s_row & _
					    "		<td align='center' valign='top' class='MC MD'>" & chr(13) & _
					    "			<span class='C'" & s_html_color & ">" & s & "</span>" & chr(13) & _
					    "		</td>" & chr(13)

            '	CHAVE ACESSO
			    s = ChaveAcesso
			    if s = "" then s = "&nbsp;"
			    s_row = s_row & _
					    "		<td align='center' valign='top' class='MC MD'>" & chr(13) & _
					    "			<input type='text' class='PLLd TxtClip' readonly style='width:130px' name='c_chave_acesso' onclick='this.select();' value='" & s & "' />" & chr(13) & _
					    "		</td>" & chr(13)

            '   CNPJ/CPF
                s_cnpj_cpf = cnpj_formata(Trim("" & r("dest__CNPJ")))
                if s_cnpj_cpf = "" then s_cnpj_cpf = cpf_formata(Trim("" & r("dest__CPF")))
                if s_cnpj_cpf = "" then s_cnpj_cpf = "&nbsp;"
			    s_row = s_row & _
					    "		<td align='center' valign='top' class='MC MD'>" & chr(13) & _
					    "			<input type='text' class='PLLd TxtClip' readonly style='text-align:center' name='c_cnpj_cpf' onclick='this.select();' value='" & s_cnpj_cpf & "' />" & chr(13) & _
					    "		</td>" & chr(13)

            '	CLIENTE
			    s = Trim("" & r("dest__xNome"))
			    if s = "" then s = "&nbsp;"
			    s_row = s_row & _
					    "		<td align='left' valign='top' class='MC MD' nowrap>" & chr(13) & _
					    "			<input type='text' class='PLLd TxtClip' readonly style='width:130px;text-align:left' name='c_nome_cliente' onclick='this.select();' value='" & s & "' />" & chr(13) & _
					    "		</td>" & chr(13)


		    '	TRANSPORTADORA
			    s = Trim("" & r("transportadora_id"))
			    if s = "" then s = "&nbsp;"
			    s_row = s_row & _
					    "		<td align='left' valign='top' class='MC MD'>" & chr(13) & _
					    "			<span class='C'" & s_html_color & ">" & s & "</span>" & chr(13) & _
					    "			<input type='hidden' name='c_pedido_transportadora' value='" & Trim("" & r("transportadora_id")) & "' />" & chr(13) & _
					    "		</td>" & chr(13)

            '	CIDADE
			    s = Trim("" & r("dest__xMun"))
			    if s = "" then s = "&nbsp;"
			    s_row = s_row & _
					    "		<td align='left' valign='top' class='MC MD' nowrap>" & chr(13) & _
					    "			<input type='text' class='PLLd TxtClip' readonly style='width:130px;text-align:left' name='c_cidade' onclick='this.select();' value='" & s & "' />" & chr(13) & _
					    "		</td>" & chr(13)


            '	UF
			    s = Trim("" & r("dest__UF"))
			    if s = "" then s = "&nbsp;"
			    s_row = s_row & _
					    "		<td align='center' valign='top' class='MC MD'>" & chr(13) & _
					    "			<span class='C'" & s_html_color & ">" & s & "</span>" & chr(13) & _
					    "		</td>" & chr(13)

		    '	FCP
			    valor = converte_numero(Trim("" & r("total__vFCPUFDest")))
			    s = formata_moeda(valor)
			    if s = "" then s = "&nbsp;"
			    s_row = s_row & _
					    "		<td align='right' valign='top' class='MC MD'>" & chr(13) & _
					    "			<span class='C'" & s_html_color & ">" & s & "</span>" & chr(13) & _
					    "		</td>" & chr(13)

		    '	ICMS DESTINO
			    valor = converte_numero(Trim("" & r("total__vICMSUFDest")))
			    s = formata_moeda(valor)
			    if s = "" then s = "&nbsp;"
			    s_row = s_row & _
					    "		<td align='right' valign='top' class='MC MD'>" & chr(13) & _
					    "			<span class='C'" & s_html_color & ">" & s & "</span>" & chr(13) & _
					    "		</td>" & chr(13)

		    '	GUIA OK
			    if r("controle_impostos_status") = CInt(COD_CONTROLE_IMPOSTOS_STATUS__OK) then s = "S" else s = "N"
			    s_row = s_row & _
				    "		<td align='center' valign='top' class='MC MD' style='padding:0px;'>" & chr(13)
			    s_row = s_row & _
			    "			<input type='checkbox' name='ckb_controle_impostos' id='ckb_controle_impostos' value='" & Trim("" & r("id")) & "|" & s_num_nfe & "|"  & s_pedido & "|" & s & "'"

			    if s = "S" then s_row = s_row & " checked disabled"

			    s_row = s_row & chr(13)

			    s_row = s_row & _
				    "		</td>" & chr(13)

			    s_row = s_row & "	</tr>" & chr(13)

			    x = x & s_row
			    if (qtdePedidosExibida mod 100) = 0 then
				    Response.Write x
				    x = ""
				    end if

			    total_fcp = total_fcp + CCur(converte_numero(r("total__vFCPUFDest")))
			    total_icms_destino = total_icms_destino + CCur(converte_numero(r("total__vICMSUFDest")))
			
                end if 'if (Not blnPedidoBloqueado) And (qtdePedidosExibida < maxQtdePedidosResultado)
			end if 'if IsNFeCompletamenteEmitida()
		
		r.MoveNext
		loop
	
    'VERIFICAÇÃO FINAL SE HOUVE CONFLITO CAUSADO POR ACESSO CONCORRENTE
    s = "SELECT " & _
            "*" & _
        " FROM t_CTRL_RELATORIO_USUARIO_X_PEDIDO" & _
        " WHERE" & _
            " (id_relatorio = " & ID_CTRL_RELATORIO_RelControleImpostos & ")" & _
	        " AND (usuario <> '" & QuotedStr(usuario) & "')" & _
	        " AND (locked = 1)" & _
	        " AND (data_hora >= DATEADD(minute, -" & tempoTimeoutEmMin & ", getdate()))" & _
	        " AND (" & _
		        "pedido IN (" & _
			        "SELECT pedido" & _
			        " FROM t_CTRL_RELATORIO_USUARIO_X_PEDIDO" & _
			        " WHERE (id_relatorio = " & ID_CTRL_RELATORIO_RelControleImpostos & ")" & _
				    " AND (usuario = '" & QuotedStr(usuario) & "')" & _
				    " AND (locked = 1)" & _
				    " AND (data_hora >= DATEADD(minute, -" & tempoTimeoutEmMin & ", getdate()))" & _
			        ")" & _
		        ")" & _
        " ORDER BY data_hora DESC"
	if tLock.State <> 0 then tLock.Close
	tLock.Open s, cn
    if Not tLock.Eof then
        'OCORREU CONFLITO DE ACESSO CONCORRENTE, POIS OUTRO USUÁRIO TAMBÉM INSERIU REGISTRO DE BLOQUEIO
        blnErroAcessoConcorrente = True
        do while Not tLock.Eof
            if s_msg_erro_acesso_concorrente <> "" then s_msg_erro_acesso_concorrente = s_msg_erro_acesso_concorrente & ", "
            s_msg_erro_acesso_concorrente = s_msg_erro_acesso_concorrente & Trim("" & tLock("pedido")) & " (" & Trim("" & tLock("usuario")) & ")"
            tLock.MoveNext
            loop

        s_msg_erro_acesso_concorrente = "FALHA DEVIDO A ACESSO CONCORRENTE: OS SEGUINTES PEDIDOS FORAM BLOQUEADOS SIMULTANEAMENTE POR OUTROS USUÁRIOS: " & s_msg_erro_acesso_concorrente
        'INTERROMPE GERAÇÃO DO RELATÓRIO (A PÁGINA DEVE EXIBIR A MENSAGEM DE ERRO, LIMPAR OS LOCKS E EXIBIR O BOTÃO P/ RETORNAR)
        exit sub
        end if

	x = x & "	<tr style='background: #FFFFDD'>" & chr(13) & _
			"       <td class='MTBE' nowrap colspan=9' align='right'><span class='C'>" & _
			"TOTAL:   </span></td>" & chr(13) & _
			"       <td class='MTB' nowrap colspan=1' align='right'><span class='C'>"& formata_moeda(total_fcp) & "</span></td>" & chr(13) & _
			"       <td class='MTB' nowrap colspan=1' align='right'><span class='C'>"& formata_moeda(total_icms_destino) & "</span></td>" & chr(13) & _
			"		<td class='MTBD' nowrap colspan=2' align='right'><span class='C'>" & _
			"	</tr>" & chr(13)

' MOSTRA O TOTAL DE REGISTROS
	x = x & "	<tr style='background: #FFFFDD'>" & chr(13) & _
			"       <td class='MDBE' nowrap colspan=12' align='right'>" & _
            "       <table cellspacing=0 cellpadding=0 style='width:100%;' border='0'>" & chr(13) & _
            "           <tr>" & chr(13) & _
            "               <td class='MB' style='text-align:right;'><span class='C'>TOTAL DE REGISTRO(S) EXIBIDO:</span></td>" & chr(13) & _
            "               <td class='MB' style='text-align:right;width:40px;'><span class='C'>" & formata_inteiro(qtdePedidosExibida) & "</span></td>" & chr(13) & _
            "           </tr>" & chr(13) & _
            "           <tr>" & chr(13) & _
            "               <td style='text-align:right;'><span class='C'>TOTAL DE REGISTRO(S) NO RESULTADO:</span></td>" & chr(13) & _
            "               <td style='text-align:right;width:40px;'><span class='C'>" & formata_inteiro(qtdeTotalPedidosATratar) & "</span></td>" & chr(13) & _
            "           </tr>" & chr(13) & _
            "       </table>" & chr(13) & _
            "       </td>" & chr(13) & _
			"	</tr>" & chr(13)

  ' MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	if qtdePedidosExibida = 0 then
		x = cab_table & cab
		x = x & "	<tr>" & chr(13) & _
				"		<td class='MT ALERTA' colspan='15' align='center'><span class='ALERTA'>&nbsp;NENHUM PEDIDO ENCONTRADO&nbsp;</span></td>" & chr(13) & _
				"	</tr>" & chr(13) & _
                "   <tr>" & chr(13) & _
				"		<td class='MDBE ALERTA' colspan='15' align='center'><span class='ALERTA'>&nbsp;TOTAL DE REGISTRO(S) NO RESULTADO: &nbsp; " & formata_inteiro(qtdeTotalPedidosATratar) & " &nbsp;</span></td>" & chr(13) & _
				"	</tr>" & chr(13)
		end if

  ' FECHA TABELA
	x = x & "</table>" & chr(13)
	
	Response.write x
    x = ""

    if qtdePedidoBloqueado > 0 then
        'EXIBE PEDIDOS BLOQUEADOS POR OUTROS USUÁRIOS
        x = x & "<br /><br />" & chr(13) & _
                "<table cellspacing=0 id='tabelaLock' style='min-width:600px;'>" & chr(13) & _
                "   <thead>" & chr(13) & _
                "   <tr>" & chr(13) & _
                "       <th colspan='2' class='MB'><span class='C'>Pedido(s) bloqueado(s) por outro(s) usuário(s)</span></th>" & chr(13) & _
                "   </tr>" & chr(13) & _
                "   <tr style='background:azure'>" & chr(13) & _
                "       <th class='MB ME MD' style='text-align:left;'><span class='R'>Usuário</span></th>" & chr(13) & _
                "       <th class='MB MD' style='text-align:left;'><span class='R'>Pedido(s)</span></th>" & chr(13) & _
                "   </tr>" & chr(13) & _
                "   </thead>" & chr(13) & _
                "   <tbody>" & chr(13)

        for i=LBound(vLocked) to UBound(vLocked)
            if Trim("" & vLocked(i).c1) <> "" then
                x = x & _
                        "   <tr>" & chr(13) & _
                        "       <td class='MB ME MD'><span class='C'>" & Trim("" & vLocked(i).c1) & "</span></td>" & chr(13) & _
                        "       <td class='MB MD'><span class='C'>" & Trim("" & vLocked(i).c2) & "</span></td>" & chr(13) & _
                        "   </tr>" & chr(13)
                end if
            next

        x = x & _
                "   </tbody>" & chr(13) & _
                "</table>" & chr(13) & _
                "<br />" & chr(13)
        
        if x <> "" then Response.Write x
        end if 'if qtdePedidoBloqueado > 0

	qtde_notas = qtdePedidosExibida
	
	if r.State <> 0 then r.Close
	set r=nothing
	
end sub

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



<% if False then 'APENAS P/ HABILITAR O INTELLISENSE DURANTE O DESENVOLVIMENTO!! %>
<script src="../Global/jquery.js" language="JavaScript" type="text/javascript"></script>
<% end if %>

<script src="<%=URL_FILE__JQUERY%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
window.status='Aguarde, executando a consulta ...';

function realca_cor_mouse_over(c) {
	c.style.backgroundColor = 'palegreen';
}

function realca_cor_mouse_out(c) {
	c.style.backgroundColor = '';
}

function fRELConcluir( id_pedido ){
	window.status = "Aguarde ...";
	fREL.pedido_selecionado.value=id_pedido;
	fREL.action = "Pedido.asp";
	fREL.submit();
}

function fRELVoltarTelaFiltro(f) {
    f.action = "RelControleImpostosFiltro.asp";
    f.submit();
}

function fRELGravaDados(f) {
	var i, intQtdeVerificadas, dtColeta, dtHoje;
	var s;

	intQtdeVerificadas = 0;
	for (i = 0; i < f.ckb_controle_impostos.length; i++) {
		if (f.ckb_controle_impostos[i].checked && !f.ckb_controle_impostos[i].disabled) intQtdeVerificadas++;
	}

	if (intQtdeVerificadas == 0) {
		alert('Nenhuma NFe foi selecionada!!');
		return;
	}

	if (!isDate(f.c_dt_coleta)) {
		alert('Data de coleta inválida!!');
		f.c_dt_coleta.focus();
		return;
	}

	window.status = "Aguarde ...";
	f.action = "RelControleImpostosGravaDados.asp";
	f.submit();
}

function fRELMarcarTodos(f) {
	var i;
	for (i = 0; i < f.ckb_controle_impostos.length; i++) {
		if (!f.ckb_controle_impostos[i].disabled) f.ckb_controle_impostos[i].checked = true;
	}
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
<link href="<%=URL_FILE__EPRINTER_CSS%>" rel="stylesheet" type="text/css" media="print">

<style type="text/css">
.TxtClip
{
	background:transparent;
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
<div class="MtAlerta" style="width:680px;font-weight:bold;" align="center"><p style='margin:5px 2px 5px 2px;'><%=alerta%></p></div>
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
<body onload="window.status='Concluído';" link=#000000 alink=#000000 vlink=#000000>

<center>

<form id="fREL" name="fREL" method="post">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<!-- Nº DO PEDIDO P/ CONSULTAR O PEDIDO AO CLICAR SOBRE O NÚMERO -->
<input type="hidden" name="pedido_selecionado" id="pedido_selecionado" value="">
<!-- FILTROS -->
<input type="hidden" name="c_transportadora" id="c_transportadora" value="<%=c_transportadora%>" />
<input type="hidden" name="c_dt_coleta" id="c_dt_coleta" value="<%=c_dt_coleta%>" />
<input type="hidden" name="c_dt_coleta_inicio" id="c_dt_coleta_inicio" value="<%=c_dt_coleta_inicio%>" />
<input type="hidden" name="c_dt_coleta_termino" id="c_dt_coleta_termino" value="<%=c_dt_coleta_termino%>" />
<input type="hidden" name="ckb_exibir_verificados" id="ckb_exibir_verificados" value="<%=ckb_exibir_verificados%>" />
<input type="hidden" name="c_nfe_emitente" id="c_nfe_emitente" value="<%=c_nfe_emitente%>" />
<input type="hidden" name="c_uf" id="c_uf" value="<%=c_uf%>" />
<!--  ASSEGURA CRIAÇÃO DE UM ARRAY DE CAMPOS, MESMO QUANDO HOUVER SOMENTE 1 LINHA!! -->
<input type="hidden" name="ckb_controle_impostos" id="ckb_controle_impostos" value="">

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="935" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Controle de Impostos</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>

<!-- FILTROS -->
<% 
	s_filtro = "<table width='935' cellpadding='0' cellspacing='0' style='border-bottom:1px solid black' border='0'>"
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>Tipo de Relatório:&nbsp;</span></td><td align='left' valign='top'>" & _
			   "<span class='N'>Controle de Impostos</span></td></tr>"

	s = c_transportadora
	if s = "" then 
		s = "todas"
	else
		if (s_nome_transportadora <> "") And (Ucase(s_nome_transportadora) <> Ucase(c_transportadora)) then s = s & "  (" & s_nome_transportadora & ")"
		end if
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
				"<span class='N'>Transportadora:&nbsp;</span></td><td align='left' valign='top'>" & _
				"<span class='N'>" & s & "</span></td></tr>"

    s = c_uf
	if s = "" then s = "todas"
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
				"<span class='N'>UF:&nbsp;</span></td><td align='left' valign='top'>" & _
				"<span class='N'>" & s & "</span></td></tr>"
	
	s = c_dt_coleta
	if s = "" then s = "N.I."
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
				"<span class='N'>Data Coleta:&nbsp;</span></td><td align='left' valign='top'>" & _
				"<span class='N'>" & s & "</span></td></tr>"
	
	s = c_dt_coleta_inicio
	if s <> "" then s = s + " a " + c_dt_coleta_termino
	if s = "" then s = "N.I."
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
				"<span class='N'>Período de Coleta:&nbsp;</span></td><td align='left' valign='top'>" & _
				"<span class='N'>" & s & "</span></td></tr>"
	
    s = ckb_exibir_verificados
    if s = "" then
        s = "Não"
    else
        s = "Sim"
        end if
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
				"<span class='N'>Exibir Já Verificados:&nbsp;</span></td><td align='left' valign='top'>" & _
				"<span class='N'>" & s & "</span></td></tr>"

    s = c_nfe_emitente
    if s = "" then
        s = "N.I."
    else
        s = obtem_apelido_empresa_NFe_emitente(s)
        end if
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
				"<span class='N'>CD:&nbsp;</span></td><td align='left' valign='top'>" & _
				"<span class='N'>" & s & "</span></td></tr>"

	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>Emissão:&nbsp;</span></td><td align='left' valign='top' width='99%'>" & _
			   "<span class='N'>" & formata_data_hora(Now) & "</span></td></tr>"

	s_filtro = s_filtro & "</table>"
	Response.Write s_filtro
%>

<!--  RELATÓRIO  -->
<br>
<% consulta_executa %>

<% if blnErroAcessoConcorrente then %>
    <br />
    <p class="ALERTA"><%=s_msg_erro_acesso_concorrente%></p>
<% end if %>

<% if (qtde_notas > 0) And (Not blnErroAcessoConcorrente) then %>
<br />
<table>
	<tr>
		<td align="right">
		<input name="bMarcarOK" id="bMarcarOK" type="button" class="Button" onclick="fRELMarcarTodos(fREL)" value="Marcar todas as Notas Fiscais como OK" title="assinala todas as notas" style="margin-left:6px;margin-bottom:10px">
		</td>
	</tr>
</table>
<% end if %>

<!-- ************   SEPARADOR   ************ -->
<table width="935" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc" align="left">&nbsp;</td></tr>
</table>
<br>


<table class="notPrint" width="935" cellspacing="0">
<tr>
	<% if (qtde_notas > 0) And (Not blnErroAcessoConcorrente) then %>
	<td align="left">
		<a name="bVOLTA" id="bVOLTA" href="<%=strUrlBotaoVoltar%>" title="volta para a página anterior"><img src="../botao/voltar.gif" width="176" height="55" border="0"></a>
	</td>
	<td align="left">&nbsp;</td>
	<td align="right">
		<a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fRELGravaDados(fREL)" title="grava os dados"><img src="../botao/confirmar.gif" width="176" height="55" border="0"></a>
	</td>
	<% else %>
	<td align="center">
		<a name="bVOLTA" id="bVOLTA" href="<%=strUrlBotaoVoltar%>" title="volta para a página anterior"><img src="../botao/voltar.gif" width="176" height="55" border="0"></a>
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
    'SE HOUVE CONFLITO DE ACESSO CONCORRENTE, CANCELA O RELATÓRIO E REMOVE OS LOCKS
    if blnErroAcessoConcorrente then
    '   LIMPA EVENTUAIS LOCKS REMANESCENTES
        s = "UPDATE t_CTRL_RELATORIO_USUARIO_X_PEDIDO SET" & _
                " locked = 0," & _
                " cod_motivo_lock_released = " & CTRL_RELATORIO_CodMotivoLockReleased_ConflitoAcessoConcorrente & "," & _
                " dt_hr_lock_released = getdate()" & _
            " WHERE" & _
                " (usuario = '" & QuotedStr(usuario) & "')" & _
                " AND (id_relatorio = " & ID_CTRL_RELATORIO_RelControleImpostos & ")" & _
                " AND (locked = 1)"
        cn.Execute(s)
        end if
%>

<%
	if tLock.State <> 0 then tLock.Close
	set tLock = nothing

'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>
