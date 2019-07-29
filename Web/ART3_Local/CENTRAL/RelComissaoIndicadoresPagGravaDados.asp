<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp" -->

<%
'     =================================================================
'	  RelComissaoIndicadoresPagGravaDados.asp
'     =================================================================
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
'	REVISADO P/ IE10

	On Error GoTo 0
	Err.Clear
	
	Server.ScriptTimeout = MAX_SERVER_SCRIPT_TIMEOUT_EM_SEG
	
	const VENDA_NORMAL = "VENDA_NORMAL"
	const DEVOLUCAO = "DEVOLUCAO"
	const PERDA = "PERDA"

'	COMO O TRATAMENTO DO RELATÓRIO PODE SER DEMORADO, CASO A SESSÃO EXPIRE E O TRATAMENTO
'	DE SESSÃO EXPIRADA NÃO CONSIGA RESTAURÁ-LA, OBTÉM A IDENTIFICAÇÃO DO USUÁRIO A PARTIR DE
'	UM CAMPO HIDDEN CRIADO NA PÁGINA CHAMADORA EXCLUSIVAMENTE P/ ISSO.
	dim s, s2, usuario, msg_erro, s_log
	usuario = Trim(Session("usuario_atual"))
	if (usuario = "") then usuario = Trim(Request("c_usuario_sessao"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO)

	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	if Not operacao_permitida(OP_CEN_FLAG_COMISSAO_PAGA, s_lista_operacoes_permitidas) then 
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if
	
	dim s_id, s_pedido, qtde_total_reg_update, i, s_devolucao, s_perda, indicador_a, mes, ano, s_operacao, max_caracteres_favorecido
    dim v_pedido, v_devolucao, v_perda, v_operacao
    dim mes_competencia, ano_competencia, favorecido, intNsuNovoFluxoCaixa, mensagem,vendedor,vendedor_a, descricao_fluxo_caixa, dt_mes_competencia
    dim empresa_fluxo_caixa, conta_corrente_fluxo_caixa, dt_competencia_fluxo_caixa, conta_comissao_fluxo_caixa, conta_RA_fluxo_caixa, conta_comissao_grupo, conta_RA_grupo, qtde_reg_fluxo_caixa
    dim vl_total_comissao, vl_total_RA, id_n3, rb_visao
    dim o
    set o = createobject( "ComPlusCalcCedulas.ComPlusCalcCedulas" )

    dt_competencia_fluxo_caixa = Trim(Request.Form("dt_competencia_fluxo_caixa"))
    empresa_fluxo_caixa = Trim(Request.Form("empresa_fluxo_caixa"))
    conta_corrente_fluxo_caixa = Trim(Request.Form("conta_corrente_fluxo_caixa"))
    conta_comissao_fluxo_caixa = Trim(Request.Form("conta_comissao_fluxo_caixa"))
    conta_RA_fluxo_caixa = Trim(Request.Form("conta_RA_fluxo_caixa"))
    rb_visao = Trim(Request.Form("rb_visao"))

    vl_total_comissao = 0
    vl_total_RA = 0
    
    dt_competencia_fluxo_caixa = StrToDate(dt_competencia_fluxo_caixa)

    s_id = Trim(Request.Form("c_id"))
    qtde_reg_fluxo_caixa=0

    if COD_FC_AMBIENTE = "L" then
        max_caracteres_favorecido = 27
    else
        max_caracteres_favorecido = 28
	end if

	dim alerta
	alerta=""

'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, cn2, r, rs, rs2
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	if Not bdd_conecta_RPIFC(cn2) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO) 

		If Not cria_recordset_pessimista(rs, msg_erro) then 
			cn.RollbackTrans
			Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
		end if
    	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
        If Not cria_recordset_otimista(rs2, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
        If Not cria_recordset_pessimista(r, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)


' ___________________________________________
' FIN GERA NSU FLUXO CAIXA DIS
'
function fin_gera_nsu_fluxo_caixa_dis(byval idNsu, byref nsu, byref msg_erro)
dim t, strSql, intRetorno, intRecordsAffected
dim intQtdeTentativas, intNsuUltimo, intNsuNovo, blnSucesso
	fin_gera_nsu_fluxo_caixa_dis=False
	msg_erro=""
	nsu=0
	strSql = "SELECT" & _
				" Count(*) AS qtde" & _
			" FROM t_FIN_CONTROLE" & _
			" WHERE" & _
				" (id='" & idNsu & "')"
	set t=cn2.Execute(strSql)
	if Not t.Eof then intRetorno=Clng(t("qtde")) else intRetorno=Clng(0)

'	NÃO ESTÁ CADASTRADO, ENTÃO CADASTRA AGORA
	if intRetorno=0 then
		strSql = "INSERT INTO t_FIN_CONTROLE (" & _
					"id, " & _
					"nsu, " & _
					"dt_hr_ult_atualizacao" & _
				") VALUES (" & _
					"'" & idNsu & "'," & _
					"0," & _
					"getdate()" & _
				")"
		cn2.Execute strSql, intRecordsAffected
		if intRecordsAffected <> 1 then
			msg_erro = "Falha ao criar o registro para geração de NSU (" & idNsu & ")!!"
			exit function
			end if
		end if

'	LAÇO DE TENTATIVAS PARA GERAR O NSU (DEVIDO A ACESSO CONCORRENTE)
	intQtdeTentativas=0
	do 
		intQtdeTentativas = intQtdeTentativas + 1
		
	'	OBTÉM O ÚLTIMO NSU USADO
		strSql = "SELECT" & _
					" nsu" & _
				" FROM t_FIN_CONTROLE" & _
				" WHERE" & _
					" id = '" & idNsu & "'"
		set t=cn2.Execute(strSql)
		if t.Eof then
			strMsgErro = "Falha ao localizar o registro para geração de NSU (" & idNsu & ")!!"
			Exit Function
		else
			intNsuUltimo = Clng(t("nsu"))
			end if

	'	INCREMENTA 1
		intNsuNovo = intNsuUltimo + 1
		
	'	TENTA ATUALIZAR O BANCO DE DADOS
		strSql = "UPDATE t_FIN_CONTROLE SET" & _
					" nsu = " & CStr(intNsuNovo) & "," & _
					" dt_hr_ult_atualizacao = getdate()" & _
				" WHERE" & _
					" (id = '" & idNsu & "')" & _
					" AND (nsu = " & CStr(intNsuUltimo) & ")"
		cn2.Execute strSql, intRecordsAffected
		If intRecordsAffected = 1 Then
			blnSucesso = True
			nsu = intNsuNovo
			end if
		
		Loop While (Not blnSucesso) And (intQtdeTentativas < 10)

	If Not blnSucesso Then
		strMsgErro = "Falha ao tentar gerar o NSU!!"
		Exit Function
		End If
	
	fin_gera_nsu_fluxo_caixa_dis = True

end function

    qtde_total_reg_update = 0
    s_pedido = ""
    s_devolucao = ""
    s_perda=""
    s_operacao=""

    rs.Open "SELECT id_plano_contas_grupo FROM t_FIN_PLANO_CONTAS_CONTA WHERE (id='" & conta_comissao_fluxo_caixa & "')", cn
        if Not rs.Eof then 
            conta_comissao_grupo = rs("id_plano_contas_grupo")
        else
            alerta=texto_add_br(alerta)
				alerta=alerta & "O plano de contas (comissão) " & conta_comissao_fluxo_caixa & " não está cadastrado."
        end if
    if rs.State <> 0 then rs.Close
    rs.Open "SELECT id_plano_contas_grupo FROM t_FIN_PLANO_CONTAS_CONTA WHERE (id='" & conta_RA_fluxo_caixa & "')", cn
        if Not rs.Eof then 
            conta_RA_grupo = rs("id_plano_contas_grupo")
        else
            alerta=texto_add_br(alerta)
				alerta=alerta & "O plano de contas (RA) " & conta_RA_fluxo_caixa & " não está cadastrado."
        end if
    if rs.State <> 0 then rs.Close

    s="SELECT * FROM t_COMISSAO_INDICADOR_N2 INNER JOIN t_COMISSAO_INDICADOR_N1 ON (t_COMISSAO_INDICADOR_N2.id_comissao_indicador_n1=t_COMISSAO_INDICADOR_N1.id) WHERE (t_COMISSAO_INDICADOR_N1.id='" & s_id & "')"
    rs.Open s, cn
    do while not rs.Eof
            
        rs("proc_automatico_qtde_tentativas")=rs("proc_automatico_qtde_tentativas")+1
        rs.Update
        rs.MoveNext
    loop
    if rs.State <> 0 then rs.Close

    if alerta = "" then
        s = "SELECT * FROM t_COMISSAO_INDICADOR_N4" & _
                        " INNER JOIN t_COMISSAO_INDICADOR_N3 ON (t_COMISSAO_INDICADOR_N3.id = t_COMISSAO_INDICADOR_N4.id_comissao_indicador_n3)" & _
                        " INNER JOIN t_COMISSAO_INDICADOR_N2 ON (t_COMISSAO_INDICADOR_N2.id = t_COMISSAO_INDICADOR_N3.id_comissao_indicador_n2)" & _
                        " INNER JOIN t_COMISSAO_INDICADOR_N1 ON (t_COMISSAO_INDICADOR_N1.id = t_COMISSAO_INDICADOR_N2.id_comissao_indicador_n1)" & _
                        " WHERE (t_COMISSAO_INDICADOR_N1.id = '" & s_id & "') AND (t_COMISSAO_INDICADOR_N3.st_tratamento_manual=0)"

        set r = cn.Execute(s)
        if r("proc_automatico_status")=1 then
            alerta=texto_add_br(alerta)
				    alerta=alerta & "O Relatório já foi processado por " & r("proc_automatico_usuario") & " em " & r("proc_automatico_data_hora") & "."
        end if
    end if

    if alerta = "" then
        cn.BeginTrans
        mes_competencia=r("competencia_mes")
        ano_competencia=r("competencia_ano")
        
        do while not r.Eof
            if s_pedido <> "" then s_pedido = s_pedido & ";"
            s_pedido = s_pedido & r("pedido")
            if s_operacao <> "" then s_operacao = s_operacao & ";"
            s_operacao = s_operacao & r("tabela_origem")


            if vendedor = "" then
                vendedor = vendedor & r("vendedor")
            else
                if r("vendedor") <> vendedor_a then 
                   vendedor = vendedor & "," & r("vendedor")
                end if
            end if
           vendedor_a = Trim("" & r("vendedor"))
        r.MoveNext
        loop

        v_pedido = split(s_pedido, ";")
        v_operacao = split(s_operacao, ";")
    
        if r.State <> 0 then r.Close

        for i=Lbound(v_pedido) to Ubound(v_pedido)
            if v_operacao(i) = "VEN" then
                rs.Open "SELECT * FROM t_PEDIDO WHERE (pedido='" & v_pedido(i) & "')", cn
                if rs.Eof then
				        alerta=texto_add_br(alerta)
				        alerta=alerta & "Pedido " & v_pedido(i) & " não foi encontrado."
                        if rs.State <> 0 then rs.Close
                else
                    rs("comissao_paga") = 1
                    rs("comissao_paga_ult_op") = "S"
                    rs("comissao_paga_data")=Date
                    rs("comissao_paga_usuario")=usuario
                    qtde_total_reg_update = qtde_total_reg_update + 1
                    rs.Update
                    if Err <> 0 then
					        alerta=texto_add_br(alerta)
					        alerta=alerta & Cstr(Err) & ": " & Err.Description
                            cn.RollbackTrans
			        end if
                    if rs.State <> 0 then rs.Close
                end if
            elseif v_operacao(i) = "DEV" then
                ' baixa comissões das devolução
                rs.Open "SELECT * FROM t_PEDIDO_ITEM_DEVOLVIDO WHERE (pedido = '" & v_pedido(i) & "') AND (comissao_descontada=0)", cn
                do while Not rs.Eof
                    rs("comissao_descontada")=1
		            rs("comissao_descontada_ult_op")="S"
		            rs("comissao_descontada_data")=Date
		            rs("comissao_descontada_usuario")=usuario
                    rs.Update
		            qtde_total_reg_update = qtde_total_reg_update + 1
                    rs.MoveNext
                Loop
                if rs.State <> 0 then rs.Close
            elseif v_operacao(i) = "PER" then
                ' baixa comissões das perdas
                rs.Open "SELECT * FROM t_PEDIDO_PERDA WHERE (pedido = '" & v_pedido(i) & "') AND (comissao_descontada=0)", cn
                do while Not rs.Eof            
                    rs("comissao_descontada")=1
		            rs("comissao_descontada_ult_op")="S"
		            rs("comissao_descontada_data")=Date
		            rs("comissao_descontada_usuario")=usuario
                    rs.Update
		            qtde_total_reg_update = qtde_total_reg_update + 1
                    rs.MoveNext
                Loop
                if rs.State <> 0 then rs.Close
            end if
        next

        s="SELECT * FROM t_COMISSAO_INDICADOR_N2 INNER JOIN t_COMISSAO_INDICADOR_N1 ON (t_COMISSAO_INDICADOR_N2.id_comissao_indicador_n1=t_COMISSAO_INDICADOR_N1.id) WHERE (t_COMISSAO_INDICADOR_N1.id='" & s_id & "')"

        rs.Open s, cn
        do while not rs.Eof
            rs("proc_automatico_status")=1
            rs("proc_automatico_data")=Date
            rs("proc_automatico_data_hora")=Now()
            rs("proc_automatico_usuario")=usuario
            rs("dt_competencia_fluxo_caixa")=StrToDate(dt_competencia_fluxo_caixa)
            rs.Update
            rs.MoveNext
        loop
        if rs.State <> 0 then rs.Close
        
    ' grava fluxo de caixa
    s = "SELECT t_COMISSAO_INDICADOR_N3.id AS id_n3, * FROM t_COMISSAO_INDICADOR_N3" & _
            " INNER JOIN t_COMISSAO_INDICADOR_N2 ON (t_COMISSAO_INDICADOR_N2.id = t_COMISSAO_INDICADOR_N3.id_comissao_indicador_n2)" & _
            " INNER JOIN t_COMISSAO_INDICADOR_N1 ON (t_COMISSAO_INDICADOR_N1.id = t_COMISSAO_INDICADOR_N2.id_comissao_indicador_n1)" & _
            " WHERE (t_COMISSAO_INDICADOR_N1.id = '" & s_id & "') AND (t_COMISSAO_INDICADOR_N3.st_tratamento_manual=0)"

    dt_mes_competencia = StrToDate("01/" + CStr(mes_competencia) + "/" + CStr(ano_competencia))

    cn2.Execute("begin tran")

    rs.Open s, cn
    do while not rs.Eof
        id_n3 = rs("id_n3")       
        s =    "SELECT * FROM t_FIN_FLUXO_CAIXA WHERE (id='-1')"

        rs2.Open s, cn2

        if rs2.Eof then
            vl_total_comissao = rs("vl_total_comissao_arredondado")
            vl_total_RA = rs("vl_total_RA_arredondado")

            ' verifica se o RA é negativo, se sim, subtrair esse valor da comissão
            if rs("vl_total_RA")<0 then vl_total_comissao = vl_total_comissao + rs("vl_total_RA")            
             if rs("meio_pagto") = "DIN" then
                vl_total_comissao = o.digitoFinal(vl_total_comissao)
            else
                vl_total_comissao = floor(vl_total_comissao)
            end if

            ' verifica se a comissão é negativa, se sim, subtrair esse valor do RA
            if rs("vl_total_comissao")<0 then vl_total_RA = vl_total_RA + rs("vl_total_comissao")
            if rs("meio_pagto") = "DIN" then
                vl_total_RA = o.digitoFinal(vl_total_RA)
            else
                vl_total_RA = floor(vl_total_RA)
            end if
            if rs("vl_total_comissao_arredondado")>0 then
                if Not fin_gera_nsu_fluxo_caixa_dis(T_FIN_FLUXO_CAIXA, intNsuNovoFluxoCaixa, msg_erro) then 
			            alerta = "FALHA AO GERAR NSU PARA O NOVO REGISTRO (" & msg_erro & ")"
		            else
			            if intNsuNovoFluxoCaixa <= 0 then
				            alerta = "NSU GERADO É INVÁLIDO (" & intNsuNovoFluxoCaixa & ")"
				            end if
			            end if                
                favorecido=rs("favorecido")
                mes_competencia=rs("competencia_mes")
                rs2.AddNew
                rs2("id")=intNsuNovoFluxoCaixa
                rs2("id_conta_corrente")=conta_corrente_fluxo_caixa
                rs2("id_plano_contas_empresa")=empresa_fluxo_caixa
                rs2("natureza")="D"
                rs2("st_sem_efeito")=0
                rs2("id_plano_contas_grupo")=conta_comissao_grupo
                rs2("id_plano_contas_conta")=conta_comissao_fluxo_caixa
                rs2("valor")=vl_total_comissao
                rs2("dt_competencia")=StrToDate(dt_competencia_fluxo_caixa)
                rs2("tipo_cadastro")="S"
                rs2("editado_manual")="N"
                rs2("dt_cadastro")=Date
                rs2("dt_hr_cadastro")=Now()
                rs2("usuario_cadastro")=usuario
                rs2("dt_ult_atualizacao")=Date
                rs2("dt_hr_ult_atualizacao")=Now()
                rs2("usuario_ult_atualizacao")=usuario
                rs2("ctrl_pagto_id_parcela")=id_n3
                rs2("ctrl_pagto_modulo")=11
                rs2("ctrl_pagto_status")=1
                rs2("ctrl_pagto_id_ambiente_origem")=ID_AMBIENTE
                rs2("dt_mes_competencia")= dt_mes_competencia
                if len(favorecido)>max_caracteres_favorecido then 
                    favorecido = mid(favorecido,1,max_caracteres_favorecido)
                    favorecido=favorecido & " "
                end if
                if len(CStr(mes_competencia))<2 then mes_competencia = "0" & mes_competencia
                if rs("meio_pagto") = "DIN" then
                    descricao_fluxo_caixa = "R"
                elseif rs("meio_pagto") = "CHQ" then
                    descricao_fluxo_caixa = "C"
                elseif rs("meio_pagto") = "DEP" Or rs("meio_pagto") = "DEP1" then
                    descricao_fluxo_caixa = "D"
                end if
                descricao_fluxo_caixa = descricao_fluxo_caixa & COD_FC_AMBIENTE
                descricao_fluxo_caixa = descricao_fluxo_caixa & "-" & favorecido & " REF " & mes_competencia
                rs2("descricao")=descricao_fluxo_caixa
                rs2.Update
                qtde_reg_fluxo_caixa = qtde_reg_fluxo_caixa + 1
            end if

            if rs("vl_total_RA_arredondado")>0 then
                 if Not fin_gera_nsu_fluxo_caixa_dis(T_FIN_FLUXO_CAIXA, intNsuNovoFluxoCaixa, msg_erro) then 
			            alerta = "FALHA AO GERAR NSU PARA O NOVO REGISTRO (" & msg_erro & ")"
		            else
			            if intNsuNovoFluxoCaixa <= 0 then
				            alerta = "NSU GERADO É INVÁLIDO (" & intNsuNovoFluxoCaixa & ")"
				            end if
			            end if
                favorecido=rs("favorecido")
                mes_competencia=rs("competencia_mes")
                rs2.AddNew
                rs2("id")=intNsuNovoFluxoCaixa
                rs2("id_conta_corrente")=conta_corrente_fluxo_caixa
                rs2("id_plano_contas_empresa")=empresa_fluxo_caixa
                rs2("natureza")="D"
                rs2("st_sem_efeito")=0
                rs2("id_plano_contas_grupo")=conta_RA_grupo
                rs2("id_plano_contas_conta")=conta_RA_fluxo_caixa
                rs2("valor")=vl_total_RA
                rs2("dt_competencia")=StrToDate(dt_competencia_fluxo_caixa)
                rs2("tipo_cadastro")="S"
                rs2("editado_manual")="N"
                rs2("dt_cadastro")=Date
                rs2("dt_hr_cadastro")=Now()
                rs2("usuario_cadastro")=usuario
                rs2("dt_ult_atualizacao")=Date
                rs2("dt_hr_ult_atualizacao")=Now()
                rs2("usuario_ult_atualizacao")=usuario
                rs2("ctrl_pagto_id_parcela")=id_n3
                rs2("ctrl_pagto_modulo")=11
                rs2("ctrl_pagto_status")=1
                rs2("ctrl_pagto_id_ambiente_origem")=ID_AMBIENTE
                rs2("dt_mes_competencia")=dt_mes_competencia
                if len(favorecido)>max_caracteres_favorecido then
                    favorecido = mid(favorecido,1,max_caracteres_favorecido)
                    favorecido=favorecido & " "
                end if
                if len(CStr(mes_competencia))<2 then mes_competencia = "0" & mes_competencia
                if rs("meio_pagto") = "DIN" then
                    descricao_fluxo_caixa = "R"
                elseif rs("meio_pagto") = "CHQ" then
                    descricao_fluxo_caixa = "C"
                elseif rs("meio_pagto") = "DEP" Or rs("meio_pagto") = "DEP1" then
                    descricao_fluxo_caixa = "D"
                end if
                descricao_fluxo_caixa = descricao_fluxo_caixa & COD_FC_AMBIENTE
                descricao_fluxo_caixa = descricao_fluxo_caixa & "-" & favorecido & " REF " & mes_competencia
                rs2("descricao")=descricao_fluxo_caixa
                rs2.Update
                
                qtde_reg_fluxo_caixa = qtde_reg_fluxo_caixa + 1
            end if

            if Err <> 0 then
					        alerta=texto_add_br(alerta)
					        alerta=alerta & Cstr(Err) & ": " & Err.Description
                            cn.RollbackTrans
                            cn2.Execute("rollback tran")
			        end if
        end if
        
        if rs2.State <> 0 then rs2.Close
    rs.MoveNext
    loop
    
    mensagem = "VENDEDOR(ES) ESCOLHIDO(S): " & vendedor & "; "&"Mês de competência: " & mes_competencia & "/"& ano_competencia & "; "
    mensagem = mensagem & "DATA FLUXO DE CAIXA: " & dt_competencia_fluxo_caixa & ";" & " CONTA CORRENTE: " & conta_corrente_fluxo_caixa & "; " & " EMPRESA: " &  empresa_fluxo_caixa & ";" & " COMISSÃO: " & conta_comissao_fluxo_caixa & ";" & " RA: " & conta_RA_fluxo_caixa & ";"
    mensagem = mensagem & "PEDIDO(S) ATUALIZADO(S): " & qtde_total_reg_update & ";" & " LANÇAMENTO(S) REALIZADOS NO FLUXO DE CAIXA: " & qtde_reg_fluxo_caixa
    grava_log usuario,"","","",OP_LOG_REL_COMISSAO_INDICADORES_GRAVADADOS  , mensagem

    if alerta="" then
        cn.CommitTrans   
        cn2.Execute("commit tran") 
    else
        cn.RollbackTrans
        cn2.Execute("rollback tran")

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

<script language="JavaScript" type="text/javascript">
function fRetornar(f) {
	f.action = "RelComissaoIndicadoresPag.asp";
	dVOLTAR.style.visibility = "hidden";
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


<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black;">

<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Relatório de Pedidos Indicadores</span></td>
</tr>
</table>
<br>
<br>

<!-- ************   MENSAGEM  ************ -->
<div class='MtAviso' style="width:649px;font-weight:bold;border:1pt solid black;" align="center">
	<% if qtde_total_reg_update = 0 then %>
	<span style='margin:5px 2px 5px 2px;'>Nenhuma alteração foi realizada para gravar</span>
	<% else %>
	<span style='margin:5px 2px 5px 2px;'>Dados gravados com sucesso:<br /> <%=Cstr(qtde_total_reg_update)%> pedidos(s) atualizado(s)</span><br />
    <span style="margin:5px 2px 5px 2px;"><%=Cstr(qtde_reg_fluxo_caixa)%> lançamento(s) registrado(s) no fluxo de caixa</span>
	<% end if %>
</div>
<br>


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
    <% if rb_visao = "" then %>
<table class="notPrint" width="649" cellspacing="0">
<tr>
	<td align="center"><div name="dVOLTAR" id="dVOLTAR"><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.go(-2);" title="Retornar para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a></div>
	</td>
</tr>
</table>
    <% else %>

<input type="hidden" name="rb_visao" id="rb_visao" value="<%=rb_visao%>" />

<table class="notPrint" width="649" cellspacing="0">
<tr>
	<td align="center"><div name="dVOLTAR" id="dVOLTAR"><a name="bVOLTAR" id="bVOLTAR" href="javascript:fRetornar(f);" title="Retornar para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a></div>
	</td>
</tr>
</table>
    <% end if %>
</form>

</center>
</body>

<% end if %>

</html>


<%

'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing

    cn2.Close
    set cn2 = nothing

%>