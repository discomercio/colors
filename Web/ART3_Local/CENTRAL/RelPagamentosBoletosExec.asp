<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=False %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ===========================================================
'	  RelPagamentosBoletosExec.asp
'     ===========================================================
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
	dim cn, rs, msg_erro
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	if Not operacao_permitida(OP_CEN_REL_PERFIL_PAGAMENTO_BOLETOS, s_lista_operacoes_permitidas) then 
		cn.Close
    	Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	dim alerta
	dim s, s_aux, s_filtro
	dim c_dt_inicio, c_dt_termino
	dim c_loja, lista_loja, s_filtro_loja, v_loja, v, i
	dim rb_visao, blnVisaoSintetica
    
	alerta = ""

	c_dt_inicio = Trim(Request("c_dt_inicio"))
	c_dt_termino = Trim(Request("c_dt_termino"))


	
	if alerta = "" then
		if c_dt_inicio <> "" then
			if Not IsDate(StrToDate(c_dt_inicio)) then
				alerta = "DATA DE INÍCIO DO PERÍODO É INVÁLIDA."
				end if
			end if
		end if
	
	if alerta = "" then
		if c_dt_termino <> "" then
			if Not IsDate(StrToDate(c_dt_termino)) then
				alerta = "DATA DE TÉRMINO DO PERÍODO É INVÁLIDA."
				end if
			end if
		end if

' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________

function monta_link_pedido(byval id_pedido)
dim strLink
	monta_link_pedido = ""
	id_pedido = Trim("" & id_pedido)
	if id_pedido = "" then exit function
	strLink = "<a href='javascript:fPEDConsulta(" & _
				chr(34) & id_pedido & chr(34) & _
				")' title='clique para consultar o pedido " & id_pedido & "' style='color:black'>" & _
				id_pedido & "</a>"
	monta_link_pedido=strLink
end function

' _____________________________________
' CONSULTA EXECUTA
'
sub consulta_executa
dim r
dim s, s_aux, s_sql, x, cab_table, cab, n_reg, n_reg_total
dim qtde_sub_total, vl_sub_total, min_atraso_sub_total, media_atraso_sub_total, max_atraso_sub_total
dim qtde_total, vl_total, mes_a, mes
dim qtde, valor, min_atraso, media_atraso, max_atraso
dim s_where, s_where_venda, s_where_devolucao, s_where_perdas, s_where_loja
dim s_where_comissao_paga, s_where_comissao_descontada, s_where_st_pagto, s_lista_pedidos_atrasados, n_pedidos_atrasados
		
'	BOLETOS COM VENCIMENTOS NO PERÍODO (IGNORA BOLETOS DE PEDIDOS CANCELADOS)

	s_sql = "SELECT" & _
	           " tC.tipo," & _
	           " Count(*) AS qtde," & _
	           " Sum(valor) AS vl_previsto" & _
           " FROM t_FIN_BOLETO_ITEM tFBI" & _
	           " INNER JOIN t_FIN_BOLETO tFB ON (tFBI.id_boleto=tFB.id)" & _
	           " LEFT JOIN t_CLIENTE tC ON (tFB.id_cliente = tC.id)" & _
           " WHERE" & _ 
	           " (dt_vencto BETWEEN " & bd_formata_data(StrToDate(c_dt_inicio)) & " AND " & bd_formata_data(StrToDate(c_dt_termino)) & ")" & _
	           " AND (dt_entrada_confirmada IS NOT NULL)" & _
	           " AND (tFBI.id NOT IN (SELECT ctrl_pagto_id_parcela FROM t_FIN_FLUXO_CAIXA WHERE (ctrl_pagto_modulo = 1) AND (st_sem_efeito = 1) AND (st_confirmacao_pendente = 1)))" & _
           " GROUP BY" & _
	           " tC.tipo"



  ' cabeçalho
	cab_table = "<table cellspacing='0' cellpadding='0' style='border:0;width:520px'>" & chr(13) & _
                "<tr>" & chr(13) & _
                "       <td style='border-bottom: 1px solid #000'><span class='N'>BOLETOS COM VENCIMENTO NO PERÍODO</span></td>" & chr(13) & _
                "</tr>" & chr(13) & _
                "   <td align='center'><br>" & chr(13)

	cab = "<table cellspacing='0'>" & chr(13) & _
          "	<tr style='background:azure' nowrap>" & chr(13) & _
		  "		<td class='Col1' style='background-color:#fff;border:0;'>&nbsp;</td>" & chr(13) & _
		  "		<td class='MTD ME Col2' align='center' valign='bottom' nowrap><span class='Rc'>Qtde</span></td>" & chr(13) & _
		  "		<td class='MTD Col3' align='right' valign='bottom' nowrap><span class='R'>Valor previsto</span></td>" & chr(13) & _
		  "	</tr>" & chr(13)
	
	x = ""
	n_reg = 0
	n_reg_total = 0
	qtde_sub_total = 0
    vl_sub_total = 0

    x = cab_table & cab
	rs.Open s_sql, cn
	do while Not rs.Eof
       
       x = x & "<tr>" & chr(13) & _
                    "   <td class='ME MC MD Col1' align='center' valign='middle'><span class='Cn'>" & rs("tipo") & "</span></td>" & chr(13) & _
                    "   <td class='MC MD Col2' align='center' valign='middle'><span class='Cn'>" & FormatNumber(rs("qtde"), 0) & "</span></td>" & chr(13) & _
                    "   <td class='MC MD Col3' align='right' valign='middle'><span class='Cn'>" & formata_moeda(rs("vl_previsto")) & "</span></td>" & chr(13) & _
               "</tr>" & chr(13)
        qtde_sub_total = qtde_sub_total + rs("qtde")
        vl_sub_total = vl_sub_total + rs("vl_previsto")        
		rs.MoveNext
    loop
    if rs.State <> 0 then rs.Close

    x = x & "<tr>" & chr(13) & _
                    "   <td class='ME MD MC MB Col1' align='center' valign='middle'><span class='C'>TOTAL</span></td>" & chr(13) & _
                    "   <td class='MC MD MB Col2' align='center' valign='middle'><span class='C'>" & FormatNumber(qtde_sub_total, 0) & "</span></td>" & chr(13) & _
                    "   <td class='MC MD MB Col3' align='right' valign='middle'><span class='C'>" & formata_moeda(vl_sub_total) & "</span></td>" & chr(13) & _                    
	        "</tr>" & chr(13) & _
        "</table><br><br><br>" & chr(13)

    x = x & "   </td>" & chr(13) & _
            "</tr>" & chr(13)
        
	Response.write x


'	BOLETOS COM VENCIMENTO NO PERÍODO E PAGOS EM DIA

	s_sql = "SELECT" & _
	           " tipo," & _
	           " Count(*) AS qtde," & _
	           " Sum(valor) AS vl_pago_em_dia" & _
           " FROM (" & _
	           " SELECT" & _
		           " tC.tipo," & _
		           " dt_competencia," & _
		           " tFFC.valor," & _
		           " dt_vencto," & _
		           " descricao" & _
	           " FROM t_FIN_FLUXO_CAIXA tFFC" & _
			           " LEFT JOIN t_FIN_BOLETO_ITEM tFBI ON (tFBI.id = tFFC.ctrl_pagto_id_parcela) AND (tFFC.ctrl_pagto_modulo = 1)" & _
			           " LEFT JOIN t_FIN_BOLETO tFB ON (tFBI.id_boleto = tFB.id)" & _
			           " LEFT JOIN t_CLIENTE tC ON (tFB.id_cliente = tC.id)" & _
	           " WHERE" & _
		           " (ctrl_pagto_modulo = 1)" & _
		           " AND (st_sem_efeito = 0)" & _
		           " AND (st_confirmacao_pendente = 0)" & _ 
                   " AND (ctrl_pagto_id_parcela IN (SELECT id FROM t_FIN_BOLETO_ITEM WHERE (dt_vencto BETWEEN " & bd_formata_data(StrToDate(c_dt_inicio)) & " AND " & bd_formata_data(StrToDate(c_dt_termino)) & ") AND (dt_entrada_confirmada IS NOT NULL)))" & _
            ") t" & _
           " WHERE" & _
	           " dt_competencia <= dt_vencto" & _
           " GROUP BY" & _
	           " tipo"


  ' cabeçalho
	cab_table = "<tr>" & chr(13) & _
                "       <td style='border-bottom: 1px solid #000'><span class='N'>BOLETOS COM VENCIMENTO NO PERÍODO E PAGOS EM DIA</span></td>" & chr(13) & _
                "</tr>" & chr(13) & _
                "   <td align='center'><br>" & chr(13)

	cab = "<table cellspacing='0'>" & chr(13) & _
          "	<tr style='background:azure' nowrap>" & chr(13) & _
		  "		<td class='Col1' style='background-color:#fff;border:0;'>&nbsp;</td>" & chr(13) & _
		  "		<td class='MTD ME Col2' align='center' valign='bottom' nowrap><span class='Rc'>Qtde</span></td>" & chr(13) & _
		  "		<td class='MTD Col3' align='right' valign='bottom' nowrap><span class='R'>Valor pago em dia</span></td>" & chr(13) & _
		  "	</tr>" & chr(13)
	
	x = ""
	qtde_sub_total = 0
    vl_sub_total = 0

    x = cab_table & cab
	rs.Open s_sql, cn
	do while Not rs.Eof
       
       x = x & "<tr>" & chr(13) & _
                    "   <td class='ME MC MD Col1' align='center' valign='middle'><span class='Cn'>" & rs("tipo") & "</span></td>" & chr(13) & _
                    "   <td class='MC MD Col2' align='center' valign='middle'><span class='Cn'>" & FormatNumber(rs("qtde"), 0) & "</span></td>" & chr(13) & _
                    "   <td class='MC MD Col3' align='right' valign='middle'><span class='Cn'>" & formata_moeda(rs("vl_pago_em_dia")) & "</span></td>" & chr(13) & _
               "</tr>" & chr(13)
        qtde_sub_total = qtde_sub_total + rs("qtde")
        vl_sub_total = vl_sub_total + rs("vl_pago_em_dia")        
		rs.MoveNext
    loop
    if rs.State <> 0 then rs.Close

    x = x & "<tr>" & chr(13) & _
                    "   <td class='ME MD MC MB Col1' align='center' valign='middle'><span class='C'>TOTAL</span></td>" & chr(13) & _
                    "   <td class='MC MD MB Col2' align='center' valign='middle'><span class='C'>" & FormatNumber(qtde_sub_total, 0) & "</span></td>" & chr(13) & _
                    "   <td class='MC MD MB Col3' align='right' valign='middle'><span class='C'>" & formata_moeda(vl_sub_total) & "</span></td>" & chr(13) & _                    
	        "</tr>" & chr(13) & _
        "</table><br><br><br>" & chr(13)

    x = x & "   </td>" & chr(13) & _
            "</tr>" & chr(13)
        
	Response.write x




'	BOLETOS COM VENCIMENTO NO PERÍODO E PAGOS COM ATRASO
    min_atraso_sub_total = 999
    max_atraso_sub_total = 0
    media_atraso_sub_total = 0
	s_sql = "SELECT" & _
	           " tipo," & _
	           " Count(*) AS qtde," & _
	           " Min(DATEDIFF(day, dt_vencto, dt_competencia)) AS min_atraso," & _
	           " Avg(DATEDIFF(day, dt_vencto, dt_competencia)) AS media_atraso," & _
	           " Max(DATEDIFF(day, dt_vencto, dt_competencia)) AS max_atraso," & _
	           " Sum(valor) AS vl_pago_com_atraso" & _
           " FROM (" & _
	           " SELECT" & _
		           " tC.tipo," & _
		           " dt_competencia," & _
		           " tFFC.valor," & _
		           " dt_vencto," & _
		           " descricao" & _
	           " FROM t_FIN_FLUXO_CAIXA tFFC" & _
			           " LEFT JOIN t_FIN_BOLETO_ITEM tFBI ON (tFBI.id = tFFC.ctrl_pagto_id_parcela) AND (tFFC.ctrl_pagto_modulo = 1)" & _
			           " LEFT JOIN t_FIN_BOLETO tFB ON (tFBI.id_boleto = tFB.id)" & _
			           " LEFT JOIN t_CLIENTE tC ON (tFB.id_cliente = tC.id)" & _
	           " WHERE" & _
		           " (ctrl_pagto_modulo = 1)" & _
		           " AND (st_sem_efeito = 0)" & _
		           " AND (st_confirmacao_pendente = 0)" & _   
                   " AND (ctrl_pagto_id_parcela IN (SELECT id FROM t_FIN_BOLETO_ITEM WHERE (dt_vencto BETWEEN " & bd_formata_data(StrToDate(c_dt_inicio)) & " AND " & bd_formata_data(StrToDate(c_dt_termino)) & ") AND (dt_entrada_confirmada IS NOT NULL)))" & _
            ") t" & _
           " WHERE" & _
	           " dt_competencia > dt_vencto" & _
           " GROUP BY" & _
	           " tipo"


  ' cabeçalho
	cab_table = "<tr>" & chr(13) & _
                "       <td style='border-bottom: 1px solid #000'><span class='N'>BOLETOS COM VENCIMENTO NO PERÍODO E PAGOS COM ATRASO</span></td>" & chr(13) & _
                "</tr>" & chr(13) & _
                "   <td align='center'><br>" & chr(13)

	cab = "<table cellspacing='0'>" & chr(13) & _
          "	<tr style='background:azure' nowrap>" & chr(13) & _
		  "		<td class='Col1' style='background-color:#fff;border:0;'>&nbsp;</td>" & chr(13) & _
		  "		<td class='MTD ME Col2' align='center' valign='bottom' nowrap><span class='Rc'>Qtde</span></td>" & chr(13) & _
		  "		<td class='MTD Col3' align='right' valign='bottom' nowrap><span class='R'>Valor pago com atraso</span></td>" & chr(13) & _
		  "		<td class='MTD Col4' align='center' valign='bottom' nowrap><span class='Rc'>Min atraso</span></td>" & chr(13) & _
		  "		<td class='MTD Col4' align='center' valign='bottom' nowrap><span class='Rc'>Media atraso</span></td>" & chr(13) & _
		  "		<td class='MTD Col4' align='center' valign='bottom' nowrap><span class='Rc'>Max atraso</span></td>" & chr(13) & _
		  "	</tr>" & chr(13)
	
	x = ""
	qtde_sub_total = 0
    vl_sub_total = 0

    x = cab_table & cab
	rs.Open s_sql, cn
	do while Not rs.Eof

       
       x = x & "<tr>" & _
                    "   <td class='ME MC MD Col1' align='center' valign='middle'><span class='Cn'>" & rs("tipo") & "</span></td>" & _
                    "   <td class='MC MD Col2' align='center' valign='middle'><span class='Cn'>" & FormatNumber(rs("qtde"), 0) & "</span></td>" & _
                    "   <td class='MC MD Col3' align='right' valign='middle'><span class='Cn'>" & formata_moeda(rs("vl_pago_com_atraso")) & "</span></td>" & _
		            "   <td class='MC MD Col4' align='center' valign='bottom' nowrap><span class='Cn'>" & FormatNumber(rs("min_atraso"), 0) & "</span></td>" & chr(13) & _
		            "   <td class='MC MD Col4' align='center' valign='bottom' nowrap><span class='Cn'>" & FormatNumber(rs("media_atraso"), 0) & "</span></td>" & chr(13) & _
		            "   <td class='MC MD Col4' align='center' valign='bottom' nowrap><span class='Cn'>" & FormatNumber(rs("max_atraso"), 0) & "</span></td>" & chr(13) & _
               "</tr>" & chr(13)
        
		rs.MoveNext
    loop
    if rs.State <> 0 then rs.Close
    s_sql = "SELECT" & _
	           " Count(*) AS qtde," & _
	           " Coalesce(Min(DATEDIFF(day, dt_vencto, dt_competencia)),0) AS min_atraso," & _
	           " Coalesce(Avg(DATEDIFF(day, dt_vencto, dt_competencia)),0) AS media_atraso," & _
	           " Coalesce(Max(DATEDIFF(day, dt_vencto, dt_competencia)),0) AS max_atraso," & _
	           " Coalesce(Sum(valor),0) AS vl_pago_com_atraso" & _
           " FROM (" & _
	           " SELECT" & _
		           " dt_competencia," & _
		           " tFFC.valor," & _
		           " dt_vencto," & _
		           " descricao" & _
	           " FROM t_FIN_FLUXO_CAIXA tFFC LEFT JOIN t_FIN_BOLETO_ITEM tFBI ON (tFBI.id = tFFC.ctrl_pagto_id_parcela) AND (tFFC.ctrl_pagto_modulo = 1)" & _
	           " WHERE" & _
		           " (ctrl_pagto_modulo = 1)" & _
		           " AND (st_sem_efeito = 0)" & _
		           " AND (st_confirmacao_pendente = 0)" & _
		           " AND (ctrl_pagto_id_parcela IN (SELECT id FROM t_FIN_BOLETO_ITEM WHERE (dt_vencto BETWEEN " & bd_formata_data(StrToDate(c_dt_inicio)) & " AND " & bd_formata_data(StrToDate(c_dt_termino)) & ") AND (dt_entrada_confirmada IS NOT NULL)))" & _
            ") t" & _
           " WHERE" & _
	           " dt_competencia > dt_vencto"

    rs.Open s_sql, cn
    if Not rs.Eof then
        qtde_sub_total = rs("qtde")
        vl_sub_total = rs("vl_pago_com_atraso")    
        max_atraso_sub_total = rs("max_atraso")
        min_atraso_sub_total = rs("min_atraso")
        media_atraso_sub_total = rs("media_atraso")
    end if
    if rs.State <> 0 then rs.Close

    x = x & "<tr>" & chr(13) & _
                    "   <td class='ME MD MC MB Col1' align='center' valign='middle'><span class='C'>TOTAL</span></td>" & chr(13) & _
                    "   <td class='MC MD MB Col2' align='center' valign='middle'><span class='C'>" & FormatNumber(qtde_sub_total, 0) & "</span></td>" & chr(13) & _
                    "   <td class='MC MD MB Col3' align='right' valign='middle'><span class='C'>" & formata_moeda(vl_sub_total) & "</span></td>" & chr(13) & _                    
                    "   <td class='MC MD MB Col4' align='center' valign='middle'><span class='C'>" & FormatNumber(min_atraso_sub_total, 0) & "</span></td>" & chr(13) & _
                    "   <td class='MC MD MB Col4' align='center' valign='middle'><span class='C'>" & FormatNumber(media_atraso_sub_total, 0) & "</span></td>" & chr(13) & _
                    "   <td class='MC MD MB Col4' align='center' valign='middle'><span class='C'>" & FormatNumber(max_atraso_sub_total, 0) & "</span></td>" & chr(13) & _                    
	        "</tr>" & chr(13) & _
        "</table><br><br><br>" & chr(13)

    x = x & "   </td>" & chr(13) & _
            "</tr>" & chr(13)
    

	Response.write x

'	BOLETOS COM VENCIMENTO NO PERÍODO E NÃO PAGOS

	s_sql = "SELECT" & _ 
	           " tipo," & _
	           " Count(*) AS qtde," & _
	           " Sum(valor) AS vl_nao_pago" & _
           " FROM (" & _
	           " SELECT" & _
		           " tC.tipo," & _
		           " dt_competencia," & _
		           " tFFC.valor," & _
		           " dt_vencto," & _
		           " descricao" & _
	           " FROM t_FIN_FLUXO_CAIXA tFFC" & _
			           " LEFT JOIN t_FIN_BOLETO_ITEM tFBI ON (tFBI.id = tFFC.ctrl_pagto_id_parcela) AND (tFFC.ctrl_pagto_modulo = 1)" & _
			           " LEFT JOIN t_FIN_BOLETO tFB ON (tFBI.id_boleto = tFB.id)" & _
			           " LEFT JOIN t_CLIENTE tC ON (tFB.id_cliente = tC.id)" & _
	           " WHERE" & _
		           " (ctrl_pagto_modulo = 1)" & _
		           " AND (st_sem_efeito = 0)" & _
		           " AND (st_confirmacao_pendente = 1)" & _
		           " AND (dt_entrada_confirmada IS NOT NULL)" & _
		           " AND (dt_vencto BETWEEN " & bd_formata_data(StrToDate(c_dt_inicio)) & " AND " & bd_formata_data(StrToDate(c_dt_termino)) & ")" & _
            ") t" & _
           " GROUP BY" & _
	           " tipo"



  ' cabeçalho
	cab_table = "<tr>" & chr(13) & _
                "       <td style='border-bottom: 1px solid #000'><span class='N'>BOLETOS COM VENCIMENTO NO PERÍODO E NÃO PAGOS</span></td>" & chr(13) & _
                "</tr>" & chr(13) & _
                "   <td align='center'><br>" & chr(13)

	cab = "<table cellspacing='0'>" & chr(13) & _
          "	<tr style='background:azure' nowrap>" & chr(13) & _
		  "		<td class='Col1' style='background-color:#fff;border:0;'>&nbsp;</td>" & chr(13) & _
		  "		<td class='MTD ME Col2' align='center' valign='bottom' nowrap><span class='Rc'>Qtde</span></td>" & chr(13) & _
		  "		<td class='MTD Col3' align='right' valign='bottom' nowrap><span class='R'>Valor não pago</span></td>" & chr(13) & _
          "     <td style='border: 0; background-color: #fff;'>&nbsp;</td>" & chr(13) & _
		  "	</tr>" & chr(13)
	
	x = ""
	qtde_sub_total = 0
    vl_sub_total = 0

    x = cab_table & cab
	rs.Open s_sql, cn
	do while Not rs.Eof
       
       x = x & "<tr>" & chr(13) & _
                    "   <td class='ME MC MD Col1' align='center' valign='middle'><span class='Cn'>" & rs("tipo") & "</span></td>" & chr(13) & _
                    "   <td class='MC MD Col2' align='center' valign='middle'><span class='Cn'>" & FormatNumber(rs("qtde"), 0) & "</span></td>" & chr(13) & _
                    "   <td class='MC MD Col3' align='right' valign='middle'><span class='Cn'>" & formata_moeda(rs("vl_nao_pago")) & "</span></td>" & chr(13) & _
                    "   <td style='border: 0; background-color: #fff;'>&nbsp;</td>" & chr(13) & _
               "</tr>" & chr(13)
        qtde_sub_total = qtde_sub_total + rs("qtde")
        vl_sub_total = vl_sub_total + rs("vl_nao_pago")        
		rs.MoveNext
    loop
    if rs.State <> 0 then rs.Close

    x = x & "<tr>" & chr(13) & _
                    "   <td class='ME MD MC MB Col1' align='center' valign='middle'><span class='C'>TOTAL</span></td>" & chr(13) & _
                    "   <td class='MC MD MB Col2' align='center' valign='middle'><span class='C'>" & FormatNumber(qtde_sub_total, 0) & "</span></td>" & chr(13) & _
                    "   <td class='MC MD MB Col3' align='right' valign='middle'><span class='C'>" & formata_moeda(vl_sub_total) & "</span></td>" & chr(13) & _     
                    "   <td style='border: 0; background-color: #fff;'><a href='javascript:fExibeOcultaLinhaPedidosAtrasados()'><img src='../botao/view_bottom.png' title='exibe/oculta relação de pedidos com boletos atrasados'></a></td>" & chr(13) & _                                   
	        "</tr>" & chr(13)

    x = x & "   </td>" & chr(13) & _
            "</tr>" & chr(13)

    ' mostrar relação de pedidos não pagos
    s_sql = "SELECT DISTINCT" & _
                 " tP.data," & _
                 " tFBIR.pedido AS pedido" & _
             " FROM t_FIN_FLUXO_CAIXA tFFC" & _
             " LEFT JOIN t_FIN_BOLETO_ITEM tFBI" & _
                " ON (tFBI.id = tFFC.ctrl_pagto_id_parcela)" & _
                    " AND (tFFC.ctrl_pagto_modulo = 1)" & _
            " LEFT JOIN t_FIN_BOLETO tFB" & _
                " ON (tFBI.id_boleto = tFB.id)" & _
            " LEFT JOIN t_CLIENTE tC" & _
                " ON (tFB.id_cliente = tC.id)" & _
            " LEFT JOIN t_FIN_BOLETO_ITEM_RATEIO tFBIR" & _
                " ON (tFBI.id = tFBIR.id_boleto_item)" & _
            " LEFT JOIN t_PEDIDO tP" & _
                " ON (tFBIR.pedido = tP.pedido)" & _
            " WHERE (ctrl_pagto_modulo = 1)" & _
                " AND (st_sem_efeito = 0)" & _
                " AND (st_confirmacao_pendente = 1)" & _
                " AND (dt_entrada_confirmada IS NOT NULL)" & _
                " AND (" & _
                    " dt_vencto BETWEEN " & bd_formata_data(StrToDate(c_dt_inicio)) & " AND " & bd_formata_data(StrToDate(c_dt_termino)) & "" & _
                     ")" & _
            " ORDER BY" & _
                " tP.data," & _
                " tFBIR.pedido"

    x = x &  "<tr>" & chr(13) & _
          "     <td id='tdPedidosAtrasados' colspan='3' align='left' class='MB ME MD'>" & chr(13) & _
          "         <table id='tblPedidosAtrasados' style='width:100%'>" & chr(13) & _
          "             <tr>" & chr(13) & _
          "                 <td class='Rf' align='left' colspan='3'>PEDIDOS COM BOLETOS NÃO PAGOS</td>" & chr(13) & _
          "             </tr>" & chr(13)
    
    n_pedidos_atrasados = 0
    s_lista_pedidos_atrasados = "<tr>"
    rs.Open s_sql, cn
	do while Not rs.Eof
        n_pedidos_atrasados = n_pedidos_atrasados + 1

        if n_pedidos_atrasados > 4 then 
            n_pedidos_atrasados = 1
            s_lista_pedidos_atrasados = s_lista_pedidos_atrasados & _
                "</tr>" & chr(13) & _
                "<tr>" & chr(13)            
        end if

        s_lista_pedidos_atrasados = s_lista_pedidos_atrasados & "   <td class='C' align='left' width='25%'>" & monta_link_pedido(rs("pedido")) & "</td>" & chr(13) & _
        

        rs.MoveNext
    loop
    if rs.State <> 0 then rs.Close

    for i=n_pedidos_atrasados+1 to 4
						s_lista_pedidos_atrasados = s_lista_pedidos_atrasados & _
							"					<td align='left' width='25%'>&nbsp;</td>" & chr(13)
						next
					s_lista_pedidos_atrasados = s_lista_pedidos_atrasados & "				</tr>" & chr(13)
    
    x = x & s_lista_pedidos_atrasados
    x = x & "   </table>" & chr(13)
    x = x & "</table><br><br><br>" & chr(13)
        
	Response.write x

'	TOTAL DE RECEBIMENTOS REFERENTE A BOLETOS NO PERÍODO

	s_sql = "SELECT" & _
	           " tipo," & _
	           " Count(*) AS qtde," & _
	           " Sum(valor) AS vl_recebimentos" & _
           " FROM (" & _
	           " SELECT" & _
		           " CASE LEN(Coalesce(cnpj_cpf,'')) WHEN 11 THEN 'PF' WHEN 14 THEN 'PJ' ELSE '' END AS tipo," & _
		           " dt_competencia," & _
		           " tFFC.valor," & _
		           " dt_vencto," & _
		           " descricao" & _
	           " FROM t_FIN_FLUXO_CAIXA tFFC" & _
			           " LEFT JOIN t_FIN_BOLETO_ITEM tFBI ON (tFBI.id = tFFC.ctrl_pagto_id_parcela) AND (tFFC.ctrl_pagto_modulo = 1) AND (dt_entrada_confirmada IS NOT NULL)" & _
	           " WHERE" & _
		           " (" & _
			           " (ctrl_pagto_modulo = 1)" & _ 
			           " OR" & _ 
			           " ((id_plano_contas_conta = 9921) AND (natureza='C') AND ((descricao LIKE '%-DEP%') OR (descricao LIKE '% DEP%')) AND (Len(Coalesce(tFFC.cnpj_cpf,''))>0))" & _
		            ")" & _
		           " AND (st_sem_efeito = 0)" & _
		           " AND (st_confirmacao_pendente = 0)" & _
		           " AND (dt_competencia BETWEEN " & bd_formata_data(StrToDate(c_dt_inicio)) & " AND " & bd_formata_data(StrToDate(c_dt_termino)) & ")" & _
            ") t" & _
           " GROUP BY" & _
	           " tipo"



  ' cabeçalho
	cab_table = "<tr>" & chr(13) & _
                "       <td style='border-bottom: 1px solid #000'><span class='N'>TOTAL DE RECEBIMENTOS REFERENTE A BOLETOS NO PERÍODO</span></td>" & chr(13) & _
                "</tr>" & chr(13) & _
                "   <td align='center'><br>" & chr(13)

	cab = "<table cellspacing='0'>" & chr(13) & _
          "	<tr style='background:azure' nowrap>" & chr(13) & _
		  "		<td class='Col1' style='background-color:#fff;border:0;'>&nbsp;</td>" & chr(13) & _
		  "		<td class='MTD ME Col2' align='center' valign='bottom' nowrap><span class='Rc'>Qtde</span></td>" & chr(13) & _
		  "		<td class='MTD Col3' align='right' valign='bottom' nowrap><span class='R'>Valor recebimentos</span></td>" & chr(13) & _
		  "	</tr>" & chr(13)
	
	x = ""
	qtde_sub_total = 0
    vl_sub_total = 0

    x = cab_table & cab
	rs.Open s_sql, cn
	do while Not rs.Eof
       
       x = x & "<tr>" & chr(13) & _
                    "   <td class='ME MC MD Col1' align='center' valign='middle'><span class='Cn'>" & rs("tipo") & "</span></td>" & chr(13) & _
                    "   <td class='MC MD Col2' align='center' valign='middle'><span class='Cn'>" & FormatNumber(rs("qtde"), 0) & "</span></td>" & chr(13) & _
                    "   <td class='MC MD Col3' align='right' valign='middle'><span class='Cn'>" & formata_moeda(rs("vl_recebimentos")) & "</span></td>" & chr(13) & _
               "</tr>" & chr(13)
        qtde_sub_total = qtde_sub_total + rs("qtde")
        vl_sub_total = vl_sub_total + rs("vl_recebimentos")        
		rs.MoveNext
    loop
    if rs.State <> 0 then rs.Close

    x = x & "<tr>" & _
                    "   <td class='ME MD MC MB Col1' align='center' valign='middle'><span class='C'>TOTAL</span></td>" & chr(13) & _
                    "   <td class='MC MD MB Col2' align='center' valign='middle'><span class='C'>" & FormatNumber(qtde_sub_total, 0) & "</span></td>" & chr(13) & _
                    "   <td class='MC MD MB Col3' align='right' valign='middle'><span class='C'>" & formata_moeda(vl_sub_total) & "</span></td>" & chr(13) & _                    
	        "</tr>" & chr(13) & _
        "</table><br>" & chr(13)

    x = x & "   </td>" & chr(13) & _
            "</tr>" & chr(13)

    
        
	Response.write x


    '----
    s_sql = "SELECT" & _
	           " origem," & _
	           " tipo," & _
	           " Count(*) AS qtde," & _
	           " Sum(valor) AS vl_recebimentos" & _
           " FROM (" & _
	           " SELECT" & _
		           " 'BOLETO' AS origem," & _
		           " CASE LEN(Coalesce(cnpj_cpf,'')) WHEN 11 THEN 'PF' WHEN 14 THEN 'PJ' ELSE '' END AS tipo," & _
		           " dt_competencia," & _
		           " tFFC.valor," & _
		           " dt_vencto," & _
		           " descricao" & _
	           " FROM t_FIN_FLUXO_CAIXA tFFC" & _
			           " LEFT JOIN t_FIN_BOLETO_ITEM tFBI ON (tFBI.id = tFFC.ctrl_pagto_id_parcela) AND (tFFC.ctrl_pagto_modulo = 1) AND (dt_entrada_confirmada IS NOT NULL)" & _
	           " WHERE" & _
		           " (" & _
						"(ctrl_pagto_modulo = 1) AND (NOT ((descricao LIKE '%-DEP%') OR (descricao LIKE '% DEP%')))" & _
					")" & _
		           " AND (st_sem_efeito = 0)" & _
		           " AND (st_confirmacao_pendente = 0)" & _
		           " AND (dt_competencia BETWEEN " & bd_formata_data(StrToDate(c_dt_inicio)) & " AND " & bd_formata_data(StrToDate(c_dt_termino)) & ")" & _
	           " UNION ALL" & _
	           " SELECT" & _
		           " 'DEPOSITO' AS origem," & _
		           " CASE LEN(Coalesce(cnpj_cpf,'')) WHEN 11 THEN 'PF' WHEN 14 THEN 'PJ' ELSE '' END AS tipo," & _
		           " dt_competencia," & _
		           " tFFC.valor," & _
		           " dt_vencto," & _
		           " descricao" & _
	           " FROM t_FIN_FLUXO_CAIXA tFFC" & _
			           " LEFT JOIN t_FIN_BOLETO_ITEM tFBI ON (tFBI.id = tFFC.ctrl_pagto_id_parcela) AND (tFFC.ctrl_pagto_modulo = 1) AND (dt_entrada_confirmada IS NOT NULL)" & _
	           " WHERE" & _
					"(" & _
						" ((id_plano_contas_conta = 9921) AND (natureza='C') AND ((descricao LIKE '%-DEP%') OR (descricao LIKE '% DEP%')) AND (Len(Coalesce(tFFC.cnpj_cpf,''))>0))" & _
						" OR " & _
						"((ctrl_pagto_modulo = 1) AND ((descricao LIKE '%-DEP%') OR (descricao LIKE '% DEP%')))" & _
					")" & _
		           " AND (st_sem_efeito = 0)" & _
		           " AND (st_confirmacao_pendente = 0)" & _
		           " AND (dt_competencia BETWEEN " & bd_formata_data(StrToDate(c_dt_inicio)) & " AND " & bd_formata_data(StrToDate(c_dt_termino)) & ")" & _
            ") t" & _
           " GROUP BY" & _
	           " origem," & _
	           " tipo" & _
           " ORDER BY" & _
	           " origem," & _
	           " tipo"


  ' cabeçalho
	cab_table = "<tr>" & chr(13) & _
                "       <td style='border:0'><span class='N'>&nbsp;</span></td>" & chr(13) & _
                "</tr>" & chr(13) & _
                "   <td align='center'><br>" & chr(13)

	cab = "<table cellspacing='0'>" & chr(13) & _
          "	<tr style='background:azure' nowrap>" & chr(13) & _
		  "		<td class='Col1' style='background-color:#fff;border:0;' colspan='2'>&nbsp;</td>" & chr(13) & _
		  "		<td class='MTD ME Col2' align='center' valign='bottom' nowrap><span class='Rc'>Qtde</span></td>" & chr(13) & _
		  "		<td class='MTD Col3' align='right' valign='bottom' nowrap><span class='R'>Valor recebimentos</span></td>" & chr(13) & _
		  "	</tr>" & chr(13)
	
	x = ""
	qtde_sub_total = 0
    vl_sub_total = 0

    x = cab_table & cab
	rs.Open s_sql, cn
	do while Not rs.Eof
       
       x = x & "<tr>" & chr(13) & _
                    "   <td class='ME MC MD Col1' align='center' valign='middle'><span class='Cn'>" & rs("origem") & "</span></td>" & chr(13) & _
                    "   <td class='MC MD Col1' align='center' valign='middle'><span class='Cn'>" & rs("tipo") & "</span></td>" & chr(13) & _
                    "   <td class='MC MD Col2' align='center' valign='middle'><span class='Cn'>" & FormatNumber(rs("qtde"), 0) & "</span></td>" & chr(13) & _
                    "   <td class='MC MD Col3' align='right' valign='middle'><span class='Cn'>" & formata_moeda(rs("vl_recebimentos")) & "</span></td>" & chr(13) & _
               "</tr>" & chr(13)
        qtde_sub_total = qtde_sub_total + rs("qtde")
        vl_sub_total = vl_sub_total + rs("vl_recebimentos")        
		rs.MoveNext
    loop
    if rs.State <> 0 then rs.Close

    x = x & "<tr>" & _
                    "   <td class='ME MD MC MB Col1' align='center' valign='middle' colspan='2'><span class='C'>TOTAL</span></td>" & chr(13) & _
                    "   <td class='MC MD MB Col2' align='center' valign='middle'><span class='C'>" & FormatNumber(qtde_sub_total, 0) & "</span></td>" & chr(13) & _
                    "   <td class='MC MD MB Col3' align='right' valign='middle'><span class='C'>" & formata_moeda(vl_sub_total) & "</span></td>" & chr(13) & _                    
	        "</tr>" & chr(13) & _
        "</table><br><br><br>" & chr(13)

    x = x & "   </td>" & chr(13) & _
            "</tr>" & chr(13)

    
        
	Response.write x


'	TOTAL DE RECEBIMENTOS NO PERÍODO REFERENTE A BOLETOS COM VENCTO EM MESES ANTERIORES
    min_atraso_sub_total = 999
    max_atraso_sub_total = 0
    media_atraso_sub_total = 0
	s_sql = "SELECT" & _
	           " tipo," & _
	           " Count(*) AS qtde," & _
	           " Min(DATEDIFF(day, dt_vencto, dt_competencia)) AS min_atraso," & _
	           " Avg(DATEDIFF(day, dt_vencto, dt_competencia)) AS media_atraso," & _
	           " Max(DATEDIFF(day, dt_vencto, dt_competencia)) AS max_atraso," & _
	           " Sum(valor) AS vl_recebimentos_ref_meses_anteriores" & _
           " FROM (" & _
	           " SELECT" & _
		           " CASE LEN(Coalesce(cnpj_cpf,'')) WHEN 11 THEN 'PF' WHEN 14 THEN 'PJ' ELSE '' END AS tipo," & _
		           " dt_competencia," & _
		           " tFFC.valor," & _
		           " dt_vencto," & _
		           " descricao" & _
	           " FROM t_FIN_FLUXO_CAIXA tFFC LEFT JOIN t_FIN_BOLETO_ITEM tFBI ON (tFBI.id = tFFC.ctrl_pagto_id_parcela) AND (tFFC.ctrl_pagto_modulo = 1) AND (dt_entrada_confirmada IS NOT NULL)" & _
	           " WHERE" & _
		           " (" & _ 
			           " (ctrl_pagto_modulo = 1)" & _
			           " OR" & _ 
			           " ((id_plano_contas_conta = 9921) AND (natureza='C') AND ((descricao LIKE '%-DEP%') OR (descricao LIKE '% DEP%')) AND (Len(Coalesce(cnpj_cpf,''))>0))" & _
		            ")" & _
		           " AND (st_sem_efeito = 0)" & _
		           " AND (st_confirmacao_pendente = 0)" & _
		           " AND (dt_competencia BETWEEN " & bd_formata_data(StrToDate(c_dt_inicio)) & " AND " & bd_formata_data(StrToDate(c_dt_termino)) & ")" & _
            ") t" & _
           " WHERE" & _
	           " dt_vencto < " & bd_formata_data(StrToDate(c_dt_inicio)) & "" & _
           " GROUP BY" & _
	           " tipo"


  ' cabeçalho
	cab_table = "   <tr>" & chr(13) & _
                "       <td style='border-bottom: 1px solid #000'><span class='N'>TOTAL DE RECEBIMENTOS NO PERÍODO REFERENTE A BOLETOS COM VENCIMENTO EM MESES ANTERIORES</span></td>" & chr(13) & _
                "   </tr>" & chr(13) & _
                "<td align='center'><br>" & chr(13)

	cab = "<table cellspacing='0'>" & chr(13) & _
          "	<tr style='background:azure' nowrap>" & chr(13) & _
		  "		<td class='Col1' style='background-color:#fff;border:0;'>&nbsp;</td>" & chr(13) & _
		  "		<td class='MTD ME Col2' align='center' valign='bottom' nowrap><span class='Rc'>Qtde</span></td>" & chr(13) & _
		  "		<td class='MTD Col3' align='right' valign='bottom' nowrap><span class='R'>Valor recebimentos</span></td>" & chr(13) & _
		  "		<td class='MTD Col4' align='center' valign='bottom' nowrap><span class='Rc'>Min atraso</span></td>" & chr(13) & _
		  "		<td class='MTD Col4' align='center' valign='bottom' nowrap><span class='Rc'>Media atraso</span></td>" & chr(13) & _
		  "		<td class='MTD Col4' align='center' valign='bottom' nowrap><span class='Rc'>Max atraso</span></td>" & chr(13) & _
		  "	</tr>" & chr(13)
	
	x = ""
	qtde_sub_total = 0
    vl_sub_total = 0

    x = cab_table & cab
	rs.Open s_sql, cn
	do while Not rs.Eof
           
       x = x & "<tr>" & chr(13) & _
                    "   <td class='ME MC MD Col1' align='center' valign='middle'><span class='Cn'>" & rs("tipo") & "</span></td>" & chr(13) & _
                    "   <td class='MC MD Col2' align='center' valign='middle'><span class='Cn'>" & FormatNumber(rs("qtde"), 0) & "</span></td>" & chr(13) & _
                    "   <td class='MC MD Col3' align='right' valign='middle'><span class='Cn'>" & formata_moeda(rs("vl_recebimentos_ref_meses_anteriores")) & "</span></td>" & chr(13) & _
		            "   <td class='MC MD Col4' align='center' valign='bottom' nowrap><span class='Cn'>" & FormatNumber(rs("min_atraso"), 0) & "</span></td>" & chr(13) & _
		            "   <td class='MC MD Col4' align='center' valign='bottom' nowrap><span class='Cn'>" & FormatNumber(rs("media_atraso"), 0) & "</span></td>" & chr(13) & _
		            "   <td class='MC MD Col4' align='center' valign='bottom' nowrap><span class='Cn'>" & FormatNumber(rs("max_atraso"), 0) & "</span></td>" & chr(13) & _
               "</tr>" & chr(13)
        qtde_sub_total = qtde_sub_total + rs("qtde")
        vl_sub_total = vl_sub_total + rs("vl_recebimentos_ref_meses_anteriores")    
        if rs("max_atraso") > max_atraso_sub_total then max_atraso_sub_total = rs("max_atraso")
        if rs("min_atraso") < min_atraso_sub_total then min_atraso_sub_total = rs("min_atraso")
        media_atraso_sub_total = media_atraso_sub_total + rs("media_atraso")
		rs.MoveNext
    loop
    if rs.State <> 0 then rs.Close
    s_sql = "SELECT" & _
	           " Count(*) AS qtde," & _
	           " Coalesce(Min(DATEDIFF(day, dt_vencto, dt_competencia)), 0) AS min_atraso," & _
	           " Coalesce(Avg(DATEDIFF(day, dt_vencto, dt_competencia)), 0) AS media_atraso," & _
	           " Coalesce(Max(DATEDIFF(day, dt_vencto, dt_competencia)), 0) AS max_atraso," & _
	           " Coalesce(Sum(valor), 0) AS vl_recebimentos_ref_meses_anteriores" & _
           " FROM (" & _
	           " SELECT " & _
		           " dt_competencia," & _
		           " tFFC.valor," & _
		           " dt_vencto," & _
		           " descricao" & _
	           " FROM t_FIN_FLUXO_CAIXA tFFC LEFT JOIN t_FIN_BOLETO_ITEM tFBI ON (tFBI.id = tFFC.ctrl_pagto_id_parcela) AND (tFFC.ctrl_pagto_modulo = 1) AND (dt_entrada_confirmada IS NOT NULL)" & _
	           " WHERE" & _
		           " ( " & _
			           " (ctrl_pagto_modulo = 1) " & _
			           " OR " & _
			           " ((id_plano_contas_conta = 9921) AND (natureza='C') AND ((descricao LIKE '%-DEP%') OR (descricao LIKE '% DEP%')) AND (Len(Coalesce(cnpj_cpf,''))>0))" & _
		            ")" & _
		           " AND (st_sem_efeito = 0)" & _
		           " AND (st_confirmacao_pendente = 0)" & _
		           " AND (dt_competencia BETWEEN " & bd_formata_data(StrToDate(c_dt_inicio)) & " AND " & bd_formata_data(StrToDate(c_dt_termino)) & ")" & _
            ") t" & _
           " WHERE" & _
	           " dt_vencto < " & bd_formata_data(StrToDate(c_dt_inicio))

    rs.Open s_sql, cn
    if Not rs.Eof then
        qtde_sub_total = rs("qtde")
        vl_sub_total = rs("vl_recebimentos_ref_meses_anteriores") + 0 
        max_atraso_sub_total = rs("max_atraso") + 0
        min_atraso_sub_total = rs("min_atraso") + 0
        media_atraso_sub_total = rs("media_atraso") + 0
    
    end if
    if rs.State <> 0 then rs.Close


    x = x & "<tr>" & chr(13) & _
                    "   <td class='ME MD MC MB Col1' align='center' valign='middle'><span class='C'>TOTAL</span></td>" & chr(13) & _
                    "   <td class='MC MD MB Col2' align='center' valign='middle'><span class='C'>" & FormatNumber(qtde_sub_total, 0) & "</span></td>" & chr(13) & _
                    "   <td class='MC MD MB Col3' align='right' valign='middle'><span class='C'>" & formata_moeda(vl_sub_total) & "</span></td>" & chr(13) & _                    
                    "   <td class='MC MD MB Col4' align='center' valign='middle'><span class='C'>" & FormatNumber(min_atraso_sub_total, 0) & "</span></td>" & chr(13) & _
                    "   <td class='MC MD MB Col4' align='center' valign='middle'><span class='C'>" & FormatNumber(media_atraso_sub_total, 0) & "</span></td>" & chr(13) & _
                    "   <td class='MC MD MB Col4' align='center' valign='middle'><span class='C'>" & FormatNumber(max_atraso_sub_total, 0) & "</span></td>" & chr(13) & _                    
	        "</tr>" & chr(13) & _
        "</table><br><br><br>" & chr(13)

    x = x & "   </td>" & chr(13) & _
            "</tr>" & chr(13)
    if rs.State <> 0 then rs.Close

	Response.write x

'	TOTAL DE RECEBIMENTOS NO PERÍODO REFERENTE A BOLETOS COM VENCIMENTO EM MESES ANTERIORES (MÊS A MÊS)
    mes_a = "xxx"
    vl_total = 0
    qtde_total = 0
	s_sql = "SELECT" & _
	           " Left(Convert(varchar(10), dt_vencto, 121), 7) AS mes_ref_dt_vencto," & _
	           " tipo," & _
	           " Count(*) AS qtde," & _
	           " Sum(valor) AS vl_recebimentos_ref_meses_anteriores" & _
           " FROM (" & _
	           " SELECT" & _
		           " CASE LEN(Coalesce(cnpj_cpf,'')) WHEN 11 THEN 'PF' WHEN 14 THEN 'PJ' ELSE '' END AS tipo," & _
		           " dt_competencia," & _
		           " tFFC.valor," & _
		           " dt_vencto," & _
		           " descricao" & _
	           " FROM t_FIN_FLUXO_CAIXA tFFC LEFT JOIN t_FIN_BOLETO_ITEM tFBI ON (tFBI.id = tFFC.ctrl_pagto_id_parcela) AND (tFFC.ctrl_pagto_modulo = 1) AND (dt_entrada_confirmada IS NOT NULL)" & _
	           " WHERE" & _
		           " (" & _
			           " (ctrl_pagto_modulo = 1)" & _
			           " OR" & _ 
			           " ((id_plano_contas_conta = 9921) AND (natureza='C') AND ((descricao LIKE '%-DEP%') OR (descricao LIKE '% DEP%')) AND (Len(Coalesce(cnpj_cpf,''))>0))" & _
		            ")" & _
		           " AND (st_sem_efeito = 0)" & _
		           " AND (st_confirmacao_pendente = 0)" & _
		           " AND (dt_competencia BETWEEN " & bd_formata_data(StrToDate(c_dt_inicio)) & " AND " & bd_formata_data(StrToDate(c_dt_termino)) & ")" & _
            ") t" & _
           " WHERE" & _
	           " dt_vencto < " & bd_formata_data(StrToDate(c_dt_inicio)) & "" & _
           " GROUP BY" & _
	           " Left(Convert(varchar(10), dt_vencto, 121), 7)," & _
	           " tipo" & _
           " ORDER BY" & _
	           " Left(Convert(varchar(10), dt_vencto, 121), 7)," & _
	           " tipo"




  ' cabeçalho
	cab_table = "<tr>" & chr(13) & _
                "       <td style='border-bottom: 1px solid #000'><span class='N'>TOTAL DE RECEBIMENTOS NO PERÍODO REFERENTE A BOLETOS COM VENCIMENTO EM MESES ANTERIORES (MÊS A MÊS)</span></td>" & chr(13) & _
                "</tr>" & chr(13) & _
                "   <td align='center'><br>" & chr(13)

	cab = "<table cellspacing='0'>" & chr(13) & _
          "	<tr style='background:azure' nowrap>" & chr(13) & _
		  "		<td class='Col1' style='background-color:#fff;border:0;' colspan='2'>&nbsp;</td>" & chr(13) & _
		  "		<td class='MTD ME Col2' align='center' valign='bottom' nowrap><span class='Rc'>Qtde</span></td>" & chr(13) & _
		  "		<td class='MTD Col3' align='right' valign='bottom' nowrap><span class='R'>Valor recebimentos</span></td>" & chr(13) & _
		  "	</tr>" & chr(13)
	
	x = ""
	qtde_sub_total = 0
    vl_sub_total = 0

    x = cab_table & cab
	rs.Open s_sql, cn
	do while Not rs.Eof
       if mes_a <> rs("mes_ref_dt_vencto") then
            mes_a = rs("mes_ref_dt_vencto")             
            if n_reg > 0 then
                x = x & "<tr>" & chr(13) & _
                        "   <td class='ME MD MC MB Col1' align='center' valign='middle' colspan='2'><span class='C'>TOTAL</span></td>" & chr(13) & _
                        "   <td class='MC MD MB Col2' align='center' valign='middle'><span class='C'>" & FormatNumber(qtde_sub_total, 0) & "</span></td>" & chr(13) & _
                        "   <td class='MC MD MB Col3' align='right' valign='middle'><span class='C'>" & formata_moeda(vl_sub_total) & "</span></td>" & chr(13) & _                    
	                "</tr>" & chr(13) & _
                    "<tr>" & chr(13) & _
                            "   <td colspan='4' style='border:0'>&nbsp;</td>" & chr(13) & _
                    "</tr>" & chr(13)

                qtde_sub_total = 0
                vl_sub_total = 0
            end if
        end if
        mes = DatePart("M", rs("mes_ref_dt_vencto"))
        if Len(mes) = 1 then mes = "0" & mes
        x = x & "<tr>" & _
                    "   <td class='ME MC MD Col1' align='center' valign='middle'><span class='C'>" & mes & "/" & DatePart("yyyy", rs("mes_ref_dt_vencto")) & "</span></td>" & chr(13) & _                    
                    "   <td class='MC MD Col1' align='center' valign='middle'><span class='Cn'>" & rs("tipo") & "</span></td>" & chr(13) & _
                    "   <td class='MC MD Col2' align='center' valign='middle'><span class='Cn'>" & FormatNumber(rs("qtde"), 0) & "</span></td>" & chr(13) & _
                    "   <td class='MC MD Col3' align='right' valign='middle'><span class='Cn'>" & formata_moeda(rs("vl_recebimentos_ref_meses_anteriores")) & "</span></td>" & chr(13) & _
               "</tr>"
        qtde_total = qtde_total + rs("qtde")
        vl_total = vl_total + rs("vl_recebimentos_ref_meses_anteriores") 
        qtde_sub_total = qtde_sub_total + rs("qtde")
        vl_sub_total = vl_sub_total + rs("vl_recebimentos_ref_meses_anteriores") 

        n_reg = n_reg + 1
		rs.MoveNext
    loop
    ' total do último
    x = x & "<tr>" & chr(13) & _
                        "   <td class='ME MD MC MB Col1' align='center' valign='middle' colspan='2'><span class='C'>TOTAL</span></td>" & chr(13) & _
                        "   <td class='MC MD MB Col2' align='center' valign='middle'><span class='C'>" & FormatNumber(qtde_sub_total, 0) & "</span></td>" & chr(13) & _
                        "   <td class='MC MD MB Col3' align='right' valign='middle'><span class='C'>" & formata_moeda(vl_sub_total) & "</span></td>" & chr(13) & _                    
	            "</tr>" & chr(13) & _
                "<tr>" & chr(13) & _
                        "<td colspan='4' style='border:0'>&nbsp;</td>" & chr(13) & _
                "</tr>" & chr(13)

    if rs.State <> 0 then rs.Close

    ' total geral
    x = x & "<tr>" & chr(13) & _
                    "   <td class='ME MD MC MB Col1' align='center' valign='middle' colspan='2' style='background:honeydew'><span class='C'>TOTAL GERAL</span></td>" & chr(13) & _
                    "   <td class='MC MD MB Col2' align='center' valign='middle' style='background:honeydew'><span class='C'>" & FormatNumber(qtde_total, 0) & "</span></td>" & chr(13) & _
                    "   <td class='MC MD MB Col3' align='right' valign='middle' style='background:honeydew'><span class='C'>" & formata_moeda(vl_total) & "</span></td>" & chr(13) & _                    
	        "</tr>" & chr(13) & _
        "</table><br><br><br>" & chr(13)

    x = x & "   </td>" & chr(13) & _
            "</tr>" & chr(13)
        
	Response.write x


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
<script src="<%=URL_FILE__JQUERY_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
var windowScrollTopAnterior;
window.status = 'Aguarde, executando a consulta ...';

$(function() {
    $("#tblPedidosAtrasados").hide();
    $("#tdPedidosAtrasados").removeClass('MB ME MD');
});

function fPEDConsulta(id_pedido) {
    window.status = "Aguarde ...";
    
    fREL.pedido_selecionado.value = id_pedido;
    fREL.action = "pedido.asp"
    fREL.submit();
}

function fExibeOcultaLinhaPedidosAtrasados() {
    if ($("#tblPedidosAtrasados").is(":visible")) {
        $("#tblPedidosAtrasados").hide();
        $("#tdPedidosAtrasados").removeClass('MB ME MD');
    }
    else {
        $("#tblPedidosAtrasados").show();
        $("#tdPedidosAtrasados").addClass('MB ME MD');
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
.Col1 {
    width: 80px;
}
.Col2 {
    width: 100px;
}
.Col3 {
    width: 130px;
}
.Col4 {
    width: 70px;
}
</style>

<% if alerta <> "" then %>
<!-- ************************************************************ -->
<!-- **********  PÁGINA PARA EXIBIR MENSAGENS DE ERRO  ********** -->
<!-- ************************************************************ -->
<body onload="bVOLTAR.focus();" alink="#000000" vlink="#000000" link="#000000">
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



<% else
     %>
<!-- ***************************************************** -->
<!-- **********  PÁGINA PARA EXIBIR RESULTADO   ********** -->
<!-- ***************************************************** -->
<body onload="window.status='Concluído';">

<center>

<form id="fREL" name="fREL" method="post">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="pedido_selecionado" id="pedido_selecionado" value="">

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="709" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Perfil de Pagamento dos Boletos</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>

<!-- FILTROS -->
<%
	s_filtro = "<table width='709' cellpadding='0' cellspacing='0' style='border-bottom:1px solid black' border='0'>" & chr(13)

'	PERÍODO: ENTRE
	s = ""

	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' nowrap><span class='N'>Período:&nbsp;</span></td>" & chr(13) & _
				"		<td align='left' valign='top' width='99%'><span class='N'>" & c_dt_inicio & " a " & c_dt_termino  & "</span></td>" & chr(13) & _
				"	</tr>" & chr(13)

'	EMISSÃO
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>Emissão:&nbsp;</span></td><td align='left' valign='top' width='99%'>" & _
			   "<span class='N'>" & formata_data_hora(Now) & "</span></td></tr>" & chr(13)

	s_filtro = s_filtro & "</table>" & chr(13)
	Response.Write s_filtro
%>

<!--  RELATÓRIO  -->
<br>
<% consulta_executa %>

<!-- ************   SEPARADOR   ************ -->
<table class="notPrint" width="709" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
	<td class="Rc" align="left">&nbsp;</td>
</tr>
</table>


<br />
<table class="notPrint" width="709" cellspacing="0">
<tr>
	<td align="center" style="width:100%"><a name="bVOLTA" id="bVOLTA" href="javascript:history.back()" title="volta para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a>
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
