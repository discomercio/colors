<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=False %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =======================================================================================
'	  R E L C O M I S S A O I N D I C A D O R E S C O N S U L T A P E D I D O E X E C . A S P
'     =======================================================================================
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
	dim cn, rs, msg_erro,rs2
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
    If Not cria_recordset_otimista(rs2, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)


	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	if Not operacao_permitida(OP_CEN_REL_COMISSAO_INDICADORES, s_lista_operacoes_permitidas) then 
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	dim blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos
	blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos = isActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos

	dim alerta
	dim s, s_aux, s_filtro, s_id,mes_competencia,ano_competencia,vendedor,c_dt_entregue_ano,c_dt_entregue_mes,ckb_Desc
    dim aviso, rb_visao, blnVisaoSintetica, s_vendedor, s_mes_competencia, s_ano_competencia
    
    blnVisaoSintetica = False
	if rb_visao = "SINTETICA" then blnVisaoSintetica = True
	alerta = ""
    s_vendedor=""
    ckb_Desc= ""

    c_dt_entregue_ano =  Trim(Request("c_dt_entregue_ano"))
    c_dt_entregue_mes = Trim(Request("c_dt_entregue_mes"))
    s_id = Request("id")

    s_vendedor = Request("vendedor")
    s_ano_competencia = Request("ano_competencia")
    s_mes_competencia = Request("mes_competencia")
    ckb_Desc = Trim(Request("ckb_Desc"))

dim o
dim resultadoCalculo,resultadoDigito,limitador(5)
dim dadosCalculo
set o = createobject( "ComPlusCalcCedulas.ComPlusCalcCedulas" )
dim v_cedulas,aux(5),y,cont
dim cedulas()
dim qtdeCedula()

' _____________________________________
' CONSULTA EXECUTA
'
sub consulta_executa
const VENDA_NORMAL = "VENDA_NORMAL"
const DEVOLUCAO = "DEVOLUCAO"
const PERDA = "PERDA"
dim r,qtdeVendedor
dim s, s_aux, s_sql, x, cab_table, cab, meio_pagto_a, n_reg, n_reg_total, qtde_indicadores, indicador_a, vendedor_a
dim idx_bloco, inc, comissao ,indice
dim v_Banco,strAuxbanco,blnAchou,vl_aux,n_reg_BD,intIdxBanco,strAuxBancoAnterior,intIdxVetor,strCampoOrdenacao,v_OutrosBancos
dim ind_anterior,atual,vl_preco_venda,perc_RT,vl_RT,vl_preco_NF, vl_RA,vl_RA_liquido,vl_RA_diferenca,vl_sub_total_preco_venda,vl_total_preco_venda
dim vl_sub_total_RT, vl_sub_total_RT_arredondado,vl_total_RT,vl_sub_total_RA,vl_total_RA,vl_sub_total_RA_liquido,vl_total_RA_liquido ,vl_sub_total_RA_arredondado 
dim vl_sub_total_RA_diferenca,vl_total_RA_diferenca,sub_total_comissao,s_lista_completa_pedidos ,s_class ,s_class_td ,s_cor,s_cor_sinal,s_sinal,s_nome_cliente   
 dim s_lista_vl_pedido,s_lista_comissao,s_lista_RA_bruto,s_lista_RA_liquido,banco,s_lista_total_comissao,s_checked
 dim msg_desconto,s_lista_meio_pagto,s_banco,s_agencia,s_conta,s_favorecido,s_banco_nome,s_desempenho_nota,s_new_cab 
dim vl_sub_total_preco_NF,s_disabled,s_lista_total_comissao_arredondado,vl_comissao,vl_total_RT_arredondado,totalComChqOutros,qtdeChq,totalComChqBradesco,totalComDin 
dim totalDEP,totalCHQ,totalDIN,totalComissaoArredondado, meio_pagto,st_tratamento_manual,vl_sub_total_comissao,qtdeCedula(),v_cedulas,z,cont,cedulas_codificado,aux(5)
dim textDin,sub_total_com_RT,sub_total_com_RA,sub_total_com,contador,v_desconto_descricao(),v_desconto_valor(),qtde_registro_desc,valor_desconto
dim vChequesPorBanco(), vDinheiroPorBanco(), i, idx
dim regex

    set regex = New RegExp
    regex.Pattern = "1+$"

sub_total_com = 0
textDin = ""
z = 0  
totalDIN = 0
totalDEP = 0 
totalCHQ = 0
qtdeChq = 0   
st_tratamento_manual = 0   
meio_pagto = ""   
s_disabled = " disabled"  
cedulas_codificado= 0
            	
    s_sql = "SELECT t_COMISSAO_INDICADOR_N3.indicador as ind3,t_COMISSAO_INDICADOR_N3.id as id_n3,t_COMISSAO_INDICADOR_N2.id as id_n2, * FROM t_COMISSAO_INDICADOR_N1" & _
            " INNER JOIN t_COMISSAO_INDICADOR_N2 ON (t_COMISSAO_INDICADOR_N1.id = t_COMISSAO_INDICADOR_N2.id_comissao_indicador_n1)" & _
            " INNER JOIN t_COMISSAO_INDICADOR_N3 ON (t_COMISSAO_INDICADOR_N2.id = t_COMISSAO_INDICADOR_N3.id_comissao_indicador_n2)" & _            
            " INNER JOIN t_COMISSAO_INDICADOR_N4 ON (t_COMISSAO_INDICADOR_N3.id = t_COMISSAO_INDICADOR_N4.id_comissao_indicador_n3)" & _
            " INNER JOIN t_PEDIDO ON (t_COMISSAO_INDICADOR_N4.pedido = t_PEDIDO.pedido)" & _
            " INNER JOIN t_CLIENTE ON (t_PEDIDO.id_cliente=t_CLIENTE.id)" & _
            " WHERE (t_COMISSAO_INDICADOR_N1.id = '" & s_id & "')"
            if ckb_Desc = "1" then s_sql = s_sql & " AND (t_COMISSAO_INDICADOR_N3.cod_motivo_tratamento_manual = '1')" 
            s_sql = s_sql &_
            " ORDER BY t_COMISSAO_INDICADOR_N2.vendedor," & _
            " t_COMISSAO_INDICADOR_N3.indicador," & _
            " t_COMISSAO_INDICADOR_N3.numero_banco," & _
            " t_COMISSAO_INDICADOR_N3.vl_total_comissao_arredondado"

    if s_vendedor <> "" then
        s_sql = "SELECT t_COMISSAO_INDICADOR_N3.indicador as ind3,t_COMISSAO_INDICADOR_N3.id as id_n3,t_COMISSAO_INDICADOR_N2.id as id_n2,* FROM t_COMISSAO_INDICADOR_N1" & _
            " INNER JOIN t_COMISSAO_INDICADOR_N2 ON (t_COMISSAO_INDICADOR_N1.id = t_COMISSAO_INDICADOR_N2.id_comissao_indicador_n1)" & _
            " INNER JOIN t_COMISSAO_INDICADOR_N3 ON (t_COMISSAO_INDICADOR_N2.id = t_COMISSAO_INDICADOR_N3.id_comissao_indicador_n2)" & _
            " INNER JOIN t_COMISSAO_INDICADOR_N4 ON (t_COMISSAO_INDICADOR_N3.id = t_COMISSAO_INDICADOR_N4.id_comissao_indicador_n3)" & _
            " INNER JOIN t_PEDIDO ON (t_COMISSAO_INDICADOR_N4.pedido = t_PEDIDO.pedido)" & _
            " INNER JOIN t_CLIENTE ON (t_PEDIDO.id_cliente=t_CLIENTE.id)" & _
            " WHERE (t_COMISSAO_INDICADOR_N2.vendedor = '" & s_vendedor & "')" & _
            " AND (t_COMISSAO_INDICADOR_N2.competencia_ano = '" & s_ano_competencia & "')" & _
            " AND (t_COMISSAO_INDICADOR_N2.competencia_mes = '" & s_mes_competencia & "')" &_                       
            " ORDER BY t_COMISSAO_INDICADOR_N2.vendedor," & _
            " t_COMISSAO_INDICADOR_N3.indicador," & _
            " t_COMISSAO_INDICADOR_N3.numero_banco," & _
            " t_COMISSAO_INDICADOR_N3.vl_total_comissao_arredondado"
    end if

    '	AS ROTINAS DE ORDENAÇÃO USAM VETORES QUE SE INICIAM NA POSIÇÃO 1
	redim vChequesPorBanco(1)
	for i = Lbound(vChequesPorBanco) to Ubound(vChequesPorBanco)
		set vChequesPorBanco(i) = New cl_DUAS_COLUNAS
		with vChequesPorBanco(i)
			.c1 = ""
			.c2 = 0
			end with
		next

    redim vDinheiroPorBanco(1)
	for i = Lbound(vDinheiroPorBanco) to Ubound(vDinheiroPorBanco)
		set vDinheiroPorBanco(i) = New cl_DUAS_COLUNAS
		with vDinheiroPorBanco(i)
			.c1 = ""
			.c2 = 0
			end with
		next
		
  ' CABEÇALHO
	cab_table = "<table cellspacing='0' id='tableDados'>" & chr(13)
	cab = "	<tr style='background:azure' nowrap>" & chr(13) & _
		  "		<td class='MDTE tdLoja' align='center' valign='bottom' nowrap><span class='Rc VISAO_ANALIT'>Loja</span></td>" & chr(13) & _
		  "		<td class='MTD tdOrcamento' align='left' valign='bottom' nowrap><span class='R VISAO_ANALIT'>Nº Orçam</span></td>" & chr(13) & _
		  "		<td class='MTD tdPedido' align='left' valign='bottom' nowrap><span class='R VISAO_ANALIT'>Nº Pedido</span></td>" & chr(13) & _
		  "		<td class='MTD tdData' align='center' valign='bottom' nowrap><span class='Rc VISAO_ANALIT'>Data</span></td>" & chr(13) & _
		  "		<td class='MTD tdVlPedido' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;'>VL Pedido</span></td>" & chr(13) & _
		  "		<td class='MTD tdVlRT' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;'>COM (" & SIMBOLO_MONETARIO & ")</span></td>" & chr(13) & _
		  "		<td class='MTD tdVlRABruto' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;'>RA Bruto (" & SIMBOLO_MONETARIO & ")</span></td>" & chr(13) & _
		  "		<td class='MTD tdVlRALiq' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;'>RA Líq (" & SIMBOLO_MONETARIO & ")</span></td>" & chr(13) & _
		  "		<td class='MTD tdVlRADif' align='right' valign='bottom'><span class='Rd' style='font-weight:bold;'>RA Dif (" & SIMBOLO_MONETARIO & ")</span></td>" & chr(13) & _
		  "		<td class='MTD tdStPagto' align='left' valign='bottom'><span class='R VISAO_ANALIT' style='font-weight:bold;'>St Pagto</span></td>" & chr(13) & _
		  "		<td class='MTD tdSinal' align='center' valign='bottom'><span class='Rc VISAO_ANALIT' style='font-weight:bold;'>+/-</span></td>" & chr(13) & _
		  "		<td valign='bottom' class='notPrint BkgWhite' align='left'>&nbsp;<a name='bExibeOcultaCampos' id='bExibeOcultaCampos' href='javascript:fExibeOcultaCampos(" & chr(34) & "_NNNNN_" & chr(34) & ");' title='exibe ou oculta os dados'><img src='../botao/view_bottom.png' border='0'></a></td>" & chr(13) & _
		  "	</tr>" & chr(13)
	
	x = ""
    st_tratamento_manual = 0
	n_reg = 0
	n_reg_total = 0
	idx_bloco = 0
    qtdeVendedor = 0
    totalComissaoArredondado = 0
	meio_pagto_a = "XXXXXXXXXXXX"
    indicador_a = "XXXXXXXXXX"
    vendedor_a = "XXXXXXXX"
    vendedor = ""

	set r = cn.execute(s_sql)

    if r.Eof then
        aviso="Não há indicadores com registros de descontos."
    end if

    if aviso = "" then
        c_dt_entregue_ano = r("competencia_ano")
        c_dt_entregue_mes = r("competencia_mes")
        redim vRelat(1)
	    for intIdxVetor = Lbound(vRelat) to Ubound(vRelat)
		    set vRelat(intIdxVetor) = New cl_VINTE_COLUNAS
		    vRelat(intIdxVetor).CampoOrdenacao = ""
		    next

        aviso=""
        if r("proc_automatico_status")=1 then
            aviso= "O Relatório já foi processado por " & r("proc_automatico_usuario") & " em " & r("proc_automatico_data_hora") & "."
        end if
        if aviso <> "" then
            x = "<div class='MtAlerta notPrint' style='width:649px;font-weight:bold;' align='center'><p style='margin:5px 2px 5px 2px;'>" & aviso & "</p></div><br />"
            Response.Write x
        end if 
        
        x = ""

	    do while Not r.Eof
    
	    '	MUDOU DE INDICADOR?
		    if Trim("" & r("ind3"))<>indicador_a then
			    indicador_a = Trim("" & r("ind3"))            
			    idx_bloco = idx_bloco + 1
			    qtde_indicadores = qtde_indicadores + 1
        
		      ' FECHA TABELA DO INDICADOR ANTERIOR
			    if n_reg_total > 0 then 
 
                   s_lista_total_comissao = s_lista_total_comissao & sub_total_comissao & ";"            
              ' SOMATORIA DO MEIOS DE PAGAMENTO CHECADOS
                if st_tratamento_manual <> 0 then 
                    s_checked=""                         
                else  
                    s_checked=" checked"
                    if meio_pagto = "CHQ" then                        
                        totalCHQ = totalCHQ + sub_total_com                            
                        
                        qtdeChq = qtdeChq +1
                    elseif meio_pagto = "DEP" Or meio_pagto = "DEP1" then                        
                        totalDEP = totalDEP + sub_total_com                                           
                    elseif meio_pagto = "DIN" then                        
                        totalDIN = totalDIN + sub_total_com
                        v_cedulas = Split(cedulas_codificado,",")
                        for cont=0 to Ubound(v_cedulas)
                            redim preserve qtdeCedula(cont)
                            qtdeCedula(cont) = cint(v_cedulas(cont))  
                        next 

                    end if
                end if                 
				    s_cor="black"
				    if vl_sub_total_preco_venda < 0 then s_cor="red"
				    if vl_sub_total_RT < 0 then s_cor="red"
				    if vl_sub_total_RA < 0 then s_cor="red"
				    if vl_sub_total_RA_liquido < 0 then s_cor="red"
				    x = x & "	<tr nowrap style='background: #FFFFDD'>" & chr(13) & _
						    "		<td class='MTBE' colspan='4' align='right' nowrap><span class='Cd' style='color:" & s_cor & ";'>" & _
						    "TOTAL:</span></td>" & chr(13) & _
						    "		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_preco_venda) & "</span></td>" & chr(13) & _                        
						    "		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_RT) & "</span></td>" & chr(13) & _                       
						    "		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_RA) & "</span></td>" & chr(13) & _                        
						    "		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_RA_liquido) & "</span></td>" & chr(13) & _                       
						    "		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_RA_diferenca) & "</span></td>" & chr(13) & _                      
						    "		<td class='MTBD' align='right' colspan='2'><span class='Cd' style='color:" & s_cor & ";'>&nbsp;</span></td>" & chr(13) & _
						    "		<td class='notPrint BkgWhite'>&nbsp;</td>" & chr(13) & _
						    "	</tr>" & chr(13) & _ 
		                    "   <tr>" & chr(13) 
                'SUB TOTAL COMISSÃO
                     if sub_total_comissao = 0 Or meio_pagto = "CHQ" Or (meio_pagto = "DEP" Or meio_pagto = "DEP1") then
                            x = x & "       <td align='left' colspan='5' nowrap>" 
                            if msg_desconto <> "" then x = x & "<span class='Cd' ><a href='javascript:abreDesconto(" & idx_bloco -1 & ")' title='Exibe ou oculta os registros de descontos' style='color: red;'>"& msg_desconto & "</a></span>"& chr(13)
                               if vl_sub_total_RT_arredondado >=0 then 
                                  s_cor = "black "
                               end if 
                               if vl_sub_total_RT <0 or vl_sub_total_RA_liquido <0 then
                                   x = x &             "<td align='right' nowrap><span class='Rd' style='color:" & s_cor & ";'>"& formata_moeda(vl_sub_total_RT)&"</span>" & chr(13) 
                               else
                                   x = x &             "<td align='right' nowrap><span class='Rd' style='color:" & s_cor & ";'>"& formata_moeda(vl_sub_total_RT_arredondado)&"</span>" & chr(13) 
                               end if
                               if vl_sub_total_RA_arredondado >=0 then 
                                  s_cor = "black "
                               else
                                  s_cor = "red"
                               end if
                               if vl_sub_total_RT <0 or vl_sub_total_RA <0 then
                                    x = x &             "            <td align='right' colspan='2' nowrap><span class='Rd' style='color:" & s_cor & ";'>"& formata_moeda(vl_sub_total_RA_liquido)&"</span>" & chr(13)
                               else        
                                    x = x &             "            <td align='right' colspan='2' nowrap><span class='Rd' style='color:" & s_cor & ";'>"& formata_moeda(vl_sub_total_RA_arredondado)&"</span>" & chr(13)
                               end if              
                     else if sub_total_comissao >=0 then
                            x = x & "       <td align='left' colspan='5' nowrap>"
                            if msg_desconto <> "" then 
                               x = x & "<span class='Cd'><a href='javascript:abreDesconto(" & idx_bloco -1 & ")' title='Exibe ou oculta os registros de descontos' style='color: red;'>"& msg_desconto & "</a></span>"& chr(13)
                               elseif st_tratamento_manual = 0 and meio_pagto = "DIN" then  
                                    x = x & "<span class='Rd' style='color: black;'> Cédulas: "                  
                                     for cont = 0 to UBound(qtdeCedula)
                                            if (cont = 0 And qtdeCedula(cont) <> 0) then                 
                                                 if (qtdeCedula(cont) > 1) then x = x & qtdeCedula(cont) & "&times;"
                                                        x = x & formata_moeda("100") 
                                                        aux(cont) =  aux(cont) + qtdeCedula(cont)
                                                        if (qtdeCedula(1) <> 0 Or qtdeCedula(2) <> 0 Or qtdeCedula(3) <> 0 Or qtdeCedula(4) <> 0 Or qtdeCedula(5) <> 0) then x = x & " + "
                          
                                            elseif (cont = 1 And qtdeCedula(cont) <> 0) then
                                                 if (qtdeCedula(cont) > 1) then x = x & qtdeCedula(cont) & "&times;"
                                                        x = x & formata_moeda("50") 
                                                        aux(cont) =  aux(cont) + qtdeCedula(cont)
                                                        if (qtdeCedula(2) <> 0 Or qtdeCedula(3) <> 0 Or qtdeCedula(4) <> 0 Or qtdeCedula(5) <> 0) then x = x & " + "
                                       
                           
                                            elseif (cont = 2 And qtdeCedula(cont) <> 0) then
                                                 if (qtdeCedula(cont) > 1) then x = x & qtdeCedula(cont) & "&times;"
                                                        x = x & formata_moeda("20") 
                                                        aux(cont) =  aux(cont) + qtdeCedula(cont)
                                                        if (qtdeCedula(3) <> 0 Or qtdeCedula(4) <> 0 Or qtdeCedula(5) <> 0) then x = x & " + "
                        
                                            elseif (cont = 3 And qtdeCedula(cont) <> 0) then
                                                if (qtdeCedula(cont) > 1) then x = x & qtdeCedula(cont) & "&times;"
                                                    x = x & formata_moeda("10") 
                                                    aux(cont) =  aux(cont) + qtdeCedula(cont)
                                                    if (qtdeCedula(4) <> 0 Or qtdeCedula(5) <> 0) then x = x & " + "

                                            elseif (cont = 4 And qtdeCedula(cont) <> 0) then
                                                if (qtdeCedula(cont) > 1) then x = x & qtdeCedula(cont) & "&times;"
                                                    x = x & formata_moeda("5") 
                                                    aux(cont) =  aux(cont) + qtdeCedula(cont)
                                                    if (qtdeCedula(5) <> 0) then x = x & " + "
                       
                                            elseif (cont = 5 And qtdeCedula(cont) <> 0) then
                                                if (qtdeCedula(cont) > 1) then x = x & qtdeCedula(cont) & "&times;"
                                                    x = x & formata_moeda("2") 
                                                    aux(cont) =  aux(cont) + qtdeCedula(cont)
                                            end if
                                     next                                                                                    
                                 end if
                                                   
                                if vl_sub_total_RT_arredondado >=0 then 
                                    s_cor = "black "
                                end if                               
                                    x = x & " <td align='right' nowrap><span class='Rd' style='color:" & s_cor & ";margin-right:0.3;'>"& formata_moeda(vl_sub_total_RT_arredondado)&"</span>"& chr(13) 
                                if vl_sub_total_RA_arredondado >=0 then 
                                    s_cor = "black "
                                else
                                    s_cor = "red"
                                end if                               
                                    x = x &  " <td align='right' colspan='2' nowrap><span class='Rd' style='color:" & s_cor & ";'>"& formata_moeda(vl_sub_total_RA_arredondado)& "</span>"& chr(13) 
                                     
                                end if
                                if sub_total_com >= 0 then 
                                    s_cor = "black"
                                else 
                                    s_cor = "red"
                                end if
                              end if
                                    x = x & "       <td align='right' colspan='2' nowrap><span class='Cd' style='color:" & s_cor & ";'>"
                    if sub_total_comissao <= 0 then
                        x = x & "&nbsp;</span></td>" & chr(13)
                    else 
                        x = x & ""& regex.Replace(meio_pagto,"") &":</span></td>" & chr(13) 
                    end if
                
                    if sub_total_comissao >= 0 then
                        x = x & "<td align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(sub_total_com) & "</span></td>" & chr(13)
                    end if
                      
                    x = x & "</tr>" & chr(13) 
                    if msg_desconto <> "" then
                        x = x &"   <tr>" & chr(13)& _
                               "          <td  class='table_Desconto' id='table_Desconto_"& idx_bloco -1  &"'"" colspan='15' >" & chr(13)& _
                               "          <table colspan='2' align='left' >"& chr(13)
                        for contador = 0 to Ubound(v_desconto_descricao)                                 
                            x = x & "<tr>" & chr(13)& _
                                "       <td width='15'>&nbsp;</td>" & chr(13)& _
                                "       <td  align='left' width='400' ><span class='Cd'style='color: red;' >"& v_desconto_descricao(contador)& "</span></td>"& _
                                "       <td align='left' ><span class='Cd'style='color: red;' > R$ "& formata_moeda(v_desconto_valor(contador))& "</span></td>"& _
                                "   </tr>"                                     
                        next          
                        x = x & "        </table>"& chr(13)& _
                                "        </td>"& chr(13)& _
                                "</tr>"
                    end if
				    x = x & "</table>" & chr(13)

                     x = "   <table cellpadding='0' cellspacing='0'><tr><td valign='top'><br />" & chr(13) & _
                         "      <input type='checkbox' name='ckb_comissao_paga_tit_bloco' class='CKB_COM' id='ckb_comissao_paga_tit_bloco_" & idx_bloco -1 & "' onclick='trata_ckb_onclick();calculaTotalComissao();alternaCheck(" & idx_bloco -1 & ");' value='" & atual & "' " & s_checked & s_disabled & " />" & chr(13) & _ 
                         "      <input type='checkbox' style='display:none' name='ckb_comissao_paga_tit_bloco_indicador' id='ckb_comissao_paga_tit_bloco_indicador_" & idx_bloco -1 & "' value='" & ind_anterior & "' />" & chr(13) & _
                         "   </td><td valign='top'>" & x & "</td></tr></table>" & chr(13)
                     s_lista_total_comissao_arredondado = s_lista_total_comissao_arredondado & sub_total_comissao & ";"
             
                   atual = ""
                   s_checked= ""
			       Response.Write x
			       x="<BR>" & chr(13)
			       end if
               
                s_sql = "SELECT * FROM t_COMISSAO_INDICADOR_N3 WHERE (indicador = '" & Trim("" & r("ind3")) & "') AND (id_comissao_indicador_n2 = " & Trim("" & r("id_n2")) & ")" 
			    if rs.State <> 0 then rs.Close
                
             
           
			    rs.Open s_sql, cn
			    if Not rs.Eof then
				    s_banco = Trim("" & rs("banco"))
				    s_agencia = Trim("" & rs("agencia"))
				    s_conta = Trim("" & rs("conta"))
				    s_favorecido = Trim("" & rs("favorecido"))
				    s_banco_nome = x_banco(s_banco)
				    if (s_banco <> "") And (s_banco_nome <> "") then s_banco = s_banco & " - " & s_banco_nome
			    else
				    s_banco = ""
				    s_banco_nome = ""
				    s_agencia = ""
				    s_conta = ""
				    s_favorecido = ""
			    end if

                if rs("st_tratamento_manual") = 0 then
                    if Trim("" & rs("meio_pagto")) = "CHQ" then
                        s = converte_numero(Trim("" & rs("banco")))
                        s = CStr(s)
		                if localiza_cl_duas_colunas(vChequesPorBanco, s, idx) then
			                with vChequesPorBanco(idx)
				                .c2 = .c2 + 1
				                end with
		                else
			                if (vChequesPorBanco(Ubound(vChequesPorBanco)).c1<>"") then
				                redim preserve vChequesPorBanco(Ubound(vChequesPorBanco)+1)
				                set vChequesPorBanco(Ubound(vChequesPorBanco)) = New cl_DUAS_COLUNAS
				                end if
			                with vChequesPorBanco(Ubound(vChequesPorBanco))
				                .c1 = s
				                .c2 = 1
				        end with
			                ordena_cl_duas_colunas vChequesPorBanco, 1, Ubound(vChequesPorBanco)
			                end if
                    elseif Trim("" & rs("meio_pagto")) = "DIN" then
                        s = converte_numero(Trim("" & rs("banco")))
                        s = CStr(s)
		                if localiza_cl_duas_colunas(vDinheiroPorBanco, s, idx) then
			                with vDinheiroPorBanco(idx)
				                .c2 = .c2 + 1
				                end with
		                else
			                if (vDinheiroPorBanco(Ubound(vDinheiroPorBanco)).c1<>"") then
				                redim preserve vDinheiroPorBanco(Ubound(vDinheiroPorBanco)+1)
				                set vDinheiroPorBanco(Ubound(vDinheiroPorBanco)) = New cl_DUAS_COLUNAS
				                end if
			                with vDinheiroPorBanco(Ubound(vDinheiroPorBanco))
				                .c1 = s
				                .c2 = 1
				                end with
			                ordena_cl_duas_colunas vDinheiroPorBanco, 1, Ubound(vDinheiroPorBanco)
			            end if
                    end if
                end if

                x = x & Replace(cab_table, "tableDados", "tableDados_" & idx_bloco)
                x = x & "	<tr>" & chr(13)
            
                s = Trim("" & r("ind3"))
			    s_aux = x_orcamentista_e_indicador(s)
			    if (s<>"") And (s_aux<>"") then s = s & " - "
			    s = s & s_aux
			
			    if s <> "" then x = x & "		<td class='MDTE' colspan='11' align='left' valign='bottom' class='MB' style='background:azure;'><span class='N'>&nbsp;" & s_desempenho_nota & s & "</span></td>" & chr(13) & _
									    "		<td class='notPrint BkgWhite'>&nbsp;</td>" & chr(13) & _
									    "	</tr>" & chr(13) & _
									    "	<tr>" & chr(13) & _
									    "		<td class='MDTE' colspan='11' align='left' valign='bottom' class='MB' style='background:whitesmoke;'>" & chr(13) & _
									    "			<table width='100%' cellspacing='0' cellpadding='0'>" & chr(13) & _
									    "				<tr>" & chr(13) & _
									    "					<td colspan='3' align='left' valign='bottom' style='vertical-align:middle'><div valign='bottom' style='height:14px;max-height:14px;overflow:hidden;vertical-align:middle'><span class='Cn'>Banco: " & rs("banco") & " - " & x_banco(rs("banco")) &  "</span></div></td>" & chr(13) & _
									    "				</tr>" & chr(13) & _
									    "				<tr>" & chr(13) & _
									    "					<td class='MTD' align='left' valign='bottom' style='height:15px;vertical-align:middle'><span class='Cn'>Agência: " & rs("agencia")
                            if Trim("" & rs("agencia_dv")) <> "" then
                                x = x & "-" & rs("agencia_dv") & chr(13)
                            end if
    
                            x = x & "</span></td>" & chr(13) & _
									    "					<td class='MC MD' align='left' valign='bottom' style='height:15px;vertical-align:middle'><span class='Cn'>"

                            if Trim("" & rs("tipo_conta")) <> "" then
                                if rs("tipo_conta") = "P" then
                                    x = x & "C/P: "
                                elseif rs("tipo_conta") = "C" then
                                    x = x & "C/C: "
                                end if
                            else
                                x = x & "Conta: "
                            end if

                            if Trim("" & rs("conta_operacao")) <> "" then
                                x = x & rs("conta_operacao") & "-"
                            end if               
    
                            x = x & rs("conta")
    
                            if Trim("" & rs("conta_dv")) <> "" then
                                x = x & "-" & rs("conta_dv") & chr(13)
                            end if
							    x =  x &"					<td class='MC' width='60%' align='left' valign='bottom'><span class='Cn'>Favorecido: " & s_favorecido & "</span></td>" & chr(13) & _
									    "				</tr>" & chr(13) & _
									    "			</table>" & chr(13) & _
									    "		</td>" & chr(13) & _
									    "		<td class='notPrint BkgWhite'>&nbsp;</td>" & chr(13) & _
									    "	</tr>" & chr(13)
			        s_new_cab = Replace(cab, "ckb_comissao_paga_tit_bloco", "ckb_comissao_paga_tit_bloco_" & idx_bloco)
			        s_new_cab = Replace(s_new_cab, "trata_ckb_onclick();", "trata_ckb_onclick(" & chr(34) & idx_bloco & chr(34) & ");")
			        s_new_cab = Replace(s_new_cab, "_NNNNN_", CStr(idx_bloco))
			        x = x & s_new_cab

			        n_reg = 0
			        vl_sub_total_preco_venda = 0
			        vl_sub_total_preco_NF = 0
			        vl_sub_total_RT = 0
			        vl_sub_total_RA = 0
			        vl_sub_total_RA_liquido = 0
			        vl_sub_total_RA_diferenca = 0
                    vl_total_RT_arredondado = 0
                    vl_sub_total_RA_arredondado = 0
                    vl_sub_total_comissao = 0
                    meio_pagto = r("meio_pagto")    
                    st_tratamento_manual = r("st_tratamento_manual")
                    cedulas_codificado=r("cedulas_codificado")
	            end if
   


                if atual <> "" then atual = atual & ", "
                atual = atual & Trim("" & r("pedido"))

	     ' CONTAGEM
		        n_reg = n_reg + 1
		        n_reg_total = n_reg_total + 1

        ' CÁLCULOS

        '	EVITA DIFERENÇAS DE ARREDONDAMENTO
		        vl_preco_venda = converte_numero(formata_moeda(r("vl_pedido")))
		        perc_RT = r("perc_RT")
		        vl_RT = r("vl_comissao")
		        vl_RA = r("vl_RA_bruto")
		        vl_comissao = ("vl_comissao_arredondado")
		        vl_RA_liquido = r("vl_RA_liq")
		        vl_RA_diferenca = vl_RA - vl_RA_liquido

        ' CÁLCULOS DE SUB TOTAL
                vl_sub_total_preco_venda = vl_sub_total_preco_venda + r("vl_pedido")
                if st_tratamento_manual = 0 then vl_total_preco_venda = vl_total_preco_venda + r("vl_pedido")		
		        vl_sub_total_RT = vl_sub_total_RT +  vl_RT
                vl_sub_total_RT_arredondado = r("vl_total_comissao_arredondado")
		        if st_tratamento_manual = 0 then vl_total_RT = vl_total_RT + vl_RT
                vl_total_RT_arredondado = vl_total_RT_arredondado + vl_sub_total_RT_arredondado
		        vl_sub_total_RA = vl_sub_total_RA +  vl_RA
                if st_tratamento_manual = 0 then vl_total_RA = vl_total_RA + vl_RA
		        vl_sub_total_RA_liquido = vl_sub_total_RA_liquido + vl_RA_liquido
		        if st_tratamento_manual = 0 then vl_total_RA_liquido = vl_total_RA_liquido + vl_RA_liquido
                vl_sub_total_RA_arredondado = r("vl_total_RA_arredondado")
		        vl_sub_total_RA_diferenca = vl_sub_total_RA_diferenca + vl_RA_diferenca
		        if st_tratamento_manual = 0 then vl_total_RA_diferenca = vl_total_RA_diferenca + vl_RA_diferenca
                sub_total_comissao = vl_sub_total_RT + vl_sub_total_RA_liquido
                vl_sub_total_comissao = vl_sub_total_RA_arredondado + vl_sub_total_RT_arredondado

                sub_total_com = r("vl_total_pagto")

                if s_lista_completa_pedidos <> "" then s_lista_completa_pedidos = s_lista_completa_pedidos & ";"
			        s_lista_completa_pedidos = s_lista_completa_pedidos & Trim("" & r("pedido"))
                
                rs2.Open "SELECT descricao,valor,ordenacao  FROM t_COMISSAO_INDICADOR_N3_DESCONTO" & _
                " INNER JOIN t_COMISSAO_INDICADOR_N3 ON (t_COMISSAO_INDICADOR_N3.id = t_COMISSAO_INDICADOR_N3_DESCONTO.id_comissao_indicador_n3)" & _
                " WHERE (id_comissao_indicador_n3 = '" & r("id_n3") & "') GROUP BY  descricao,valor,ordenacao ORDER BY ordenacao", cn
                
                if r("qtde_reg_descontos_planilha") > 0 then
                    contador = 0
                    qtde_registro_desc = 0
                    valor_desconto = 0
                    Erase v_desconto_descricao
                    Erase v_desconto_valor
                    do while Not rs2.EoF         
                        redim preserve v_desconto_descricao(contador) 
                        redim preserve  v_desconto_valor(contador)                 
                        v_desconto_descricao(contador) = rs2("descricao")
                        v_desconto_valor(contador) = rs2("valor")
                        qtde_registro_desc = qtde_registro_desc + 1
                        valor_desconto = valor_desconto + v_desconto_valor(contador)                        
                        contador = contador + 1
                        rs2.MoveNext
		            loop    
                    msg_desconto =  "Planilha de Desconto:&nbsp; " & Cstr(qtde_registro_desc ) & " registro(s) no valor total de&nbsp;" & formata_moeda(valor_desconto)                
                else
                    msg_desconto= ""
                end if
                if rs2.State <> 0 then rs2.Close
         '> CHECK BOX
	     '	É USADO O CÓDIGO DA OPERAÇÃO (VENDA NORMAL, DEVOLUÇÃO, PERDA) P/ NÃO CORRER O RISCO DE HAVER CONFLITO DEVIDO A ID'S REPETIDOS ENTRE AS OPERAÇÕES
		        s_class = " CKB_COM_BL_" & idx_bloco
		        s_class_td = ""

		        x = x & "	<tr nowrap class='VISAO_ANALIT'>"  & chr(13)
		
		        if (vl_preco_venda < 0) Or (vl_RT < 0) Or (vl_RA < 0) Or (vl_RA_liquido < 0) then
			        s_cor = "red"
			        s_cor_sinal = "red"
			        s_sinal = "-"
		        else
			        s_cor = "black"
			        s_cor_sinal = "green"
			        s_sinal = "+"
			    end if

            '	x = x & "		<input type='hidden' class='CKB_COM " & s_class & "' name='" & s_id & "' id='" & s_id & "' value='" & Trim("" & r("id_registro")) & "|" & Trim("" & r("operacao")) & "' />" & chr(13)
		
	         '> LOJA
		        x = x & "		<td class='MDTE tdLoja' align='center'><span class='Cnc' style='color:" & s_cor & ";'>"& Trim("" & r("loja")) &"</span></td>" & chr(13)

	         '> Nº ORÇAMENTO
		        s = Trim("")
		        if s = "" then s = "&nbsp;"
		        x = x & "		<td class='MTD tdOrcamento' align='left'><span class='Cn'><a style='color:" & s_cor & ";' href='javascript:fORCConsulta(" & _
				        chr(34) & s & chr(34) & "," & chr(34) & usuario & chr(34) & ")' title='clique para consultar o orçamento'>" & _
				        s & "</a></span></td>" & chr(13)

	         '> Nº PEDIDO
                if blnActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos then
		            s_nome_cliente = Trim("" & r("endereco_nome_iniciais_em_maiusculas"))
                else
		            s_nome_cliente = Trim("" & r("nome_iniciais_em_maiusculas"))
                    end if

		        s_nome_cliente = Left(s_nome_cliente, 15)
		
		        x = x & "		<td class='MTD tdPedido' align='left'><span class='Cn'><a style='color:" & s_cor & ";' href='javascript:fPEDConsulta(" & _
				        chr(34) & Trim("" & r("pedido")) & chr(34) & "," & chr(34) & usuario & chr(34) & ")' title='clique para consultar o pedido'>" & _
				        Trim("" & r("pedido")) & "<br>" & s_nome_cliente & "</a></span></td>" & chr(13)

	         '> DATA
		        s = formata_data(r("data"))
		        x = x & "		<td align='center' class='MTD tdData'><span class='Cnc' style='color:" & s_cor & ";'>" & s & "</span></td>" & chr(13)

	         '> VALOR DO PEDIDO (PREÇO DE VENDA)
		        x = x & "		<td align='right' class='MTD tdVlPedido'><span class='Cnd' style='color:" & s_cor & ";'>" & formata_moeda(r("vl_pedido")) & "</span></td>" & chr(13)
                s_lista_vl_pedido = s_lista_vl_pedido & vl_preco_venda & ";"

	         '> COMISSÃO (ANTERIORMENTE CHAMADO DE RT)
		        x = x & "		<td align='right' class='MTD tdVlRT'><span class='Cnd' style='color:" & s_cor & ";'>" & formata_moeda(r("vl_comissao")) & "</span></td>" & chr(13)
                s_lista_comissao = s_lista_comissao & vl_RT & ";"

	         '> RA BRUTO
		        x = x & "		<td align='right' class='MTD tdVlRABruto'><span class='Cnd' style='color:" & s_cor & ";'>" & formata_moeda(r("vl_RA_bruto")) & "</span></td>" & chr(13)
                s_lista_RA_bruto = s_lista_RA_bruto & formata_moeda(vl_RA) & ";"

	         '> RA LÍQUIDO
		        x = x & "		<td align='right' class='MTD tdVlRALiq'><span class='Cnd' style='color:" & s_cor & ";'>" & formata_moeda(r("vl_RA_liq")) & "</span></td>" & chr(13)
                s_lista_RA_liquido = s_lista_RA_liquido & formata_moeda(vl_RA_liquido) & ";"

	         '> RA DIFERENÇA
		        x = x & "		<td align='right' class='MTD tdVlRADif'><span class='Cnd' style='color:" & s_cor & ";'>" & formata_moeda(vl_RA_diferenca) & "</span></td>" & chr(13)

	         '> STATUS DE PAGAMENTO
		        x = x & "		<td class='MTD tdStPagto' align='left'><span class='Cn' style='color:" & s_cor & ";'>" & x_status_pagto(Trim("" & r("st_pagto"))) & "</span></td>" & chr(13)

	         '> +/-
		        x = x & "		<td align='center' class='MTD tdSinal'><span class='C' style='font-family:Courier,Arial;color:" & s_cor_sinal & "'>" & s_sinal & "</span></td>" & chr(13)
		
	         '> COLUNA DA FIGURA (EXPANDE/RECOLHE)
		        x = x & "		<td class='notPrint BkgWhite'>&nbsp;</td>" & chr(13)
		
		        x = x & "	</tr>" & chr(13)

			    ind_anterior = r("ind3")
			    banco = r("banco")
        	    meio_pagto = r("meio_pagto")

           
               if vendedor = "" then
                   vendedor = vendedor & r("vendedor")
               else
                   if r("vendedor") <> vendedor_a then 
                     vendedor = vendedor & "," & r("vendedor")
                   end if
               end if
               vendedor_a = r("vendedor")
            
            
		    r.MoveNext
        loop
	
	      ' MOSTRA TOTAL DO ÚLTIMO INDICADOR
	    if n_reg <> 0 then 
		    s_cor="black"
		    if vl_sub_total_preco_venda < 0 then s_cor="red"
		    if vl_sub_total_RT < 0 then s_cor="red"
		    if vl_sub_total_RA < 0 then s_cor="red"
		    if vl_sub_total_RA_liquido < 0 then s_cor="red"
            sub_total_comissao = vl_sub_total_RT + vl_sub_total_RA_liquido
        ' 
       
            
             
        
            if st_tratamento_manual = 0 and meio_pagto = "DIN" then
                v_cedulas = Split(cedulas_codificado,",")
                for cont=0 to Ubound(v_cedulas)
                    redim preserve qtdeCedula(cont)
                    qtdeCedula(cont) = cint(v_cedulas(cont))              
                next 
            end if
		    x = x & "	<tr nowrap style='background: #FFFFDD'>" & chr(13) & _
				    "		<td colspan='4' class='MTBE' align='right' nowrap><span class='Cd' style='color:" & s_cor & ";'>" & _
				    "                   TOTAL:</span></td>" & chr(13) & _
				    "		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_preco_venda) & "</span></td>" & chr(13) & _                      
				    "		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_RT) & "</span></td>" & chr(13) & _                                          
				    "		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_RA) & "</span></td>" & chr(13) & _                       
				    "		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_RA_liquido) & "</span></td>" & chr(13) & _                        
				    "		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_RA_diferenca) & "</span></td>" & chr(13) & _                       
				    "		<td class='MTBD' align='right' colspan='2'><span class='Cd' style='color:" & s_cor & ";'>&nbsp;</span></td>" & chr(13) & _
				    "		<td class='notPrint BkgWhite'>&nbsp;</td>" & chr(13) & _
				    "	</tr>" & chr(13) & _
				    "	</tr>" & chr(13) &_
	                "   <tr>" & chr(13) 
      '  VERIFICA A FORMA DE PAGAMENTO,SE CONTEM DESCONTO E MOSTRA AS CEDULAS                
                    if sub_total_comissao <= 0 Or sub_total_comissao > 300 Or banco = "237" then
                        x = x & "       <td align='left' colspan='5' nowrap>" 
                        if msg_desconto <> "" then x = x & "<span class='Cd'><a href='javascript:abreDesconto(" & idx_bloco  & ")' title='Exibe ou oculta os registros de descontos' style='color: red;'>"& msg_desconto & "</a></span>"
                             if vl_sub_total_RT_arredondado >=0 then 
                                  s_cor = "black "
                               end if   
                               x = x &  "<td align='right' nowrap><span class='Rd' style='color:" & s_cor & ";'>"& formata_moeda(vl_sub_total_RT_arredondado)&"</span>" & chr(13) 
                               if vl_sub_total_RA_arredondado >=0 then 
                                  s_cor = "black "
                               else
                                  s_cor = "red"
                               end if       
                               x = x &  "            <td align='right' colspan='2' nowrap><span class='Rd' style='color:" & s_cor & ";'>"& formata_moeda(vl_sub_total_RA_arredondado)&"</span>" & chr(13)
                                         
                        else if sub_total_comissao >=0 then
                             x = x & "       <td align='left' colspan='5' nowrap>"
                             if msg_desconto <> "" then 
                                 x = x & "<span class='Cd' ><a href='javascript:abreDesconto(" & idx_bloco  & ")' title='Exibe ou oculta os registros de descontos' style='color: red;'>"& msg_desconto & "</a></span>"
                             elseif st_tratamento_manual = 0 and meio_pagto = "DIN" then  
                                 x = x & "<span class='Rd' style='color: black;'> Cédulas: "                  
                             for cont = 0 to UBound(qtdeCedula)
                                 if (cont = 0 And qtdeCedula(cont) <> 0) then                 
                                      if (qtdeCedula(cont) > 1) then x = x & qtdeCedula(cont) & "&times;"
                                        x = x & formata_moeda("100") 
                                        aux(cont) =  aux(cont) + qtdeCedula(cont)
                                        if (qtdeCedula(1) <> 0 Or qtdeCedula(2) <> 0 Or qtdeCedula(3) <> 0 Or qtdeCedula(4) <> 0 Or qtdeCedula(5) <> 0) then x = x & " + "
                          
                                 elseif (cont = 1 And qtdeCedula(cont) <> 0) then
                                      if (qtdeCedula(cont) > 1) then x = x & qtdeCedula(cont) & "&times;"
                                        x = x & formata_moeda("50") 
                                        aux(cont) =  aux(cont) + qtdeCedula(cont)
                                        if (qtdeCedula(2) <> 0 Or qtdeCedula(3) <> 0 Or qtdeCedula(4) <> 0 Or qtdeCedula(5) <> 0) then x = x & " + "
                             
                                elseif (cont = 2 And qtdeCedula(cont) <> 0) then
                                      if (qtdeCedula(cont) > 1) then x = x & qtdeCedula(cont) & "&times;"
                                        x = x & formata_moeda("20") 
                                        aux(cont) =  aux(cont) + qtdeCedula(cont)
                                        if (qtdeCedula(3) <> 0 Or qtdeCedula(4) <> 0 Or qtdeCedula(5) <> 0) then x = x & " + "
                        
                                elseif (cont = 3 And qtdeCedula(cont) <> 0) then
                                      if (qtdeCedula(cont) > 1) then x = x & qtdeCedula(cont) & "&times;"
                                        x = x & formata_moeda("10") 
                                        aux(cont) =  aux(cont) + qtdeCedula(cont)
                                        if (qtdeCedula(4) <> 0 Or qtdeCedula(5) <> 0) then x = x & " + "

                                elseif (cont = 4 And qtdeCedula(cont) <> 0) then
                                      if (qtdeCedula(cont) > 1) then x = x & qtdeCedula(cont) & "&times;"
                                        x = x & formata_moeda("5") 
                                        aux(cont) =  aux(cont) + qtdeCedula(cont)
                                        if (qtdeCedula(5) <> 0) then x = x & " + "
                       
                                elseif (cont = 5 And qtdeCedula(cont) <> 0) then
                                      if (qtdeCedula(cont) > 1) then x = x & qtdeCedula(cont) & "&times;"
                                        x = x & formata_moeda("2") 
                                        aux(cont) =  aux(cont) + qtdeCedula(cont)
                                end if                 
                             next                                                                                      
                        end if
                        if vl_sub_total_RT_arredondado >=0 then 
                            s_cor = "black "
                        end if                               
                            x = x & " <td align='right' nowrap><span class='Rd' style='color:" & s_cor & ";margin-right:0.3;'>"& formata_moeda(vl_sub_total_RT_arredondado)&"</span>"& chr(13) 
                        if vl_sub_total_RA_arredondado >=0 then 
                            s_cor = "black "
                        else
                            s_cor = "red"
                        end if                               
                            x = x &  " <td align='right' colspan='2' nowrap><span class='Rd' style='color:" & s_cor & ";'>"& formata_moeda(vl_sub_total_RA_arredondado)& "</span>"& chr(13) 
                                     
                        end if
                        if sub_total_com >= 0 then 
                            s_cor = "black"
                        else 
                            s_cor = "red"
                        end if
                     end if
                                    x = x & "       <td align='right' colspan='2' nowrap><span class='Cd' style='color:" & s_cor & ";'>"
            if sub_total_comissao <= 0 then
               x = x & "&nbsp;</span></td>" & chr(13)
            else 
               x = x & ""& regex.Replace(meio_pagto,"") & ":</span></td>" & chr(13) 
            end if
        
            if sub_total_comissao >=0 then
            x = x & "       <td align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(sub_total_com) & "</span></td>" & chr(13)
            end if
            if msg_desconto <> "" then
                x = x &"   <tr>" & chr(13)& _
                        "          <td  class='table_Desconto' id='table_Desconto_"& idx_bloco &"'"" colspan='15' >" & chr(13)& _
                        "          <table colspan='2' align='left' >"& chr(13)
                        for contador = 0 to Ubound(v_desconto_descricao)                                 
                            x = x & "   <tr>" & chr(13)& _
                                    "       <td width='15'>&nbsp;</td>" & chr(13)& _
                                    "       <td  align='left' width='400' ><span class='Cd'style='color: red;' >"& v_desconto_descricao(contador)& "</span></td>"& _
                                    "       <td align='left' ><span class='Cd'style='color: red;' > R$ "& formata_moeda(v_desconto_valor(contador))& "</span></td>"& _
                                    "   </tr>"                                     
                        next          
                        x = x & "   </table>"& chr(13)& _
                                "   </td>"& chr(13)& _
                                "</tr>"
             end if
	    '>	TOTAL GERAL
		    if qtde_indicadores >= 1 then
			    s_cor="black"
			    if vl_total_preco_venda < 0 then s_cor="red"
			    if vl_total_RT < 0 then s_cor="red"
			    if vl_total_RA < 0 then s_cor="red"
			    if vl_total_RA_liquido < 0 then s_cor="red"
			    x = x & "	<tr>" & chr(13) & _
					    "		<td colspan='12' style='border-left:0px;border-right:0px;' align='left'>&nbsp;</td>" & chr(13) & _
					    "		<td class='notPrint BkgWhite'>&nbsp;</td>" & chr(13) & _
					    "	</tr>" & chr(13) & _
					    "	<tr>" & chr(13) & _
					    "		<td colspan='12' style='border-left:0px;border-right:0px;' align='left'>&nbsp;</td>" & chr(13) & _
					    "		<td class='notPrint BkgWhite'>&nbsp;</td>" & chr(13) & _
					    "	</tr>" & chr(13) & _
					    "	<tr nowrap style='background:honeydew'>" & chr(13) & _
					    "		<td class='MTBE' colspan='4' align='right' nowrap><span class='Cd' style='color:" & s_cor & ";'>" & _
					    "TOTAL GERAL:</span></td>" & chr(13) & _
					    "		<td class='MTB' align='right'><span id ='total_VlPedido'  class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_preco_venda) & "</span></td>" & chr(13) & _
					    "		<td class='MTB' align='right'><span id='totalComissao' class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_RT) & "</span></td>" & chr(13) & _
					    "		<td class='MTB' align='right'><span id='total_RA' class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_RA) & "</span></td>" & chr(13) & _
					    "		<td class='MTB' align='right'><span id='total_RAliq' class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_RA_liquido) & "</span></td>" & chr(13) & _
					    "		<td class='MTB' align='right'><span id='total_RAdif' class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_RA_diferenca) & "</span></td>" & chr(13) & _
					    "		<td class='MTBD' align='right' colspan='2'><span class='Cd' style='color:" & s_cor & ";'>&nbsp;</span></td>" & chr(13) & _
					    "		<td class='notPrint BkgWhite'>&nbsp;</td>" & chr(13) & _
					    "	</tr>" & chr(13) &_
                        " </table>" & chr(13)
         ' SELECIONAR NO CHECK BOX E CALCULO DO TOTAL DE CHQ,DEP E DIN                   
                   if st_tratamento_manual <> 0 then 
                        s_checked=""                
                   else            
                        s_checked=" checked"
                        if meio_pagto = "CHQ" then                        
                            totalCHQ = totalCHQ + sub_total_com
                            qtdeChq = qtdeChq +1
                        elseif meio_pagto = "DEP" Or meio_pagto = "DEP1" then                        
                            totalDEP = totalDEP + sub_total_com                                                
                        elseif meio_pagto = "DIN" then                        
                            totalDIN = totalDIN + sub_total_com
                        else
                            totalDIN = totalDIN + sub_total_com                           
                         end if
                     
                    end if                            
                    totalComissaoArredondado = totalCHQ +  totalDEP + totalDIN

                    x = "<table cellpadding='0' cellspacing='0'><tr><td valign='top'><br />" & chr(13) & _
                         "   <input type='checkbox' name='ckb_comissao_paga_tit_bloco' class='CKB_COM' id='ckb_comissao_paga_tit_bloco_" & idx_bloco & "' onclick='trata_ckb_onclick();calculaTotalComissao();alternaCheck(" & idx_bloco & ");' value='" & atual & "' " & s_checked & s_disabled & " />" & chr(13) & _ 
                         "   <input type='checkbox'  style='display:none' name='ckb_comissao_paga_tit_bloco_indicador' id='ckb_comissao_paga_tit_bloco_indicador_" & idx_bloco & "' value='" & ind_anterior & "' />" & chr(13) & _
                         "</td><td valign='top'>" & x & "</td></tr></table>" & chr(13)
       '    TOTAL DE CEDULAS
                    if aux(0) <> 0 then                 
                        textDin = textDin + Cstr(aux(0)) + "&times;100,00 "
                        if ((aux(1) <> 0) OR (aux(2) <> 0) OR (aux(3) <> 0) OR (aux(4) <> 0) OR (aux(5) <> 0)) then textDin = textDin + "&nbsp;&nbsp;&nbsp;+&nbsp;&nbsp;&nbsp;" 
                    end if
                    if aux(1) <> 0 then
                        textDin = textDin + Cstr(aux(1)) + "&times;50,00 "
                        if ((aux(1) <> 0) OR (aux(1) <> 0) OR (aux(1) <> 0) OR (aux(1) <> 0)) then textDin = textDin + "&nbsp;&nbsp;&nbsp;+&nbsp;&nbsp;&nbsp;"
                    end if
                    if aux(2) <> 0 then
                        textDin = textDin + Cstr(aux(2)) + "&times;20,00 "
                        if ((aux(3) <> 0) OR (aux(4) <> 0) OR (aux(5) <> 0)) then textDin = textDin + "&nbsp;&nbsp;&nbsp;+&nbsp;&nbsp;&nbsp;"
                    end if
                    if aux(3) <> 0 then
                        textDin = textDin + Cstr(aux(3)) + "&times;10,00 "
                        if ((aux(4) <> 0) OR (aux(5) <> 0)) then textDin = textDin + "&nbsp;&nbsp;&nbsp;+&nbsp;&nbsp;&nbsp;"
                    end if
                    if aux(4) <> 0 then
                        textDin = textDin + Cstr(aux(4)) + "&times;5,00 "
                        if  (aux(5) <> 0) then textDin = textDin + "&nbsp;&nbsp;&nbsp;+&nbsp;&nbsp;&nbsp;"
                    end if
                    if aux(5) <> 0 then
                        textDin = textDin + Cstr(aux(5)) + "&times;2,00 "
                    end if    


                        x = x & "<br>"& chr(13)

                        x = x & " <table  cellspacing='0'  width='700px'> " & chr(13) & _ 
                        "   <tr  nowrap style='background:honeydew'>"& chr(13) & _
                        "       <td width='30%' class='MDTE' align='left'><span class='Cd' style='color:black;' > TOTAL COMISSÃO ARREDONDADO</td> "& chr(13) & _
                        "       <td class='MD MC'  style='background:honeydew'><span id='totalComissaoAd' class='Cd'>" & formata_moeda(totalComissaoArredondado) &" </td> "& chr(13) & _
                        "   </tr>"& chr(13) & _
                        "   <tr nowrap>"& chr(13) 
                        if qtdeChq <> 0 then
                        x = x &"<td  style='background:honeydew;' width='30%' class='MDTE' align='left'><span class='Cd' style='color:black;' valign='bottom'>Comissão em CHQ </td>" & chr(13) & _
                        "       <td class='MD MC'  style='background:honeydew'><span id='totalCHQ' class='Cd'>" & formata_moeda(totalCHQ)& "&nbsp; (Nº cheques "& qtdeChq & ") </td> "& chr(13) & _
                        "   </tr>"& chr(13)
                        else
                        x = x &_
                        "       <td  style='background:honeydew;' width='30%' class='MDTE' align='left'><span class='Cd' style='color:black;' valign='bottom'>Comissão em CHQ </td>" & chr(13) & _
                        "       <td class='MD MC'  style='background:honeydew'><span id='totalCHQ' class='Cd'>" & formata_moeda(totalCHQ)& "&nbsp;</td> "& chr(13) & _
                        "   </tr>"& chr(13)
                        end if 
                        x = x &_
                        "    <tr nowrap >"& chr(13) & _
                        "       <td  style='background:honeydew' width='30%' class='MDTE' align='left'><span class='Cd' style='color:black;' >Comissão em DEP</td> "& chr(13) & _
                        "       <td class='MTD' style='background:honeydew'><span  id='totalComissaoDEP' class='Cd'>"& formata_moeda(totalDEP)&" </td> "& chr(13) & _
                        "   </tr>"& chr(13) & _
                        "   <tr nowrap >"& chr(13) & _
                        "       <td style='background:honeydew' width='30%' class='MDTE' align='left'><span  class='Cd' style='color:black;' >Comissão em DIN:</td>" & chr(13) & _
                        "       <td class='MD MC' style='background:honeydew'><span id='totalComissaoDIN' class='Cd'>"& formata_moeda(totalDIN)&"</td>"& chr(13) & _
                        "   </tr>"& chr(13) & _
                        "   <tr nowrap >"& chr(13) & _
                        "       <td  style='background:honeydew' width='30%' class='MTBE MD' align='left'><span class='Cd' style='color:black;' >Qtde Cedulas para Comissão em DIN</td>"& chr(13) & _
                        "       <td class='MTB MD' align='left'  style='background:honeydew'><span id='totalCedulasDIN' class='Cd'>" & textDin &" </td>"& chr(13) & _
                        "   </tr>"& chr(13) & _ 
                        "</table>"& chr(13) & _
                        "<br />" & chr(13)

                        ' Qtde pagamentos em cheque
                        x = x & "<table cellspacing='0' width='700px'>"& chr(13) & _
                                "   <tr>"& chr(13) & _
                                "       <td  style='background:honeydew' colspan='2' class='MTBE MD' align='left'><span class='Cd' style='color:black;'>CHQ - Pagamentos em Cheque</td>" & chr(13) & _ 
                                "   </tr>" & chr(13) & _
                                "   <tr>"& chr(13) & _
                                "       <td  style='background:honeydew' class='MB ME' align='left'><span class='Cd' style='color:black;'>Banco</td>" & chr(13) & _ 
                                "       <td  style='background:honeydew' class='MB ME MD' align='left'><span class='Cd' style='color:black;'>Quantidade</td>" & chr(13) & _ 
                                "   </tr>" & chr(13)
                        for i = Lbound(vChequesPorBanco) to Ubound(vChequesPorBanco)
			                with vChequesPorBanco(i)
                                if Trim("" & .c1) <> "" then
                                    Do While Len(.c1) < 3 
                                        .c1 = "0" & .c1
                                    Loop
                                    x = x & "<tr>" & _
                                            "       <td  style='background:honeydew' width='50%' class='MB ME' align='left'><span class='Cd' style='color:black;font-weight:normal'>" & .c1 & " - " & x_banco(.c1) & "</td>" & chr(13) & _ 
                                            "       <td  style='background:honeydew' width='50%' class='MB ME MD' align='left'><span class='Cd' style='color:black;font-weight:normal'>" & .c2 & "</td>" & chr(13) & _
                                            "   </tr>" & chr(13)
                                end if
                            end with
                        next
                        
                        x = x & "</table>" & chr(13) & _
                                "<br />" & chr(13)
                        
                        ' Qtde de pagamentos em dinheiro
                        x = x & "<table cellspacing='0' width='700px'>"& chr(13) & _
                                "   <tr>"& chr(13) & _
                                "       <td  style='background:honeydew' colspan='2' class='MTBE MD' align='left'><span class='Cd' style='color:black;'>DIN - Pagamentos em Dinheiro</td>" & chr(13) & _ 
                                "   </tr>" & chr(13) & _
                                "   <tr>"& chr(13) & _
                                "       <td  style='background:honeydew' class='MB ME' align='left'><span class='Cd' style='color:black;'>Banco</td>" & chr(13) & _ 
                                "       <td  style='background:honeydew' class='MB ME MD' align='left'><span class='Cd' style='color:black;'>Quantidade</td>" & chr(13) & _ 
                                "   </tr>" & chr(13)
                        for i = Lbound(vDinheiroPorBanco) to Ubound(vDinheiroPorBanco)
			                with vDinheiroPorBanco(i)
                                if Trim("" & .c1) <> "" then
                                    Do While Len(.c1) < 3 
                                        .c1 = "0" & .c1
                                    Loop
                                    x = x & "<tr>" & _
                                            "       <td  style='background:honeydew' width='50%' class='MB ME' align='left'><span class='Cd' style='color:black;font-weight:normal'>" & .c1 & " - " & x_banco(.c1) & "</td>" & chr(13) & _ 
                                            "       <td  style='background:honeydew' width='100' class='MB ME MD' align='left'><span class='Cd' style='color:black;font-weight:normal'>" & .c2 & "</td>" & chr(13) & _
                                            "   </tr>" & chr(13)
                                end if
                            end with
                        next
                        
                        x = x & "</table>"
                    
                       
			    end if
		    end if
        end if
      ' MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	    if n_reg_total = 0 then
		    x = cab_table & cab
		    x = x & "	<tr nowrap>" & chr(13) & _
				    "		<td class='MT ALERTA' colspan='12' align='center'><span class='ALERTA'>&nbsp;NÃO HÁ PEDIDOS NO PERÍODO ESPECIFICADO&nbsp;</span></td>" & chr(13) & _
				    "		<td class='notPrint BkgWhite'>&nbsp;</td>" & chr(13) & _
				    "	</tr>" & chr(13)
		    end if
      ' FECHA TABELA DO ÚLTIMO INDICADOR
	    x = x & "</table>" & chr(13)

         x = x & "<input type='hidden' name='ConsultaVendedor' id='ConsultaVendedor' value='"& vendedor &"'>"
        x = x & "<input type='hidden' name='filtroAno' id='filtroAno' value='"& c_dt_entregue_ano &"'>"
        x = x & "<input type='hidden' name='filtroMes' id='filtroMes' value='"& c_dt_entregue_mes &"'>"
	    Response.write x
    
    
	
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
<script src="<%=URL_FILE__JQUERY_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>
<link href="<%=URL_FILE__EPRINTER_CSS%>" rel="stylesheet" type="text/css" media="print">

<script language="JavaScript" type="text/javascript">
    var windowScrollTopAnterior;
    window.status = 'Aguarde, executando a consulta ...';

    $(function () {
        $("#divPedidoConsulta").hide();

        sizeDivPedidoConsulta();

        $('#divInternoPedidoConsulta').addClass('divFixo');


        var MostraVendedor, mes, ano;
        MostraVendedor = ""; mes = ""; ano = "";
        MostraVendedor = $("#ConsultaVendedor").val();
        mes = $("#filtroMes").val() + "/" + $("#filtroAno").val();
        $("#MostraVendedor").text(MostraVendedor);
        $("#Competencia").text(mes);


        $(document).keyup(function (e) {
            if (e.keyCode == 27) fechaDivPedidoConsulta();
        });

        $("#divPedidoConsulta").click(function () {
            fechaDivPedidoConsulta();
        });

        $("#imgFechaDivPedidoConsulta").click(function () {
            fechaDivPedidoConsulta();
        });


        
        // EXIBE O REALCE NOS CHECKBOXES QUE SÃO EXIBIDOS INICIALMENTE ASSINALADOS
        $(".CKB_COM:enabled").each(function () {
            if ($(this).is(":checked")) {
                $(this).parents("td.tdCkb").addClass("CKB_HIGHLIGHT");
            }
            else {
                $(this).parents("td.tdCkb").removeClass("CKB_HIGHLIGHT");
            }
        })

        // EVENTO P/ REALÇAR OU NÃO CONFORME SE MARCA/DESMARCA O CHECKBOX
        $(".CKB_COM:enabled").click(function () {
            if ($(this).is(":checked")) {
                $(this).parents("td.tdCkb").addClass("CKB_HIGHLIGHT");
            }
            else {
                $(this).parents("td.tdCkb").removeClass("CKB_HIGHLIGHT");
            }
        })

        // VISÃO SINTÉTICA?
        if ($("#rb_visao").val() == "SINTETICA") {
            $(".CKB_COM").attr("disabled", true);
        }

        $(".table_Desconto").hide();
    });


    function abreDesconto(idx_bloco) {
        var s_seletor = "#table_Desconto_" + idx_bloco;

        $(s_seletor).toggle();
    }

    //Every resize of window
    $(window).resize(function () {
        sizeDivPedidoConsulta();
    });

    function sizeDivPedidoConsulta() {
        var newHeight = $(document).height() + "px";
        $("#divPedidoConsulta").css("height", newHeight);
    }

    function fechaDivPedidoConsulta() {
        $(window).scrollTop(windowScrollTopAnterior);
        $("#divPedidoConsulta").fadeOut();
        $("#iframePedidoConsulta").attr("src", "");
    }

    function marcar_todos() {
        $(".CKB_COM:enabled")
            .prop("checked", true)
            .parents("td.tdCkb").addClass("CKB_HIGHLIGHT");
    }

    function desmarcar_todos() {
        $(".CKB_COM:enabled")
            .prop("checked", false)
            .parents("td.tdCkb").removeClass("CKB_HIGHLIGHT");
    }

    function trata_ckb_onclick(idx_bloco) {
        var s_id, s_class;
        s_id = "#ckb_comissao_paga_tit_bloco_" + idx_bloco;
        s_class = ".CKB_COM_BL_" + idx_bloco;
        if ($(s_id).is(":checked")) {
            $(s_class).prop("checked", true);
            $(s_class).parents("td.tdCkb").addClass("CKB_HIGHLIGHT");
        }
        else {
            $(s_class).prop("checked", false);
            $(s_class).parents("td.tdCkb").removeClass("CKB_HIGHLIGHT");
        }
    }

    function fExibeOcultaCampos(indice_bloco) {
        var s_seletor = "#tableDados_" + indice_bloco + " .VISAO_ANALIT";
        $(s_seletor).toggle();
    }

    function expandir_todos() {
        $(".VISAO_ANALIT").show();
    }

    function recolher_todos() {
        $(".VISAO_ANALIT").hide();
    }

    function fPEDConsulta(id_pedido, usuario) {
        windowScrollTopAnterior = $(window).scrollTop();
        sizeDivPedidoConsulta();
        $("#iframePedidoConsulta").attr("src", "PedidoConsultaView.asp?pedido_selecionado=" + id_pedido + "&pedido_selecionado_inicial=" + id_pedido + "&usuario=" + usuario);
        $("#divPedidoConsulta").fadeIn();
    }

    function fORCConsulta(id_orcamento, usuario) {
        windowScrollTopAnterior = $(window).scrollTop();
        sizeDivPedidoConsulta();
        $("#iframePedidoConsulta").attr("src", "OrcamentoConsultaView.asp?orcamento_selecionado=" + id_orcamento + "&orcamento_selecionado_inicial=" + id_orcamento + "&usuario=" + usuario);
        $("#divPedidoConsulta").fadeIn();
    }

    function fRELGravaDados(f) {
        window.status = "Aguarde ...";
        dCONFIRMA.style.visibility = "hidden";
        f.action = "RelComissaoIndicadoresPagGravaDados.asp";
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
<link href="<%=URL_FILE__EPRINTER_CSS%>" rel="stylesheet" type="text/css" media="print">

<style type="text/css">
.tdCkb
{
	width: 20px;
}

.break {
    page-break-before:always;
}

.tdLoja{
	width: 28px;
	}
.tdOrcamento{
	width: 58px;
	}
.tdPedido{
	width: 83px;
	}
.tdData{
	width: 60px;
	}
.tdVlPedido{
	width: 70px;
	}
.tdVlRT{
	width: 60px;
	}
.tdVlRABruto{
	width: 60px;
	}
.tdVlRALiq{
	width: 60px;
	}
.tdVlRADif{
	width: 60px;
	}
.tdStPagto{
	width: 60px;
	}
.tdSinal{
	width: 18px;
	}
.BTN_LNK
{
	min-width:140px;
}
.CKB_HIGHLIGHT
{
	background-color:#90EE90;
}
#divPedidoConsulta
{
	position:absolute;
	top:0;
	left:0;
	width:100%;
	z-index:1000;
	background-color:#808080;
	opacity: 1;
}
#divInternoPedidoConsulta
{
	position:absolute;
	top:6%;
	left:5%;
	width:90%;
	height:90%;
	z-index:1000;
	background-color:#808080;
	opacity: 1;
}
#divInternoPedidoConsulta.divFixo
{
	position:fixed;
	top:6%;
}
#imgFechaDivPedidoConsulta
{
	position:fixed;
	top:6%;
	left: 50%;
	margin-left: -16px; /* -1 * image width / 2 */
	margin-top: -32px;
	z-index:1001;
}
#iframePedidoConsulta
{
	position:absolute;
	top:0;
	left:0;
	width:100%;
	height:100%;
	border: solid 4px black;
}
.BkgWhite
{
	background-color:#ffffff;
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
<input type="hidden" name="orcamento_selecionado" id="orcamento_selecionado" value="">


<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="709" cellpadding="4" cellspacing="0" class="notPrint" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Relatório de Pedidos Indicadores (Consulta)</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>

<!-- FILTROS -->
<% Response.Write vendedor
	s_filtro = "<table width='709' cellpadding='0' cellspacing='0' style='border-bottom:1px solid black' border='0'>" & chr(13)

'	PERÍODO: MÊS DE COMPETÊNCIA
	s = ""
	if (c_dt_entregue_mes <> "") Or (c_dt_entregue_ano <> "") then
	'	DEVIDO AO WORD WRAP: SÓ FAZ WORD WRAP QUANDO ENCONTRA CHR(32), OU SEJA, MANTÉM AGRUPADO TEXTO COM &nbsp;
		if s <> "" then s = s & ",&nbsp; "
		s_aux = c_dt_entregue_mes
		if s_aux = "" then 
            s_aux = "N.I."
        else
            if c_dt_entregue_ano = "" then 
                s_aux = "N.I."
            else
	    	    s_aux = " " & s_aux & "/"
		        s_aux = replace(s_aux, " ", "&nbsp;")
		        s = s & s_aux
		        s_aux = ano_competencia
		        s_aux = replace(s_aux, " ", "&nbsp;")
            end if
        end if
		s = s & s_aux  
		end if

		s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='top' nowrap><span class='N'>Mês de competência:&nbsp;</span></td>" & chr(13) & _
					"		<td align='left' valign='top' width='99%'><span id='Competencia' class='N'>"&s&"</span></td>" & chr(13) & _
					"	</tr>" & chr(13)


'	COMISSÃO PAGA
	
	if s_aux<>"" then
		if s <> "" then s = s & ",&nbsp;&nbsp;"
		s = "paga"
		end if

	if s <> "" then
		s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='top' nowrap><span class='N'>Comissão:&nbsp;</span></td>" & chr(13) & _
					"		<td align='left' valign='top' width='99%'><span class='N'>paga</span></td>" & chr(13) & _
					"	</tr>" & chr(13)
		end if

'	STATUS DE PAGAMENTO
	s = ""
		s = "pago"

	if s <> "" then
		s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='top' nowrap><span class='N'>Status de Pagamento:&nbsp;</span></td>" & chr(13) & _
					"		<td align='left' valign='top' width='99%'><span class='N'>" & s & "</span></td>" & chr(13) & _
					"	</tr>" & chr(13)
		end if

'	INDICADOR
		s = "todos"
		s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='top' nowrap><span class='N'>Indicador:&nbsp;</span></td>" & chr(13) & _
					"		<td align='left' valign='top' width='99%'><span class='N'>" & s & "</span></td>" & chr(13) & _
					"	</tr>" & chr(13)

'	VENDEDOR
		s =  vendedor
		's_aux = x_usuario(vendedor)
		s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='top' nowrap><span class='N'>Vendedor(es):&nbsp;</span></td>" & chr(13) & _
					"		<td align='left'  valign='top' width='99%'><span id='MostraVendedor' class='N'></span></td>" & chr(13) & _
					"	</tr>" & chr(13)

	s_filtro = s_filtro & "</table>" & chr(13)
	Response.Write s_filtro
     
%>

<br />
<!--  RELATÓRIO  -->
<br>
<br />
<!--  RELATÓRIO  -->
<br>
<% consulta_executa %>
<input type="hidden" name="c_id" id="c_id" value="<%=s_id%>" />
<!-- ************   SEPARADOR   ************ -->
<table class="notPrint" width="709" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
	<td class="Rc" align="left">&nbsp;</td>
</tr>
</table>

<% if blnVisaoSintetica then %>
<table class="notPrint" width="709" cellpadding="0" cellspacing="0" style="margin-top:5px;">
<tr>
	<td align="right">
		<button type="button" name="bExpandirTodos" id="bExpandirTodos" class="Button BTN_LNK" onclick="expandir_todos();" title="expandir todas as linhas de dados" style="margin-left:6px;margin-bottom:2px">Expandir Tudo</button>
		&nbsp;
		<button type="button" name="bRecolherTodos" id="bRecolherTodos" class="Button BTN_LNK" onclick="recolher_todos();" title="recolher todas as linhas de dados" style="margin-left:6px;margin-right:6px;margin-bottom:2px">Recolher Tudo</button>
	</td>
</tr>
</table>
<br />
<table class="notPrint" width="709" cellspacing="0">
<tr>
	<td align="center"><a name="bVOLTA" id="A1" href="javascript:history.back()" title="volta para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a>
	</td>
</tr>
</table>
<% else %>
<table class="notPrint" width="709" cellpadding="0" cellspacing="0" style="margin-top:5px;">
<tr>
	<td align="left">
		<button type="button" name="bExpandirTodos" id="bExpandirTodos" class="Button BTN_LNK" onclick="expandir_todos();" title="expandir todas as linhas de dados" style="margin-left:6px;margin-bottom:2px">Expandir Tudo</button>
		&nbsp;
		<button type="button" name="bRecolherTodos" id="bRecolherTodos" class="Button BTN_LNK" onclick="recolher_todos();" title="recolher todas as linhas de dados" style="margin-left:6px;margin-right:6px;margin-bottom:2px">Recolher Tudo</button>
		
		
	</td>
</tr>
</table>

<br />
<table class="notPrint" width="709" cellspacing="0">
<tr>
	<td align="left"><a name="bVOLTA" id="bVOLTA" href="javascript:history.back()" title="volta para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a>
	</td>
	<td align="left">&nbsp;</td>
	<td align="right">
		<div name="dIMPRIME" id="dIMPRIME"><a name="bIMPRIME" id="bIMPRIME" href="javascript:window.print();" title="imprimir relatório"><img src="../botao/imprimir.gif" width="176" height="55" border="0"></a></div>
	</td>
</tr>
</table>
<% end if %>
</form>

</center>

<div id="divPedidoConsulta"><center><div id="divInternoPedidoConsulta"><img id="imgFechaDivPedidoConsulta" src="../imagem/close_button_32.png" title="clique para fechar o painel de consulta" /><iframe id="iframePedidoConsulta"></iframe></div></center></div>

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
