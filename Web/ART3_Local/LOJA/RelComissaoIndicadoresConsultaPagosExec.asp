<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =====================================================================================
'	  R E L C O M I S S A O I N D I C A D O R E S C O N S U L T A P A G O S E X E C . A S P
'     =====================================================================================
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
   
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))


	dim alerta
	dim s, s_aux, s_filtro, s_id, mes_competencia, ano_competencia
    dim aviso, rb_visao, blnVisaoSintetica
    
    blnVisaoSintetica = False
	if rb_visao = "SINTETICA" then blnVisaoSintetica = True
	alerta = ""

    mes_competencia = Trim(Request.Form("c_dt_entregue_mes"))
    ano_competencia = Trim(Request.Form("c_dt_entregue_ano"))


    dim o
dim strMsg
dim resultadoCalculo,resultadoDigito,QtdeCedulas,TotalCedula,limitador(5)
dim dadosCalculo
set o = createobject( "ComPlusCalcCedulas.ComPlusCalcCedulas" )
dim v_cedulas,aux(5),y,totalArredondado, cont
dim cedulas()
dim qtdeCedula(), limitador_fixo(5)


' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________

    function x_meio_pagto(x)
        dim s
        select case x
           case "DEP" : s = "Pagamento em Depósito"
           case "CHQ" : s = "Pagamento com Cheque"
           case "DIN" : s = "Pagamento em Dinheiro"
           case else : s=""
        end select
    x_meio_pagto = s
    end function

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
dim textDin,vendedor

vendedor = ""
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
            	
    s_sql = "SELECT t_COMISSAO_INDICADOR_N3.indicador as ind3,t_COMISSAO_INDICADOR_N2.id as id_n2,* FROM t_COMISSAO_INDICADOR_N1" & _
            " INNER JOIN t_COMISSAO_INDICADOR_N2 ON (t_COMISSAO_INDICADOR_N1.id = t_COMISSAO_INDICADOR_N2.id_comissao_indicador_n1)" & _
            " INNER JOIN t_COMISSAO_INDICADOR_N3 ON (t_COMISSAO_INDICADOR_N2.id = t_COMISSAO_INDICADOR_N3.id_comissao_indicador_n2)" & _
            " INNER JOIN t_COMISSAO_INDICADOR_N4 ON (t_COMISSAO_INDICADOR_N3.id = t_COMISSAO_INDICADOR_N4.id_comissao_indicador_n3)" & _
            " INNER JOIN t_PEDIDO ON (t_COMISSAO_INDICADOR_N4.pedido = t_PEDIDO.pedido)" & _
            " INNER JOIN t_CLIENTE ON (t_PEDIDO.id_cliente=t_CLIENTE.id)" & _
            " WHERE (t_COMISSAO_INDICADOR_N2.competencia_mes = '" & mes_competencia & "') AND (t_COMISSAO_INDICADOR_N2.competencia_ano = '" & ano_competencia & "') AND (t_COMISSAO_INDICADOR_N2.vendedor = '" & usuario & "') AND (t_COMISSAO_INDICADOR_N3.st_tratamento_manual='0')" & _
            " ORDER BY t_COMISSAO_INDICADOR_N3.indicador," & _
            " t_COMISSAO_INDICADOR_N3.numero_banco," & _
            " t_COMISSAO_INDICADOR_N3.vl_total_comissao_arredondado"
            
		
  ' CABEÇALHO
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
	set r = cn.execute(s_sql)

    
    redim vRelat(1)
	for intIdxVetor = Lbound(vRelat) to Ubound(vRelat)
		set vRelat(intIdxVetor) = New cl_VINTE_COLUNAS
		vRelat(intIdxVetor).CampoOrdenacao = ""
		next

    aviso=""  
    x = ""

	do while Not r.Eof

	'	MUDOU DE INDICADOR?
		if Trim("" & r("ind3"))<>indicador_a then
			indicador_a = Trim("" & r("ind3"))            
			idx_bloco = idx_bloco + 1
			qtde_indicadores = qtde_indicadores + 1
        
		  ' FECHA TABELA DO INDICADOR ANTERIOR
			if n_reg_total > 0 then 

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
                        "       <td align='right' colspan='9'><span class='Cd' style='color:" & s_cor & ";'>" & "TOTAL PAGO:</span></td>" & chr(13) & _
                        "       <td align='right' colspan='2'><span class='Cd' style='color:" & s_cor & ";'>"&formata_moeda(vl_sub_total_RT_arredondado+vl_sub_total_RA_arredondado) & "</span></td>" & chr(13) & _
                        "	</tr>" & chr(13) & _
						"</table>" & chr(13)

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
									        "					<td colspan='3' align='left' valign='bottom' style='vertical-align:middle'><div valign='bottom' style='height:14px;max-height:14px;overflow:hidden;vertical-align:middle'><span class='Cn'>Banco: " & Trim("" & rs("banco")) & " - " & x_banco(Trim("" & rs("banco"))) &  "</span></div></td>" & chr(13) & _
									        "				</tr>" & chr(13) & _
									        "				<tr>" & chr(13) & _
									        "					<td class='MTD' align='left' valign='bottom' style='height:15px;vertical-align:middle'><span class='Cn'>Agência: " & Trim("" & rs("agencia"))
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
    
                                x = x & Trim("" & rs("conta"))
    
                                if Trim("" & rs("conta_dv")) <> "" then
                                    x = x & "-" & rs("conta_dv") & chr(13)
                                end if


								x = x &	"					<td class='MC' width='60%' align='left' valign='bottom'><span class='Cn'>Favorecido: " & s_favorecido & "</span></td>" & chr(13) & _
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
        vl_sub_total_comissao = r("vl_total_pagto")
           

        if s_lista_completa_pedidos <> "" then s_lista_completa_pedidos = s_lista_completa_pedidos & ";"
		s_lista_completa_pedidos = s_lista_completa_pedidos & Trim("" & r("pedido"))

     
     '> CHECK BOX
	 '	É USADO O CÓDIGO DA OPERAÇÃO (VENDA NORMAL, DEVOLUÇÃO, PERDA) P/ NÃO CORRER O RISCO DE HAVER CONFLITO DEVIDO A ID'S REPETIDOS ENTRE AS OPERAÇÕES
'		s_id = "ckb_comissao_paga_" & Trim("" & r("operacao")) & "_" & Trim("" & r("id_registro"))
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
		s_nome_cliente = Trim("" & r("nome_iniciais_em_maiusculas"))
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
	
		    ind_anterior = Trim("" & r("ind3"))
			banco = Trim("" & r("banco"))
        	meio_pagto = r("meio_pagto")
        
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
      
		x = x & "	<tr nowrap style='background: #FFFFDD'>" & chr(13) & _
				"		<td colspan='4' class='MTBE' align='right' nowrap><span class='Cd' style='color:" & s_cor & ";'>" & _
										"TOTAL:</span></td>" & chr(13) & _
						"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_preco_venda) & "</span></td>" & chr(13) & _                      
						"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_RT) & "</span></td>" & chr(13) & _                                          
						"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_RA) & "</span></td>" & chr(13) & _                       
						"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_RA_liquido) & "</span></td>" & chr(13) & _                        
						"		<td class='MTB' align='right'><span class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_sub_total_RA_diferenca) & "</span></td>" & chr(13) & _                       
						"		<td class='MTBD' align='right' colspan='2'><span class='Cd' style='color:" & s_cor & ";'>&nbsp;</span></td>" & chr(13) & _
						"		<td class='notPrint BkgWhite'>&nbsp;</td>" & chr(13) & _
						"	</tr>" & chr(13) & _
				        "       <td align='right' colspan='9'><span class='Cd' style='color:" & s_cor & ";'>" & "TOTAL PAGO:</span></td>" & chr(13) & _
                        "       <td align='right' colspan='2'><span class='Cd' style='color:" & s_cor & ";'>"&formata_moeda(vl_sub_total_RT_arredondado+vl_sub_total_RA_arredondado) & "</span></td>" & chr(13) & _
                        "	</tr>" & chr(13) 
						
                
                
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
					"		<td class='MTB' align='right'><span id ='total_VlPedido' class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_preco_venda) & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span id='totalComissao' class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_RT) & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span id='total_RA' class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_RA) & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span id='total_RAliq' class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_RA_liquido) & "</span></td>" & chr(13) & _
					"		<td class='MTB' align='right'><span id='total_RAdif' class='Cd' style='color:" & s_cor & ";'>" & formata_moeda(vl_total_RA_diferenca) & "</span></td>" & chr(13) & _
					"		<td class='MTBD' align='right' colspan='2'><span class='Cd' style='color:" & s_cor & ";'>&nbsp;</span></td>" & chr(13) & _
					"		<td class='notPrint BkgWhite'>&nbsp;</td>" & chr(13) & _
					"	</tr>" & chr(13) &_
                    " </table>" & chr(13)

			end if
		end if

  ' MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	if n_reg_total = 0 then
		x = cab_table & cab
		x = x & "	<tr nowrap>" & chr(13) & _
				"		<td class='MT ALERTA' colspan='11' align='center'><span class='ALERTA'>&nbsp;O RELATÓRIO DO MÊS INFORMADO AINDA NÃO FOI PROCESSADO&nbsp;</span></td>" & chr(13) & _
				"		<td class='notPrint BkgWhite'>&nbsp;</td>" & chr(13) & _
				"	</tr>" & chr(13)
		end if

  ' FECHA TABELA DO ÚLTIMO INDICADOR
	x = x & "</table>" & chr(13)

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
	<title>LOJA</title>
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
    });

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
	<td align="right" valign="bottom"><span class="PEDIDO">Relatório de Pedidos Indicadores (Processado)</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>

<!-- FILTROS -->
<%
	s_filtro = "<table width='709' cellpadding='0' cellspacing='0' style='border-bottom:1px solid black' border='0'>" & chr(13)

'	PERÍODO: MÊS DE COMPETÊNCIA
	s = ""
	if (mes_competencia <> "") Or (ano_competencia <> "") then
	'	DEVIDO AO WORD WRAP: SÓ FAZ WORD WRAP QUANDO ENCONTRA CHR(32), OU SEJA, MANTÉM AGRUPADO TEXTO COM &nbsp;
		if s <> "" then s = s & ",&nbsp; "
		s_aux = mes_competencia
		if s_aux = "" then 
            s_aux = "N.I."
        else
            if ano_competencia = "" then 
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

	if s <> "" then
		s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='top' nowrap><span class='N'>Mês de competência:&nbsp;</span></td>" & chr(13) & _
					"		<td align='left' valign='top' width='99%'><span class='N'>" & s & "</span></td>" & chr(13) & _
					"	</tr>" & chr(13)
		end if

'	COMISSÃO PAGA
	
	if s_aux<>"" then
		if s <> "" then s = s & ",&nbsp;&nbsp;"
		s = "paga"
		end if

	if s <> "" then
		s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='top' nowrap><span class='N'>Comissão:&nbsp;</span></td>" & chr(13) & _
					"		<td align='left' valign='top' width='99%'><span class='N'>" & s & "</span></td>" & chr(13) & _
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
        
		s = usuario
		s_aux = x_usuario(usuario)
		if (s <> "") And (s_aux <> "") then s = s & " - "
		s = s & s_aux
		s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='top' nowrap><span class='N'>Vendedor(es):&nbsp;</span></td>" & chr(13) & _
					"		<td align='left' valign='top' width='99%'><span class='N'>" & s & "</span></td>" & chr(13) & _
					"	</tr>" & chr(13)

'	EMISSÃO
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>Emissão:&nbsp;</span></td><td align='left' valign='top' width='99%'>" & _
			   "<span class='N'>" & formata_data_hora(Now) & "</span></td></tr>" & chr(13)

	s_filtro = s_filtro & "</table>" & chr(13)
	Response.Write s_filtro
%>

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
