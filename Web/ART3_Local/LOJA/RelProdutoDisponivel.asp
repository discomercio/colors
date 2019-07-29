<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ======================================================
'	  R E L P R O D U T O D I S P O N I V E L . A S P
'     ======================================================
'
''	  S E R V E R   S I D E   S C R I P T I N G
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

	dim usuario, loja
	usuario = Trim(Session("usuario_atual"))
	loja = Trim(Session("loja_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 
	If (loja = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	CONECTA COM O BANCO DE DADOS
	dim cn, rs, msg_erro,r,tPCI,t
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
    If Not cria_recordset_otimista(tPCI, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
    If Not cria_recordset_otimista(t, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	
	dim alerta, s, s_aux,s_sql
	alerta = ""

	dim c_fabricante, c_produto, qtde_disponivel, s_nome_fabricante, s_nome_produto
    dim blnPularProdutoComposto,qtde_estoque_venda_composto,qtde_estoque_venda_aux,n_reg,blnProdutoComposto,descricao,blnvendavel,v_fabricante(),v_produto(),cont
    blnProdutoComposto = false
    blnvendavel = true
    cont = 0
	qtde_disponivel = 0
	s_nome_fabricante = ""
	s_nome_produto = ""
	descricao = ""
	c_fabricante = Trim(Request.Form("c_fabricante"))
	c_produto = Trim(Request.Form("c_produto"))

	if c_fabricante = "" then
		alerta = "Especifique o código do fabricante."
	elseif c_produto = "" then
		alerta = "Especifique o código do produto."
		end if

    
     if alerta = "" then
		c_fabricante = normaliza_codigo(c_fabricante, TAM_MIN_FABRICANTE)	
		s = normaliza_produto(c_produto)
		if s <> "" then c_produto = s
        s = "SELECT * FROM t_EC_PRODUTO_COMPOSTO WHERE produto_composto = '" & c_produto & "' AND fabricante_composto = '" & c_fabricante & "'"

        rs.open s, cn
        if not rs.Eof then
            blnProdutoComposto = true            		
		end if
    end if
		
	if alerta = "" then
        if blnProdutoComposto = false then		    
	
		    s = "SELECT Sum(qtde-qtde_utilizada) AS saldo" & _
			    " FROM t_ESTOQUE_ITEM WHERE" & _
			    " ((qtde-qtde_utilizada) > 0)" & _
			    " AND (fabricante='" & c_fabricante & "')" & _
			    " AND (produto='" & c_produto & "')" 
            if rs.State <> 0 then rs.Close
		    rs.open s, cn
		    if Not rs.Eof then
			    if IsNumeric(rs("saldo")) then qtde_disponivel = CLng(rs("saldo"))
			end if
		end if
    end if

    if alerta = "" then       
        if blnProdutoComposto = True then                 	
	
		    s_sql = " SELECT t_EC_PRODUTO_COMPOSTO_ITEM.fabricante_item" &_
	                    " ,t_EC_PRODUTO_COMPOSTO_ITEM.produto_item" &_
	                    " ,descricao" &_
                    " FROM t_EC_PRODUTO_COMPOSTO_ITEM" &_
                    " INNER JOIN t_EC_PRODUTO_COMPOSTO ON (t_EC_PRODUTO_COMPOSTO_ITEM.produto_composto = t_EC_PRODUTO_COMPOSTO.produto_composto)" &_
                    " WHERE t_EC_PRODUTO_COMPOSTO.produto_composto = '" & c_produto & "'" &_
	                    " AND t_EC_PRODUTO_COMPOSTO.fabricante_composto = '" & c_fabricante & "'"

		    n_reg = 0
		    set r = cn.execute(s_sql)
            if r.State <> 0 then r.Close
            r.open s_sql, cn
		    do while Not r.Eof                            
                redim preserve v_fabricante(cont)
                redim preserve v_produto(cont)
                v_fabricante(cont) = r("fabricante_item")
                v_produto(cont)   = r("produto_item")
                descricao = Trim("" & r("descricao"))	
                cont = cont + 1
                r.MoveNext
			loop
			blnPularProdutoComposto = False
            
			qtde_estoque_venda_composto = -1	
            for cont = Lbound(v_fabricante) to Ubound (v_fabricante)              
				    s_sql = " SELECT" & _
							    " tP.fabricante," & _
							    " tP.produto," & _						    
							    " Coalesce((SELECT Sum(qtde-qtde_utilizada) FROM t_ESTOQUE_ITEM tEI WHERE (tEI.fabricante=tP.fabricante) AND (tEI.produto=tP.produto) AND ((qtde-qtde_utilizada)>0)), 0) AS qtde_estoque_venda" & _                 
						    " FROM t_PRODUTO tPL" & _
							    " INNER JOIN t_PRODUTO tP ON (tPL.fabricante = tP.fabricante) AND (tPL.produto = tP.produto)" & _
                                " INNER JOIN t_PRODUTO_LOJA on (tPL.fabricante = t_PRODUTO_LOJA.fabricante) AND (tPL.produto = t_PRODUTO_LOJA.produto)" & _                                            
						    " WHERE " & _
                            " (tP.fabricante = '" & Trim("" & v_fabricante(cont)) & "')" & _
                       	    " AND (tP.produto = '" & Trim("" & v_produto(cont)) & "') "   
                    if loja <> "" then                         								
				          s_sql = s_sql + " AND (t_PRODUTO_LOJA.loja = '" & loja &"')"
                    end if  				    
                    
				    if t.State <> 0 then t.Close
				    t.Open s_sql, cn
				    if t.Eof then
				        blnPularProdutoComposto = true 
                               
				    else                                                  					
					    qtde_estoque_venda_aux = t("qtde_estoque_venda")
					    if qtde_estoque_venda_composto = -1 then
						    qtde_estoque_venda_composto = qtde_estoque_venda_aux
					    else
						    if qtde_estoque_venda_aux < qtde_estoque_venda_composto then
							    qtde_estoque_venda_composto = qtde_estoque_venda_aux
                                
						    end if
					    end if   
				    end if
                        
				    if blnPularProdutoComposto then exit for
				        
                        
			        if qtde_estoque_venda_composto >= 0 then                        
                       if Not blnPularProdutoComposto then				
			             '> SALDO ESTOQUE                    
				            qtde_disponivel = qtde_estoque_venda_composto
                            blnProdutoComposto = true				                
		                end if
			        end if
            next
            if r.State <> 0 then r.Close
	        set r = nothing			    
        end if
    end if


	if alerta = "" then
		s = "SELECT nome, razao_social FROM t_FABRICANTE WHERE (fabricante='" & c_fabricante & "')"
		if rs.State <> 0 then rs.Close
		rs.open s, cn
		if Not rs.Eof then
			s_nome_fabricante = Trim("" & rs("nome"))
			if s_nome_fabricante = "" then s_nome_fabricante = Trim("" & rs("razao_social"))
		else
			alerta = "Fabricante " & c_fabricante & " não está cadastrado."
			end if
		end if
		
	if alerta = "" then        
	'	O FLAG "EXCLUIDO_STATUS" INDICA SE O PRODUTO ESTÁ EXCLUÍDO LOGICAMENTE DO SISTEMA!!
	'	A TABELA BÁSICA DE PRODUTOS MANTÉM INFORMAÇÕES DE PRODUTOS EXCLUÍDOS LOGICAMENTE 
	'	P/ MANTER A REFERÊNCIA COM OUTRAS TABELAS QUE NECESSITEM DE DADOS COMO DESCRIÇÃO, ETC.
		s = "SELECT descricao, descricao_html, excluido_status FROM t_PRODUTO WHERE (fabricante='" & c_fabricante & "') AND (produto='" & c_produto & "')"
		if rs.State <> 0 then rs.Close
		rs.open s, cn
		if rs.Eof then
		'	PRODUTO NÃO ESTÁ CADASTRADO NA TABELA BÁSICA, PORTANTO NÃO ESTÁ DISPONÍVEL P/ VENDAS EM NENHUMA LOJA
            if blnProdutoComposto <> true then
			    qtde_disponivel = 0
			    alerta = "Produto " & c_produto & " do fabricante " & c_fabricante & " não está cadastrado."            
            else            
                s_nome_produto = descricao
            end if
		else
		'	PRODUTO ESTÁ EXCLUÍDO LOGICAMENTE DA TABELA BÁSICA, PORTANTO NÃO ESTÁ DISPONÍVEL P/ VENDAS EM NENHUMA LOJA
			if rs("excluido_status")<>0 then qtde_disponivel = 0
                if blnProdutoComposto = true then
                    s_nome_produto = descricao
                else
			        s_nome_produto = produto_formata_descricao_em_html(Trim("" & rs("descricao_html")))
                end if
			end if
		end if
		
	if alerta = "" then
		s = "SELECT qtde_max_venda, vendavel FROM t_PRODUTO_LOJA WHERE (loja='" & loja & "') AND (fabricante='" & c_fabricante & "') AND (produto='" & c_produto & "')"
		if rs.State <> 0 then rs.Close
		rs.open s, cn
		if rs.Eof then
		'	PRODUTO NÃO ESTÁ CADASTRADO P/ VENDA NESTA LOJA
            if blnProdutoComposto <> true then
			    qtde_disponivel = 0
			    alerta = "Produto " & c_produto & " do fabricante " & c_fabricante & " não está cadastrado."
            end if
		else
            '	ESTÁ DISPONÍVEL P/ VENDA? PRODUTO COMPOSTO/NORMAIS	
            if blnProdutoComposto = true then
                s_sql = "SELECT produto_item" & _
                            " ,vendavel" & _
	                        " ,sequencia" & _
	                        " ,qtde_max_venda" & _
                        " FROM t_EC_PRODUTO_COMPOSTO_ITEM" & _
                        " INNER JOIN t_PRODUTO_LOJA ON (t_EC_PRODUTO_COMPOSTO_ITEM.produto_item = t_PRODUTO_LOJA.produto)" & _
                        " WHERE fabricante_composto = '" & c_fabricante & "'" & _
	                        " AND produto_composto = '" & c_produto & "'" & _                      
	                        " AND loja = '" & loja & "' "
                if t.State <> 0 then t.Close
		        t.open s_sql, cn
                
                do while Not t.Eof
                    if t("vendavel") = "N" then blnvendavel = false 
                    '	A INFORMAÇÃO DE DISPONIBILIDADE DO PRODUTO NO ESTOQUE É LIMITADA 
		        '	PELA INFORMAÇÃO DE QTDE MÁXIMA POR VENDA, DE MODO QUE O VENDEDOR 
		        '	NUNCA SAIBA A POSIÇÃO VERDADEIRA DO ESTOQUE.                 
                if IsNumeric(t("qtde_max_venda")) then                                              
				        if CLng(t("qtde_max_venda")) < CLng(qtde_disponivel) then
					        qtde_disponivel = CLng(t("qtde_max_venda"))                   
					    end if
				end if 
                    t.MoveNext
                loop
                if blnvendavel <> true then qtde_disponivel = 0                
                       
            else               	    
			    if UCase(Trim("" & rs("vendavel"))) <> "S" then qtde_disponivel = 0              
                    if IsNumeric(rs("qtde_max_venda")) then
				        if CLng(rs("qtde_max_venda")) < qtde_disponivel then
					        qtde_disponivel = CLng(rs("qtde_max_venda"))                 
					    end if
				    end if
            end if
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
	<title>LOJA</title>
	</head>



<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
window.status='Aguarde, executando a consulta ...';
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
<body onload="window.status='Concluído';">

<center>

<form id="fREL" name="fREL" method="post">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Consulta Disponibilidade no Estoque</span>
	<br>
	<%	s = "<span class='N'>Emissão:&nbsp;" & formata_data_hora(Now) & "</span>"
		Response.Write s
	%>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>
<br>


<!--  RELATÓRIO  -->
<table class="Qx" cellSpacing="0">
	<% 	s = c_fabricante
		if (c_fabricante <> "") And (s_nome_fabricante <> "") then s = s & " - "
		s = s & iniciais_em_maiusculas(s_nome_fabricante)
	%>
	<tr bgColor="#FFFFFF"><td class="MT" NOWRAP><span class="PLTe">Fabricante</span>
		<br><span class="C" style="width:340px;cursor:default;"><%=s%></span></td>
	</tr>
	<% 	s = c_produto
		if (c_produto <> "") And (s_nome_produto <> "") then s = s & " - "
		s = s & s_nome_produto
	%>
	<tr bgColor="#FFFFFF"><td class="MDBE" NOWRAP><span class="PLTe">Produto</span>
		<br><span class="C" style="width:340px;cursor:default;"><%=s%></span></td>
	</tr>
	<%	if qtde_disponivel = 0 then
			s = "z e r o"
			s_aux = "red"
		else
			s = formata_inteiro(qtde_disponivel)
			s_aux = "green"
			end if
	%>
	<tr bgColor="#FFFFFF"><td class="MDBE" NOWRAP><span class="PLTe">Quantidade Disponível</span>
		<br><span class="C" style="font-size:12pt;width:100px;cursor:default;color:<%=s_aux%>;"><%=s%></span></td>
	</tr>
</table>


<!-- ************   SEPARADOR   ************ -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>
<br>


<table class="notPrint" width="649" cellSpacing="0">
<tr>
	<td align="center"><a name="bVOLTA" id="bVOLTA" href="javascript:history.back()" title="volta para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
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
