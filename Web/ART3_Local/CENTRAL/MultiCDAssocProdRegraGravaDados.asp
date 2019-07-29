<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/Global.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ===============================================================================
'	  MultiCDAssocProdRegraGravaDados.asp
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

	class cl_TIPO_GRAVA_PRODUTO_REGRA
		dim fabricante
		dim produto
		dim regra
		end class
		
	dim s, msg_erro
	dim usuario
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO)

	dim alerta
	alerta=""

    dim c_qtde_produtos, intQtdeProdutos, vProduto, regra_id

	c_qtde_produtos=Trim(Request("c_qtde_produtos"))
	intQtdeProdutos=CInt(c_qtde_produtos)
	
	redim vProduto(0)
	set vProduto(Ubound(vProduto)) = new cl_TIPO_GRAVA_PRODUTO_REGRA
	vProduto(Ubound(vProduto)).produto = ""
	
	dim i
	dim c_fabricante, c_produto, c_regra

	for i = 1 to intQtdeProdutos
		c_fabricante = Trim(Request.Form("c_fabricante_" & Cstr(i)))
		c_produto = Trim(Request.Form("c_produto_" & Cstr(i)))
		c_regra = Trim(Request.Form("c_regra_" & Cstr(i)))
			if vProduto(Ubound(vProduto)).produto <> "" then
				redim preserve vProduto(Ubound(vProduto)+1)
				set vProduto(Ubound(vProduto)) = new cl_TIPO_GRAVA_PRODUTO_REGRA
			end if
			vProduto(Ubound(vProduto)).fabricante = c_fabricante
			vProduto(Ubound(vProduto)).produto = c_produto
			vProduto(Ubound(vProduto)).regra = c_regra
		next	
	
	dim s_log_inclusao, s_log_alteracao, s_log_exclusao

	s_log_inclusao = ""
	s_log_alteracao = ""
	s_log_exclusao = ""
    
'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn, rs
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

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
			
		for i=Lbound(vProduto) to Ubound(vProduto)
			if Trim(vProduto(i).produto)<>"" then     

				if alerta = "" then
					s = "SELECT * FROM t_PRODUTO_X_WMS_REGRA_CD WHERE (fabricante = '" & vProduto(i).fabricante & "' AND produto = '" & vProduto(i).produto & "')"
					if rs.State <> 0 then rs.Close
					rs.Open s, cn
                    '  SE NÃO EXISTE REGRA CADASTRADA PARA ESTE PRODUTO, VERIFICA SE USUÁRIO SELECIONOU ALGUMA REGRA E INSERE O REGISTRO
                    if rs.Eof then           
                        if Trim(vProduto(i).regra)<>"" then     
					        rs.AddNew 
					        rs("fabricante")=vProduto(i).fabricante
                            rs("produto")=vProduto(i).produto
					        rs("id_wms_regra_cd")=vProduto(i).regra
					        rs("usuario_cadastro")=usuario
                            rs("usuario_ult_atualizacao")=usuario
					        rs.Update 
					        if Err <> 0 then
					        '	~~~~~~~~~~~~~~~~
						        cn.RollbackTrans
					        '	~~~~~~~~~~~~~~~~
						        Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
						        end if
						
                            if s_log_inclusao <> "" then s_log_inclusao = s_log_inclusao & "; "
					        s_log_inclusao = s_log_inclusao & "(" & vProduto(i).fabricante & ")" & vProduto(i).produto & ": " & "id_wms_regra_cd=" & vProduto(i).regra
						
					        if rs.State <> 0 then rs.Close                  							
                        end if
                    else ' SE EXISTE UMA REGRA CADASTRADA PARA ESTE PRODUTO
                        regra_id = Trim("" & rs("id_wms_regra_cd"))
                        if Trim(vProduto(i).regra)<>"" then
                            ' SE O USUÁRIO SELECIONOU OUTRA REGRA PARA ESTE PRODUTO, ALTERA O REGISTRO
                            if Trim("" & rs("id_wms_regra_cd"))<>Trim(vProduto(i).regra) then
                                rs("id_wms_regra_cd")=vProduto(i).regra                                
                                rs("usuario_ult_atualizacao")=usuario
                                rs("dt_ult_atualizacao")=Date
                                rs("dt_hr_ult_atualizacao")=Now
                                rs.Update
                                if Err <> 0 then
					            '	~~~~~~~~~~~~~~~~
						            cn.RollbackTrans
					            '	~~~~~~~~~~~~~~~~
						            Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
						            end if         

                                if s_log_alteracao <> "" then s_log_alteracao = s_log_alteracao & "; "
						        s_log_alteracao = s_log_alteracao & "(" & vProduto(i).fabricante & ")" & vProduto(i).produto & ": " & "id_wms_regra_cd=" & regra_id & " => " & vProduto(i).regra              
                                                       				            						
					            if rs.State <> 0 then rs.Close                  							
                            end if
                        else ' SE O USUÁRIO APAGOU A REGRA PARA ESTE PRODUTO
                            s = "DELETE FROM t_PRODUTO_X_WMS_REGRA_CD WHERE (fabricante = '" & vProduto(i).fabricante & "' AND produto = '" & vProduto(i).produto & "')"
                            cn.Execute(s)
                            
                            if Err <> 0 then
					            '	~~~~~~~~~~~~~~~~
						            cn.RollbackTrans
					            '	~~~~~~~~~~~~~~~~
						            Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_BD)
						            end if

                            if s_log_exclusao <> "" then s_log_exclusao = s_log_exclusao & "; "
                            s_log_exclusao = s_log_exclusao & "(" & vProduto(i).fabricante & ")" & vProduto(i).produto & ": " & "id_wms_regra_cd=" & regra_id
                        end if
					end if
				end if 'alerta=""  
            end if 'if Trim(vProduto(i).produto)<>"" 
		next
			
		if alerta = "" then
		'	~~~~~~~~~~~~~~
			cn.CommitTrans
		'	~~~~~~~~~~~~~~
			if Err=0 then 
		        ' GRAVA O LOG
				if s_log_inclusao <> "" then grava_log usuario, "", "", "", OP_LOG_PRODUTO_REGRA_CD_INCLUSAO, s_log_inclusao
				if s_log_alteracao <> "" then grava_log usuario, "", "", "", OP_LOG_PRODUTO_REGRA_CD_ALTERACAO, s_log_alteracao
				if s_log_exclusao <> "" then grava_log usuario, "", "", "", OP_LOG_PRODUTO_REGRA_CD_EXCLUSAO, s_log_exclusao
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

<body onload="bVOLTAR.focus();">
<center>
<% if alerta <> "" then %>
<!-- ************************************************************ -->
<!-- **********  PÁGINA PARA EXIBIR MENSAGENS DE ERRO  ********** -->
<!-- ************************************************************ -->
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
<% else %>

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black;">

<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Cadastro de Produtos: Atribuição de Regras</span></td>
</tr>
</table>
<br>
<br>

<!-- ************   MENSAGEM  ************ -->
<div class='MtAviso' style="width:649px;font-weight:bold;border:1pt solid black;padding-top: 5px; padding-bottom: 5px;" align="center">
	<span style='margin:5px 2px 5px 2px;'>Dados atualizados com sucesso</span>
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
<table class="notPrint" width="649" cellspacing="0">
<tr>
	<td align="center"><div name="dVOLTAR" id="dVOLTAR"><a name="bVOLTAR" id="bVOLTAR" href="MenuCadastro.asp" title="Retornar para o menu de cadastros">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a></div>
	</td>
</tr>
</table>
<% end if %>
</center>
</body>
</html>


<%
'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>