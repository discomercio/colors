<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=False %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =================================================================
'	  R E L C O M I S S A O I N D I C A D O R E S P A G E X E C . A S P
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
	
	dim usuario,loja
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	CONECTA COM O BANCO DE DADOS
	dim cn, rs, msg_erro
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'   VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	if Not operacao_permitida(OP_CEN_REL_PEDIDOS_CANCELADOS , s_lista_operacoes_permitidas) then 
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	dim alerta
	dim s, s_aux, s_filtro
	dim ckb_st_entrega_cancelado, c_dt_entregue_mes, c_dt_entregue_ano, str_data, mes, ano
	dim ckb_comissao_paga_sim, ckb_comissao_paga_nao
	dim ckb_st_pagto_pago, ckb_st_pagto_nao_pago, ckb_st_pagto_pago_parcial
	dim c_vendedor, c_indicador
	dim c_loja, lista_loja, s_filtro_loja, v_loja, v, i
	dim rb_visao, blnVisaoSintetica
    dim v_vendedor, vendedor_temp, j, aviso,c_dt_cancel_termino,c_dt_cancel_inicio,c_motivo,c_cod_sub_motivo
    
    v_vendedor = ""
	alerta = ""

	ckb_st_entrega_cancelado = Trim(Request.Form("ckb_st_entrega_entregue"))
	c_dt_cancel_inicio = Trim(Request.Form("c_dt_cancel_inicio"))
  '  c_dt_cancel_inicio = bd_formata_data(c_dt_cancel_inicio)
	c_dt_cancel_termino = Trim(Request.Form("c_dt_cancel_termino"))
   ' c_dt_cancel_termino = bd_formata_data(c_dt_cancel_termino)
	c_vendedor = Trim(Request.Form("c_vendedor"))
	c_indicador = Trim(Request.Form("c_indicador"))
    c_motivo = Trim(Request.Form("c_motivo"))
	c_loja = Trim(Request.Form("c_loja"))
    c_cod_sub_motivo = Trim(Request.Form("c_cod_sub_motivo"))

    if alerta = "" then
        if c_loja <> "" then
            s = "SELECT loja FROM t_PRODUTO_LOJA WHERE (loja='" & c_loja & "')"
            if rs.State <> 0 then rs.Close
			rs.open s, cn
            if rs.Eof then
                alerta = "Número da loja incorreto"
            end if       
        end if
    end if
        
' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________

' _____________________________________
' CONSULTA EXECUTA
'

sub consulta_executa
dim r,s_sql,cont,vRelat()
dim v_apelido(),v_indicador(),v_pedido(),v_qtde(),n_reg,s_where_temp,s_where,v_vendedores(),v_data()
dim num_vend

num_vend = 0
n_reg = 0
s_where = ""

    
		if c_vendedor <> "" then                         
            s_where = s_where & " AND ( vendedor = '" & c_vendedor & "')"           
        end if

		if c_motivo <> "" then         
            s_where = s_where & " AND ( codigo = '" & c_motivo & "')"           
        end if
        
        if c_loja <> "" then         
            s_where = s_where & " AND ( loja = '" & c_loja & "')"             
        end if
        
        if c_cod_sub_motivo <> "" then
           s_where = s_where & " AND ( cancelado_codigo_sub_motivo = '" & c_cod_sub_motivo & "')"    
        end if

        
	    s_where = s_where & " AND ( grupo = '" & GRUPO_T_CODIGO_DESCRICAO__CANCELAMENTOPEDIDO_MOTIVO & "')"
        s_where = s_where & " AND ( cancelado_data >= " & bd_formata_data(c_dt_cancel_inicio) & ")"  
        s_where = s_where & " AND ( cancelado_data <= " & bd_formata_data(c_dt_cancel_termino) & ")" 
        
    
	s = "SELECT pedido" & _
	        " ,t_CODIGO_DESCRICAO.descricao" & _
	        " ,cancelado_usuario" & _
	        " ,cancelado_data_hora" & _
            " ,vendedor" & _
            " ,nome" & _
            " ,Coalesce((SELECT descricao FROM t_CODIGO_DESCRICAO WHERE grupo = 'CancelamentoPedido_Motivo_Sub' AND codigo_pai = t_PEDIDO.cancelado_codigo_motivo AND codigo=t_PEDIDO.cancelado_codigo_sub_motivo),'' ) As descricao_Sub" & _
        " FROM t_PEDIDO" & _
        " INNER JOIN t_CLIENTE ON (t_PEDIDO.id_cliente = t_CLIENTE.id)" & _
        " INNER JOIN t_CODIGO_DESCRICAO ON (t_PEDIDO.cancelado_codigo_motivo = t_CODIGO_DESCRICAO.codigo)" & _
        " WHERE st_entrega = '" & ST_ENTREGA_CANCELADO & "'" & _
	        s_where & _
        " ORDER BY cancelado_data_hora DESC"

    
	set r = cn.Execute(s)
  
    s = "<table width='700' class='QS' cellSpacing='0' style='border-left:0px;'>" & chr(13) & _
		        "<tr style='background:#FFF0E0'>" & chr(13) & _
                "<td style='background:#FFF' class='MD MB'> &nbsp; </td>" & chr(13) & _
		        "<td class='MD MTB' align='left'><p class='R'>DATA / HORA </p></td>" & chr(13) & _
		        "<td class='MD MTB' align='left'><p class='R'>PEDIDO</p></td>" & chr(13) & _
                "<td class='MD MTB' align='left'><p class='R'>NOME DO CLIENTE</p></td>" & chr(13) 
    if operacao_permitida(OP_LJA_CONSULTA_UNIVERSAL_PEDIDO_ORCAMENTO , s_lista_operacoes_permitidas) then  
        s = s +"<td class='MD MTB' align='left'><p class='R'>VENDEDOR</p></td>" & chr(13) 
    end if  
             
	s = s +     "<td class='MD MTB' align='left'><p class='R'>USUÁRIO </p></td>" & chr(13) & _
                "<td class=' MD MTB' align='left'><p class='R'>MOTIVO</p></td>" & chr(13) & _
                "<td class=' MTB' align='left'><p class='R'>SUB-MOTIVO</p></td>" & chr(13) & _                 
		        "</tr>"
	i = 0
    cont = 0
    if not r.Eof then
	    do while Not r.Eof 

            n_reg = n_reg + 1  
            
           
            i = i + 1            
		    if (i AND 1)=0 then
			    s = s & "<tr nowrap class='trCor'>"
		    else
			    s = s & "<tr nowrap>"
			end if
            
            s = s & "	<td class='tdn_reg' nowrap><p class='Rd'>" & n_reg & ".</p></td>" & chr(13)
		    s = s & "	<td class='tdData' nowrap><p class='C'>" & formata_data_hora_sem_seg(r("cancelado_data_hora")) & "</p></td>" & chr(13)		     
            s = s & "   <td  class='tdPed'nowrap><p class='C'>&nbsp;<a href='javascript:fRELCon(" & _
			chr(34) & r("pedido") & chr(34) & _
			")' title='clique para consultar o pedido'>" & r("pedido") & "</a></p></td>" & chr(13)
            s = s & "	<td class='tdCliente' ><p class='C'>" & r("nome") & "</p></td>" & chr(13)           
            s = s & "<td class='tdVen'  ><p class='C'>" & iniciais_em_maiusculas( r("vendedor") ) & "</p></td>" & chr(13)
		    s = s & "	<td class='tdUsu'  nowrap><p class='C'>" & iniciais_em_maiusculas( r("cancelado_usuario")) & "</p></td>" & chr(13)
            s = s & "	<td class='tdMot'  ><p class='C'>" & iniciais_em_maiusculas( r("descricao")) & "</p></td>" & chr(13)
            s = s & "	<td class='tdSubM'  ><p class='C'>" & iniciais_em_maiusculas( r("descricao_Sub")) & "</p></td>" & chr(13)           
		    s = s & "</tr>" & chr(13)
                
            cont = cont + 1
		    r.MoveNext
	    loop
    end if
    if i = 0 then
	s = s & "<tr nowrap style='background: #FFF0E0;'>" & _
                            "<td style='background: #FFF;' class='MD ME'><p class='Rd'>0.</p></td>" & _
				            "<td align='center' colspan='7'>" & _
				            "<p class='C' style='color:red;letter-spacing:1px;'>NENHUM PEDIDO.</p>" & _
				            "</td></tr>"
	end if

	
      
        
     	
	s = s & "</table>" & chr(13)
	Response.Write(s)
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

<script language="JavaScript" type="text/javascript">
var windowScrollTopAnterior;
window.status = 'Aguarde, executando a consulta ...';

$(function() {
	$("#divPedidoConsulta").hide();

	sizeDivPedidoConsulta();

	$('#divInternoPedidoConsulta').addClass('divFixo');

	$(document).keyup(function(e) {
		if (e.keyCode == 27) fechaDivPedidoConsulta();
	});

	$("#divPedidoConsulta").click(function() {
		fechaDivPedidoConsulta();
	});

	$("#imgFechaDivPedidoConsulta").click(function() {
		fechaDivPedidoConsulta();
	});

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

function fRELCon(id_pedido) {
    window.status = "Aguarde ...";
    fRELConsulta.pedido_selecionado.value = id_pedido;
    fRELConsulta.action = "pedido.asp"
    fRELConsulta.submit();
}

function fRELGravaDados(f) {
	window.status = "Aguarde ...";
	bCONFIRMA.style.visibility = "hidden";
	f.action = "RelComissaoIndicadoresPagExecConfirma.asp";

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
.trCor{
    background: #FFF0E0
}
.tdn_reg{
    width:10px;
    background:#FFF;
    text-align:right;
    border-right: 1pt solid #C0C0C0;
    border-left: 1pt solid #C0C0C0;
}
.tdData{
    width:60px;
    text-align:left;
    vertical-align:top;
    border-right: 1pt solid #C0C0C0;
}
.tdPed{
    text-align:left;
    vertical-align:top;
    border-right: 1pt solid #C0C0C0;
}
.tdVen{
    width:60px;
    text-align:left;
    vertical-align:top;
    border-right: 1pt solid #C0C0C0;
}
.tdUsu{
    width:70px;
    text-align:left;
    vertical-align:top;
    border-right: 1pt solid #C0C0C0;
}	  
.tdMot{
    width:100px;
    text-align:left;
    vertical-align:top;
    border-right: 1pt solid #C0C0C0;
}
.tdSubM{
    width:80px;
    text-align:left;
    vertical-align:top;

}
.tdReg{
    width:5%;
    background:#FFF;
    text-align:right;
    border-right: 1pt solid #C0C0C0;
    border-left: 1pt solid #C0C0C0;
}	 
.tdDataF{
    width:10%;
    text-align:left;
    vertical-align:top;
    border-right: 1pt solid #C0C0C0;
}
.tdAnalise{
    width:35%;
    text-align:left;
    vertical-align:top;
} 	     
.tdPedido{
    width:10%;
    text-align:left;
    vertical-align:top;
    border-right: 1pt solid #C0C0C0;
}
.tdCliente{    
    max-width:400px;
    text-align:left;
    vertical-align:top;
    border-right: 1pt solid #C0C0C0;
    white-space: normal;
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

<form id="fRELConsulta" name="fRELConsulta" method="post">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="pedido_selecionado" id="pedido_selecionado" value="">

<input type="hidden" name="mes" id="mes" value="<%=c_dt_entregue_mes%>">
<input type="hidden" name="ano" id="ano" value="<%=c_dt_entregue_ano%>">
<input type="hidden" name="c_usuario_sessao" id="c_usuario_sessao" value="<%=usuario%>" />
<input type="hidden" name="rb_visao" id="rb_visao" value="<%=rb_visao%>" />

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="709" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Relatório de Pedidos Cancelados</span>
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
	if (c_dt_cancel_inicio <> "") AND (c_dt_cancel_termino <> "") then
		s = formata_data(c_dt_cancel_inicio) & " e " & formata_data(c_dt_cancel_termino)
		end if

	if s <> "" then
		s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='top' nowrap><span class='N'>Cancelado(s) entre:&nbsp;</span></td>" & chr(13) & _
					"		<td align='left' valign='top' width='99%'><span class='N'>" & s & "</span></td>" & chr(13) & _
					"	</tr>" & chr(13)
		end if

'	VENDEDOR
	if c_vendedor <> "" then
		s = c_vendedor
		s_aux = x_usuario(c_vendedor)
		if (s <> "") And (s_aux <> "") then s = s & " - "
		s = s & s_aux
		s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='top' nowrap><span class='N'>Vendedor:&nbsp;</span></td>" & chr(13) & _
					"		<td align='left' valign='top' width='99%'><span class='N'>" & s & "</span></td>" & chr(13) & _
					"	</tr>" & chr(13)
		end if

'	LOJA
	s = ""
    if c_loja <> "" then 
     s = s & c_loja
    end if
	if s = "" then s = "todas"
	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' nowrap><span class='N'>Loja:&nbsp;</span></td>" & chr(13) & _
				"		<td align='left' valign='top'><span class='N'>" & s & "</span></td>" & chr(13) & _
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
	<td align="center"><a name="bVOLTA" id="bVOLTA" href="javascript:history.back()" title="volta para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a>
	</td>
	<td align="left">&nbsp;</td>
</tr>
</table>


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
