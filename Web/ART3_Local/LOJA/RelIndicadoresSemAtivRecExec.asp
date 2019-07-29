<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=False %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =================================================================
'	  RelIndicadoresSemAtivRecExec.asp
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
	dim cn, rs, msg_erro
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

	
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))
'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
    if Not operacao_permitida(OP_LJA_REL_INDICADORES_SEM_ATIVIDADE_RECENTE, s_lista_operacoes_permitidas) then
      Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
    end if

	dim alerta
	dim s, s_aux, s_filtro 
	dim c_vendedor, c_indicador
	dim c_loja, lista_loja, s_filtro_loja, v, i,loja
    dim v_vendedor,aviso
    
    v_vendedor = ""
	alerta = ""
	c_vendedor = Trim(Request.Form("c_vendedor"))
	c_indicador = Trim(Request.Form("c_indicador"))
	if loja = "" then loja = Session("loja_atual")
    v_vendedor = split(c_vendedor, ", ")


' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________

' _____________________________________
' CONSULTA EXECUTA
'
sub consulta_executa
dim r,s_sql,cont
dim n_reg,s_where_temp,s_where

n_reg = 0
s_where = ""

    if c_vendedor <> "" then
		if c_vendedor <> "" then s_where = s_where & " AND"
            s_where_temp = ""
        for cont = LBound(v_vendedor) to UBound(v_vendedor)
            if s_where_temp <> "" then s_where_temp = s_where_temp & " OR"
            s_where_temp = s_where_temp & " (vendedor = '" & Trim(replace(v_vendedor(cont), "'", "''")) & "')"
        next
        if s_where_temp <> "" then
            s_where_temp = "(" & s_where_temp & ")"
            s_where = s_where & s_where_temp
        end if
	end if
   
    
	s = " SELECT *" & _
        " FROM (" & _
	        " SELECT DATEDIFF(day, Coalesce(entregue_data, vendedor_dt_ult_atualizacao, dt_cadastro), getdate()) AS qtde_dias," & _
		        " pedido," & _
		        " apelido," & _
		        " razao_social_nome," & _
		        " vendedor," & _
                " vendedor_dt_ult_atualizacao," & _
                " dt_cadastro," & _
                " (SELECT TOP 1 dt_cadastro FROM t_ORCAMENTISTA_E_INDICADOR_BLOCO_NOTAS WHERE apelido=t2.apelido ORDER BY dt_hr_cadastro DESC) AS dt_ult_bloco_notas," & _
                " (SELECT TOP 1 mensagem FROM t_ORCAMENTISTA_E_INDICADOR_BLOCO_NOTAS WHERE apelido=t2.apelido ORDER BY dt_hr_cadastro DESC) AS msg_bloco_notas" & _
	        " FROM (" & _
		        " SELECT (" & _
				        " SELECT entregue_data" & _
				        " FROM t_PEDIDO" & _
				        " WHERE t_PEDIDO.pedido = t.pedido" & _
				        " ) AS entregue_data," & _
			        " t.pedido," & _
			        " t.apelido," & _
			        " t.vendedor_dt_ult_atualizacao," & _
			        " t.dt_cadastro," & _
			        " t.razao_social_nome," & _
			        " t.vendedor" & _
		        " FROM (" & _
			        " SELECT (" & _
					        " SELECT TOP 1 pedido" & _
					        " FROM t_PEDIDO" & _
					        " WHERE st_entrega = '" & ST_ENTREGA_ENTREGUE & "'" & _						  
						        " AND indicador = t_ORCAMENTISTA_E_INDICADOR.apelido" & _
					        " ORDER BY entregue_data DESC," & _
						        " data_hora DESC" & _
					        " ) AS pedido," & _
				        " t_ORCAMENTISTA_E_INDICADOR.apelido," & _
				        " t_ORCAMENTISTA_E_INDICADOR.vendedor_dt_ult_atualizacao," & _
				        " t_ORCAMENTISTA_E_INDICADOR.dt_cadastro," & _
				        " t_ORCAMENTISTA_E_INDICADOR.razao_social_nome," & _
				        " t_ORCAMENTISTA_E_INDICADOR.vendedor" & _
			        " FROM t_ORCAMENTISTA_E_INDICADOR" & _
			        " WHERE  STATUS = 'A'" & _
				         s_where & _
			        " ) t" & _
		        " ) t2" & _
	        " ) t3" & _
        " WHERE (" & _
		        " (qtde_dias >= 45)" & _
		        " OR (qtde_dias IS NULL)" & _
		        " )" & _
        " ORDER BY CASE " & _
		        " WHEN qtde_dias IS NULL" & _
			        " THEN 999999" & _
		        " ELSE qtde_dias" & _
		        " END DESC"


	set r = cn.Execute(s)
  
    s = "<table width='600' class='QS' cellSpacing='0' style='border-left:0px;'>" & chr(13) & _
		        "<tr style='background:#FFF0E0'>" & chr(13) & _
                "<td style='background:#FFF' class='MD MB'> &nbsp; </td>" & chr(13) & _
		        "<td class='MD MTB' align='right'><p class='R'>QTDE <br> DIAS</p></td>" & chr(13) & _
		        "<td class='MD MTB' align='center'><p class='R'>RELAC</p></td>" & chr(13) & _
		        "<td class='MD MTB' align='left'><p class='R'>PEDIDO</p></td>" & chr(13) & _
                "<td class='MD MTB' align='left'><p class='R'>VENDEDOR</p></td>" & chr(13) & _
		        "<td class='MD MTB' align='left'><p class='R'>APELIDO</p></td>" & chr(13) & _
                "<td class='MD MTB' align='left'><p class='R'>NOME INDICADOR</p></td>" & chr(13) & _
                "<td class='MTB' align='left'><p class='R'>DATA DE<br>CADASTRO</p></td>" & chr(13) & _
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
		    s = s & "	<td class='tdQtde' nowrap><p class='C'>" & r("qtde_dias") & "</p></td>" & chr(13)				
		    s = s & "	<td class='tdQtde' align='center' nowrap><p class='C' style='cursor:pointer' title='" & Trim("" & r("msg_bloco_notas")) & "'>" & formata_data(Trim("" & r("dt_ult_bloco_notas"))) & "</p></td>" & chr(13)		     
            s = s & "   <td  class='tdPed'nowrap><p class='C'>&nbsp;<a href='javascript:fRELCon(" & _
			chr(34) & r("pedido") & chr(34) & _
			")' title='clique para consultar o pedido' style='color:black'>" & r("pedido") & "</a></p></td>" & chr(13)
            s = s & "	<td class='tdVen'  nowrap><p class='C'>" & iniciais_em_maiusculas(r("vendedor")) & "</p></td>" & chr(13)
		    s = s & "	<td class='tdAp'  nowrap><p class='C'><a href='javascript:fOrcamentistaEIndicadorConsultaView(" & _
            chr(34) & r("apelido") & chr(34) & "," & chr(34) & usuario & chr(34) & _
            ")' title='clique para consultar o cadastro do indicador' style='color:black'>" & iniciais_em_maiusculas(r("apelido")) & "</a></p></td>" & chr(13)
            s = s & "	<td class='tdInd'  nowrap><p class='C'><a href='javascript:fOrcamentistaEIndicadorConsultaView(" & _
            chr(34) & r("apelido") & chr(34) & "," & chr(34) & usuario & chr(34) & _
            ")' title='clique para consultar o cadastro do indicador' style='color:black'>" & iniciais_em_maiusculas(r("razao_social_nome")) & "</a></p></td>" & chr(13)
            s = s & "	<td class='tdData' nowrap><p class='C'>" & formata_data(r("dt_cadastro")) & "</p></td>" & chr(13)
		    s = s & "</tr>" & chr(13)
                
            cont = cont + 1
		    r.MoveNext
	    loop
    end if
    if i = 0 then
	s = s & "<tr nowrap style='background: #FFF0E0;'>" & _
                            "<td style='background: #FFF;' class='MD ME'><p class='Rd'>0.</p></td>" & _
				            "<td align='center' colspan='7'>" & _
				            "<p class='C' style='color:red;letter-spacing:1px;'>NENHUM INDICADOR ENCONTRADO.</p>" & _
				            "</td></tr>"
	end if

	if n_reg > 0 and c_vendedor <> "" then
        s_sql = "SELECT COUNT('apelido') As n_ind  FROM t_ORCAMENTISTA_E_INDICADOR WHERE (status='A') AND " & s_where_temp
        if rs.State <> 0 then rs.Close
        rs.Open s_sql,cn
        s = s & "<tr nowrap style='background: #FFF;'>" & _                           
				            "<td align='center' colspan='8' class='MTE'>" & _
				            "<p class='C' style='letter-spacing:1px;'>TOTAL DE INDICADORES NA CARTEIRA: "& rs("n_ind") &"</p>" & _
				            "</td>" & _
                "</tr>"
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
	<title>LOJA</title>
	</head>

<% if False then 'APENAS P/ HABILITAR O INTELLISENSE DURANTE O DESENVOLVIMENTO!! %>
<script src="../Global/jquery.js" language="JavaScript" type="text/javascript"></script>
<% end if %>

<script src="<%=URL_FILE__JQUERY%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>

<script type="text/javascript">$(function () { $(document).tooltip(); });</script>

<script language="JavaScript" type="text/javascript">
var windowScrollTopAnterior;
window.status = 'Aguarde, executando a consulta ...';
$(document).ready(function () {
    $("#divOrcamentistaEIndicadorConsultaView").hide();
    $('#divInternoOrcamentistaEIndicadorConsultaView').addClass('divFixo');
    sizeDivOrcamentistaEIndicadorConsultaView();

    $("#divOrcamentistaEIndicadorConsultaView").click(function () {
        fechaDivOrcamentistaEIndicadorConsultaView();
    });
    $("#imgFechaDivOrcamentistaEIndicadorConsultaView").click(function () {
        fechaDivOrcamentistaEIndicadorConsultaView();
    });
});
//Every resize of window
$(window).resize(function () {
    sizeDivOrcamentistaEIndicadorConsultaView();
});
function fRELCon(id_pedido) {
    window.status = "Aguarde ...";
    fREL.pedido_selecionado.value = id_pedido;
    fREL.action = "pedido.asp"
    fREL.submit();
}
function fOrcamentistaEIndicadorConsultaView(apelido, usuario) {
    sizeDivOrcamentistaEIndicadorConsultaView();
    $("#iframeOrcamentistaEIndicadorConsultaView").attr("src", "OrcamentistaEIndicadorConsultaView.asp?id_selecionado=" + encodeURIComponent(apelido) + "&usuario=" + usuario);
    $("#divOrcamentistaEIndicadorConsultaView").fadeIn();
}
function fechaDivOrcamentistaEIndicadorConsultaView() {
    $("#divOrcamentistaEIndicadorConsultaView").fadeOut();
    $("#iframeOrcamentistaEIndicadorConsultaView").attr("src", "");
}
function sizeDivOrcamentistaEIndicadorConsultaView() {
    var newHeight = $(document).height() + "px";
    $("#divOrcamentistaEIndicadorConsultaView").css("height", newHeight);
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
<link href="<%=URL_FILE__JQUERY_UI_CSS %>" rel="Stylesheet" type="text/css" />

<style type="text/css">

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
.trCor{
    background: #FFF0E0
}
.tdn_reg{
    background:#FFF;
    text-align:right;
    border-right: 1pt solid #C0C0C0;
    border-left: 1pt solid #C0C0C0;
}
.tdQtde{
    width:60px;
    text-align:right;
    vertical-align:top;
    border-right: 1pt solid #C0C0C0;
}
.tdPed{
    text-align:left;
    vertical-align:top;
    border-right: 1pt solid #C0C0C0;
}
.tdVen{
    width:80px;
    text-align:left;
    vertical-align:top;
    border-right: 1pt solid #C0C0C0;
}
.tdAp{
    width:80px;
    text-align:left;
    vertical-align:top;
    border-right: 1pt solid #C0C0C0;
}	  
.tdInd{
    width:80px;
    text-align:left;
    vertical-align:top;
    border-right: 1pt solid #C0C0C0;
}
.tdData{
    border-right:0px;
    text-align:center;   
}	  	     
#iframeOrcamentistaEIndicadorConsultaView
{
	position:absolute;
	top:0;
	left:0;
	width:100%;
	height:100%;
	border: solid 4px black;
}
#imgFechaDivOrcamentistaEIndicadorConsultaView
{
    position:fixed;
	top:6%;
	left: 50%;
	margin-left: -16px; /* -1 * image width / 2 */
	margin-top: -32px;
	z-index:1001;
}
#divInternoOrcamentistaEIndicadorConsultaView.divFixo
{
    position:fixed;
	top:6%;
}
#divInternoOrcamentistaEIndicadorConsultaView
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
#divOrcamentistaEIndicadorConsultaView
{
    position:absolute;
	top:0;
	left:0;
	width:100%;
	z-index:1000;
	background-color:#808080;
	opacity: 1;
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

<input type="hidden" name="c_usuario_sessao" id="c_usuario_sessao" value="<%=usuario%>" />


<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="709" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Relatório de Indicadores sem Atividade Recente</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>

<!-- FILTROS -->
<%
	s_filtro = "<table width='709' cellpadding='0' cellspacing='0' style='border-bottom:1px solid black' border='0'>" & chr(13)

'	VENDEDOR
	if c_vendedor <> "" then
		s = c_vendedor
		s_aux = x_usuario(c_vendedor)
		if (s <> "") And (s_aux <> "") then s = s & " - "
		s = s & s_aux
    else
        s = "Todos"
    end if
		s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='top' nowrap><span class='N'>Vendedor(es):&nbsp;</span></td>" & chr(13) & _
					"		<td align='left' valign='top' width='99%'><span class='N'>" & s & "</span></td>" & chr(13) & _
					"	</tr>" & chr(13)

'	LOJA(S)
	s_filtro_loja = loja
	
	s = s_filtro_loja
	if s = "" then s = "todas"
	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' nowrap><span class='N'>Loja(s):&nbsp;</span></td>" & chr(13) & _
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
	<td align="center"><a name="bVOLTA" id="bVOLTA" href="RelIndicadoresSemAtivRec.asp">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a>
	</td>
</tr>
</table>


</form>

</center>

<div id="divOrcamentistaEIndicadorConsultaView"><center><div id="divInternoOrcamentistaEIndicadorConsultaView"><img id="imgFechaDivOrcamentistaEIndicadorConsultaView" src="../imagem/close_button_32.png" title="clique para fechar o painel de consulta" /><iframe id="iframeOrcamentistaEIndicadorConsultaView"></iframe></div></center></div>
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
