<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/Global.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ========================================================
'	  RelPedidoOcorrenciaExec.asp
'     ========================================================
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

	dim usuario, loja
	usuario = Trim(Session("usuario_atual"))
	loja = Trim(Session("loja_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO)
	If (loja = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	CONECTA COM O BANCO DE DADOS
	dim cn, rs, msg_erro
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	if Not operacao_permitida(OP_LJA_OCORRENCIAS_EM_PEDIDOS_CADASTRAMENTO, s_lista_operacoes_permitidas) then 
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if
	
	dim s_filtro, intQtdeOcorrencias
	intQtdeOcorrencias = 0





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
				")' title='clique para consultar o pedido " & id_pedido & "'>" & _
				id_pedido & "</a>"
	monta_link_pedido=strLink
end function


' _____________________________________
' CONSULTA EXECUTA
'
sub consulta_executa
dim s, s_sql, x
dim r
dim cab_table, cab
dim qtde_ocorrencia_aberta, qtde_ocorrencia_em_andamento, qtde_ocorrencia_finalizada
dim s_link_rastreio

	s_sql = _
		"SELECT" & _
			" tPO.id," & _
			" tPO.pedido," & _
			" tPO.finalizado_status," & _
			" tPO.finalizado_data_hora," & _
			" tPO.usuario_cadastro," & _
			" tPO.dt_cadastro," & _
			" tPO.dt_hr_cadastro," & _
			" tPO.contato," & _
			" tPO.ddd_1," & _
			" tPO.tel_1," & _
			" tPO.ddd_2," & _
			" tPO.tel_2," & _
			" tPO.texto_ocorrencia," & _
			" tPO.tipo_ocorrencia," & _
            " tPO.cod_motivo_abertura," & _
			" tPO.texto_finalizacao," & _
			" tP.transportadora_id," & _
            " tP.loja AS pedido_loja," & _
			" tC.nome AS nome_cliente," & _
			" (" & _
				"SELECT" & _
					" TOP 1 NFe_numero_NF" & _
				" FROM t_NFe_EMISSAO tNE" & _
				" WHERE" & _
					" (tNE.pedido=tPO.pedido)" & _
					" AND (tipo_NF = '1')" & _
					" AND (st_anulado = 0)" & _
					" AND (codigo_retorno_NFe_T1 = 1)" & _
				" ORDER BY" & _
					" id DESC" & _
			") AS numeroNFe," & _
			" (" & _
				"SELECT" & _
					" Count(*)" & _
				" FROM t_PEDIDO_OCORRENCIA_MENSAGEM" & _
				" WHERE" & _
					" (id_ocorrencia=tPO.id)" & _
					" AND (fluxo_mensagem='" & COD_FLUXO_MENSAGEM_OCORRENCIAS_EM_PEDIDOS__CENTRAL_PARA_LOJA & "')" & _
			") AS qtde_msg_central," & _
            " (" & _
                " SELECT Count(*)" & _
		           " FROM t_PEDIDO_OCORRENCIA_MENSAGEM INNER JOIN t_PEDIDO_OCORRENCIA ON (t_PEDIDO_OCORRENCIA_MENSAGEM.id_ocorrencia=t_PEDIDO_OCORRENCIA.id)" & _
                   " INNER JOIN t_PEDIDO ON (t_PEDIDO_OCORRENCIA.pedido=t_PEDIDO.pedido)" & _ 
		           " WHERE (id_ocorrencia = tPO.id)" & _
                   " AND (t_PEDIDO.loja = '" & NUMERO_LOJA_ECOMMERCE_AR_CLUBE & "')" & _ 
            ") AS qtde_msg" & _
	   " FROM t_PEDIDO_OCORRENCIA tPO" & _
			" INNER JOIN t_PEDIDO tP ON (tPO.pedido=tP.pedido)" & _
			" INNER JOIN t_CLIENTE tC ON (tP.id_cliente=tC.id)" & _
	   " WHERE" & _
			" (tP.loja = '" & loja & "')" & _
			" AND " & _
				"(" & _
					"(tPO.usuario_cadastro = '" & usuario & "')" & _
					" OR " & _
					"(tP.vendedor = '" & usuario & "')" & _
				")" & _
			" AND " & _
				"(" & _
					"(tPO.finalizado_status = 0)" & _
					" OR " & _
					"( (tPO.finalizado_status <> 0) AND (tPO.finalizado_data >= DateAdd(day,-30,getdate())) )" & _
				")"
	
	s_sql = "SELECT * FROM (" & s_sql & ") t ORDER BY finalizado_status, finalizado_data_hora, dt_hr_cadastro, id"

	cab_table = "<TABLE cellSpacing=0 cellPadding=0>" & chr(13)
	cab = "	<TR style='background:azure' NOWRAP>" & chr(13) & _
		"		<TD class='MDTE tdDataHora' style='vertical-align:bottom'><P class='Rc'>DT Ocorr</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdPedido' style='vertical-align:bottom'><P class='R'>Pedido</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdNF' style='vertical-align:bottom'><P class='Rc'>NF</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdTransp' style='vertical-align:bottom'><P class='R'>Transp</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdOcorrencia' style='vertical-align:bottom'><P class='R'>Ocorrência</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdSolucao' style='vertical-align:bottom'><P class='R'>Solução</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdTipoOcorrencia' style='vertical-align:bottom'><P class='R'>Tipo Ocorrência</P></TD>" & chr(13) & _
		  "		<TD class='MTD tdStatus' style='vertical-align:bottom'><P class='R'>Status</P></TD>" & chr(13) & _
		  "		<TD style='background:white;'>&nbsp;</TD>" & chr(13) & _
		  "	</TR>" & chr(13)
	
	x = cab_table & cab
	intQtdeOcorrencias = 0
	qtde_ocorrencia_aberta = 0
	qtde_ocorrencia_em_andamento = 0
	qtde_ocorrencia_finalizada = 0
	
	set r = cn.execute(s_sql)
	do while Not r.Eof
	
	 ' CONTAGEM
		intQtdeOcorrencias = intQtdeOcorrencias + 1

		x = x & "	<TR NOWRAP>" & chr(13)

	'> DATA DA OCORRÊNCIA
		s = formata_data_hora_sem_seg(r("dt_hr_cadastro"))
		x = x & "		<TD class='MDTE tdDataHora'><P class='Cnc'>" & s & "</P></TD>" & chr(13)

	'> PEDIDO
		s = monta_link_pedido(Trim("" & r("pedido")))
		x = x & "		<TD class='MTD tdPedido'><P class='Cn'>" & s & "</P></TD>" & chr(13)

	'> NF
        s_link_rastreio = ""
		s = Trim("" & r("numeroNFe"))
		if s = "" then
            s = "&nbsp;"
        else
            s_link_rastreio = monta_link_rastreio(Trim("" & r("pedido")), Trim("" & r("numeroNFe")), Trim("" & r("transportadora_id")), Trim("" & r("pedido_loja")))
        end if
        if s_link_rastreio <> "" then s_link_rastreio = "&nbsp;" & s_link_rastreio
		x = x & "		<TD class='MTD tdNF'><P class='Cnc'>" & s & s_link_rastreio & "</P></TD>" & chr(13)

	'> TRANSPORTADORA
		s = Trim("" & r("transportadora_id"))
		if s = "" then s = "&nbsp;"
		x = x & "		<TD class='MTD tdTransp'><P class='Cn'>" & s & "</P></TD>" & chr(13)

	'> OCORRÊNCIA
        s = Trim("" & r("texto_ocorrencia"))
		if Trim("" & r("cod_motivo_abertura")) = "" then 
            x = x & "		<TD class='MTD tdOcorrencia'><P class='Cn'>" & s & "</P></TD>" & chr(13)
        else
			s = iniciais_em_maiusculas(obtem_descricao_tabela_t_codigo_descricao(GRUPO_T_CODIGO_DESCRICAO__OCORRENCIAS_EM_PEDIDOS__MOTIVO_ABERTURA, Trim("" & r("cod_motivo_abertura")))) & "<br>" & substitui_caracteres(Trim("" & r("texto_ocorrencia")), chr(13), "<br>")
            x = x & "		<TD class='MTD tdOcorrencia'><P class='Cn'>" & s & "</P></TD>" & chr(13)            
	    end if

	'> SOLUÇÃO
		s = substitui_caracteres(Trim("" & r("texto_finalizacao")), chr(13), "<br>")
		if s = "" then s = "&nbsp;"
		x = x & "		<TD class='MTD tdSolucao'><P class='Cn'>" & s & "</P></TD>" & chr(13)

	'> TIPO DE OCORRÊNCIA
		s = Trim("" & r("tipo_ocorrencia"))
		if s <> "" then 
			s = iniciais_em_maiusculas(obtem_descricao_tabela_t_codigo_descricao(GRUPO_T_CODIGO_DESCRICAO__OCORRENCIAS_EM_PEDIDOS__TIPO_OCORRENCIA, Trim("" & r("tipo_ocorrencia"))))
			end if
		if s = "" then s = "&nbsp;"
		x = x & "		<TD class='MTD tdTransp'><P class='Cn'>" & s & "</P></TD>" & chr(13)

	'> STATUS
		if CInt(r("finalizado_status")) <> 0 then
			s = "Finalizada"
			qtde_ocorrencia_finalizada = qtde_ocorrencia_finalizada + 1
		else
			if CInt(r("qtde_msg_central")) > 0 Or CInt(r("qtde_msg")) > 0 then
				s = "Em Andamento"
				qtde_ocorrencia_em_andamento = qtde_ocorrencia_em_andamento + 1
			else
				s = "Aberta"
				qtde_ocorrencia_aberta = qtde_ocorrencia_aberta + 1
				end if
			end if

		if s = "" then s = "&nbsp;"
		x = x & "		<TD class='MTD tdStatus'><P class='Cn'>" & s & "</P></TD>" & chr(13)

	'> BOTÃO P/ EXIBIR DEMAIS CAMPOS
		x = x & "		<TD valign='bottom' class='notPrint'>" & _
							"&nbsp;<a name='bExibeOcultaCampos' id='bExibeOcultaCampos' href='javascript:fExibeOcultaCampos(" & chr(34) & Cstr(intQtdeOcorrencias) & chr(34) & ")' title='exibe ou oculta os campos adicionais'><img src='../botao/view_bottom.png' border='0'></a>" & _
						"</TD>" & chr(13)
		
		x = x & "	</TR>" & chr(13)

	'> MENSAGENS
		s_sql = _
			"SELECT " & _
				"*" & _
		   " FROM t_PEDIDO_OCORRENCIA_MENSAGEM" & _
		   " WHERE" & _
				" (id_ocorrencia = " & Trim("" & r("id")) & ")" & _
		   " ORDER BY" & _
				" dt_hr_cadastro," & _
				" id"
		if rs.State <> 0 then rs.Close
		rs.open s_sql, cn
		x = x & "	<TR style='display:none;' id='TR_MSGS_" & Cstr(intQtdeOcorrencias) & "'>" & chr(13) & _
				"		<TD class='ME MD'>&nbsp;</TD>" & chr(13) & _
				"		<TD colspan='7' class='MC MD'>" & chr(13) & _
				"			<table width='100%' cellspacing='0' cellpadding='0'>" & chr(13) & _
				"				<TR>" & chr(13) & _
				"					<td class='Rf tdWithPadding'>MENSAGENS</td>" & chr(13) & _
				"				</TR>" & chr(13)
		if rs.Eof then
			x = x & _
				"				<TR>" & chr(13) & _
				"					<td>&nbsp;</td>" & chr(13) & _
				"				</TR>" & chr(13)
			end if
		
		do while Not rs.Eof
			x = x & _
				"				<TR>" & chr(13) & _
				"					<TD>" & chr(13) & _
				"						<table width='100%' cellSpacing='0' cellPadding='0'>" & chr(13) & _
				"							<TR>" & chr(13) & _
				"								<TD class='Cn MD MC tdWithPadding tdDataHoraMsg' align='center'>" & chr(13) & _
													formata_data_hora_sem_seg(rs("dt_hr_cadastro")) & _
				"								</TD>" & chr(13) & _
				"								<TD class='Cn MD MC tdWithPadding tdUsuarioMsg' align='center'>" & chr(13) & _
													rs("usuario_cadastro")
			if Trim("" & rs("loja")) <> "" then x = x & " (Loja&nbsp;" & Trim("" & rs("loja")) & ")"
			x = x & _
				"								</TD>" & chr(13) & _
				"								<TD class='Cn MC tdWithPadding tdTextoMensagem' align='left' valign='top'>" & chr(13) & _
													substitui_caracteres(Trim("" & rs("texto_mensagem")), chr(13), "<br>") & _
												"</TD>" & chr(13) & _
				"							</TR>" & chr(13) & _
				"						</table>" & chr(13) & _
				"					</TD>" & chr(13) & _
				"				</TR>" & chr(13)
			rs.MoveNext
			loop
		
		x = x & _
				"			</table>" & chr(13) & _
				"		</TD>" & chr(13) & _
				"	</TR>" & chr(13)

		if (intQtdeOcorrencias mod 100) = 0 then
			Response.Write x
			x = ""
			end if
		
		r.MoveNext
		loop
		
		
'	TOTAL GERAL
	if intQtdeOcorrencias > 0 then
		x = x & "	<TR>" & chr(13) & _
				"		<TD COLSPAN='8' class='MC' style='border-left:0px;border-right:0px;'>&nbsp;</TD>" & chr(13) & _
				"	</TR>" & chr(13) & _
				"	<TR NOWRAP style='background:honeydew'>" & chr(13) & _
				"		<TD COLSPAN='8' class='MT'><p class='C'>TOTAL: &nbsp; " & cstr(qtde_ocorrencia_aberta+qtde_ocorrencia_em_andamento+qtde_ocorrencia_finalizada) & " ocorrência(s)</p></TD>" & chr(13) & _
				"	</TR>" & chr(13)
		end if
		
  ' MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	if intQtdeOcorrencias = 0 then
		x = cab_table & cab
		x = x & "	<TR NOWRAP>" & chr(13) & _
				"		<TD class='MT' colspan='8'><P class='ALERTA'>&nbsp;NENHUM REGISTRO ENCONTRADO&nbsp;</P></TD>" & chr(13) & _
				"	</TR>" & chr(13)
		end if

  ' FECHA TABELA
	x = x & "</TABLE>" & chr(13)
	
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



<html>


<head>
	<title>LOJA</title>
	</head>



<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY%>" language="JavaScript" type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
window.status = 'Aguarde, executando a consulta ...';

function fExibeOcultaCampos(indice_row) {
var row_MSGS;

	row_MSGS = document.getElementById("TR_MSGS_" + indice_row);
	if (row_MSGS.style.display.toString() == "none") {
		row_MSGS.style.display = "";
	}
	else {
		row_MSGS.style.display = "none";
	}
}

function fPEDConsulta(id_pedido) {
	window.status = "Aguarde ...";
	fPED.pedido_selecionado.value = id_pedido;
	fPED.action = "pedido.asp"
	fPED.submit();
}
</script>
<script type="text/javascript">
    $(document).ready(function () {
        $("#divRastreioConsultaView").hide();
        $('#divInternoRastreioConsultaView').addClass('divFixo');
        sizeDivRastreioConsultaView();
        $(document).keyup(function (e) {
            if (e.keyCode == 27) {
                fechaDivRastreioConsultaView();
            }
        });
        $("#divRastreioConsultaView").click(function () {
            fechaDivRastreioConsultaView();
        });
        $("#imgFechaDivRastreioConsultaView").click(function () {
            fechaDivRastreioConsultaView();
        });
    });
    //Every resize of window
    $(window).resize(function () {
        sizeDivRastreioConsultaView();
    });
    function fRastreioConsultaView(url) {
        sizeDivRastreioConsultaView();
        $("#divRastreioConsultaView").fadeIn();
        frame = document.getElementById("iframeRastreioConsultaView");
        frame.contentWindow.location.replace(url);
    }
    function fechaDivRastreioConsultaView() {
        $("#divRastreioConsultaView").fadeOut();
    }
    function sizeDivRastreioConsultaView() {
        var newHeight = $(document).height() + "px";
        $("#divRastreioConsultaView").css("height", newHeight);
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
<link href="<%=URL_FILE__ESCREEN_CSS%>" Rel="stylesheet" Type="text/css" media="screen">

<style TYPE="text/css">
.tdWithPadding
{
	padding:1px;
}
.tdDataHora{
	vertical-align: top;
	width: 65px;
	}
.tdPedido{
	vertical-align: top;
	font-weight: bold;
	width: 65px;
	}
.tdNF{
	vertical-align: top;
	width: 60px;
	}
.tdTransp{
	vertical-align: top;
	width: 90px;
	}
.tdOcorrencia{
	vertical-align: top;
	width: 243px;
	}
.tdSolucao{
	vertical-align: top;
	width: 243px;
	}
.tdTipoOcorrencia{
	vertical-align: top;
	width: 139px;
	}
.tdStatus{
	vertical-align: top;
	width: 90px;
	}
.tdDataHoraMsg{
	vertical-align: top;
	width: 63px;
	}
.tdUsuarioMsg{
	vertical-align: top;
	width: 80px;
	}
.tdTextoMensagem{
	vertical-align: top;
	width: 785px;
	}
#divRastreioConsultaView
{
	position:absolute;
	top:0;
	left:0;
	width:100%;
	z-index:1000;
	background-color:#808080;
	opacity: 1;
}
#divInternoRastreioConsultaView
{
	position:absolute;
	top:6%;
	left:5%;
	width:90%;
	height:90%;
	z-index:1000;
	background-color:#fff;
	opacity: 1;
}
#divInternoRastreioConsultaView.divFixo
{
	position:fixed;
	top:6%;
}
#imgFechaDivRastreioConsultaView
{
	position:fixed;
	top:6%;
	left: 50%;
	margin-left: -16px; /* -1 * image width / 2 */
	margin-top: -32px;
	z-index:1001;
}
#iframeRastreioConsultaView
{
	position:absolute;
	top:0;
	left:0;
	width:100%;
	height:100%;
	border: solid 4px black;
}
</style>


<body onload="window.status='Concluído';focus();" link=#000000 alink=#000000 vlink=#000000>
<center>

<form id="fPED" name="fPED" method="post">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="pedido_selecionado" id="pedido_selecionado" value="">
</form>


<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="1024" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Ocorrências</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>

<!-- FILTROS -->
<% 
	s_filtro = "<table width='1024' cellPadding='0' CellSpacing='0' style='border-bottom:1px solid black' border='0'>" & chr(13)

	s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='top' NOWRAP>" & _
					"<p class='N'>Emissão:&nbsp;</p></td><td valign='top' width='99%'>" & _
					"<p class='N'>" & formata_data_hora(Now) & "</p></td>" & chr(13) & _
					"	</tr>" & chr(13)
	
	s_filtro = s_filtro & "</table>" & chr(13)
	Response.Write s_filtro
%>

<!--  RELATÓRIO  -->
<br>

<% consulta_executa %>

<!-- ************   SEPARADOR   ************ -->
<table width="1024" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>
<br>


<table class="notPrint" width="1024" cellSpacing="0">
<tr>
	<td align="center"><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()" title="volta para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a>
	</td>
</tr>
</table>

</center>

<div id="divRastreioConsultaView"><center><div id="divInternoRastreioConsultaView"><img id="imgFechaDivRastreioConsultaView" src="../imagem/close_button_32.png" title="clique para fechar o painel de consulta" /><iframe id="iframeRastreioConsultaView"></iframe></div></center></div>

</body>
</html>


<%
	if rs.State <> 0 then rs.Close
	set rs = nothing
	
'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>
