<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp"        -->
<!-- #include file = "../global/Global.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =============================================================
'	  EstoqueTransfereEntreCDsConsultaFiltro.asp
'     =============================================================
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

	Const COD_TRANSF_CADASTRA = "TRANSF_CADASTRA"
	Const COD_TRANSF_CONSULTA = "TRANSF_CONSULTA"
	
	dim usuario
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	CONECTA COM O BANCO DE DADOS
	dim cn
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

	dim url_back, strUrlBotaoVoltar
	url_back = Trim(Request("url_back"))
	if url_back <> "" then
		strUrlBotaoVoltar = "Resumo.asp?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))
	else
		strUrlBotaoVoltar = "javascript:history.back()"
		end if

'	CD
	dim i, qtde_nfe_emitente
	dim v_usuario_x_nfe_emitente
	dim id_nfe_emitente_selecionado
	v_usuario_x_nfe_emitente = obtem_lista_usuario_x_nfe_emitente(usuario)
	
	qtde_nfe_emitente = 0
	for i=Lbound(v_usuario_x_nfe_emitente) to UBound(v_usuario_x_nfe_emitente)
		if Not Isnull(v_usuario_x_nfe_emitente(i)) then
			qtde_nfe_emitente = qtde_nfe_emitente + 1
			id_nfe_emitente_selecionado = v_usuario_x_nfe_emitente(i)
			end if
		next
	
	if qtde_nfe_emitente > 1 then
	'	HÁ MAIS DO QUE 1 CD, ENTÃO SERÁ EXIBIDA A LISTA P/ O USUÁRIO SELECIONAR UM CD
		id_nfe_emitente_selecionado = 0
		end if
	
	if qtde_nfe_emitente = 0 then
	'	NÃO HÁ NENHUM CD CADASTRADO P/ ESTE USUÁRIO!!
		Response.Redirect("aviso.asp?id=" & ERR_NENHUM_CD_HABILITADO_PARA_USUARIO)
		end if



' ____________________________________________________________________________
' TRANSF NSU MONTA ITENS SELECT
'
function t_transf_nsu_monta_itens_select()
dim x, r, strResp, strSql, v, i

    strResp = ""

	strSql = "select id" & _
                   " from t_ESTOQUE_TRANSFERENCIA" & _
                   " where st_exclusao <> 1" & _
                   " order by id"

	set r = cn.Execute(strSql)
	x = ""
	do while Not r.eof 
	    
		x = CStr("" & r("id"))
        While Len(x) < 6: x = "0" & x: Wend 
		strResp = strResp & "<option "
		strResp = strResp & " value=" & x & ">"
		strResp = strResp & x        
        'if Trim("" & r("descricao")) <> "" then strResp = strResp & "&nbsp;-&nbsp;" & Trim("" & r("descricao"))
		strResp = strResp & "</option>" & chr(13)
		r.MoveNext	
 	loop
		
	t_transf_nsu_monta_itens_select = strResp
	r.close
	set r=nothing
end function



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
<script src="<%=URL_FILE__JQUERY_UI%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_I18N%>" Language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>

<script type="text/javascript">
	$(function() {
		$("input[type=radio]").hUtil('fix_radios');
		$("#c_filtro_dt_entrega").hUtilUI('datepicker_padrao');

		if ($("input[name='rb_op_desejada']:checked").val() == '<%=COD_TRANSF_CONSULTA%>') {
			$(".TR_LISTA_NSU").show();
		}
		else {
			$(".TR_LISTA_NSU").hide();
		}

		$("input[name='rb_op_desejada']").change(function() {
			if ($("input[name='rb_op_desejada']:checked").val() == '<%=COD_TRANSF_CONSULTA%>') {
				$(".TR_LISTA_NSU").show();
			}
			else {
				$(".TR_LISTA_NSU").hide();
			}
		});
	});
</script>

<script language="JavaScript" type="text/javascript">
    function limpaCampoSelectNSU() {
        $("#c_nsu").children().prop("selected", false);
    }

function fFILTROConfirma( f ) {
    var s_ir_para;

	if (!f.rb_op_desejada[0].checked && !f.rb_op_desejada[1].checked) {
		alert("Selecione o tipo de relatório!!");
		return;
	}

    if (f.rb_op_desejada[0].checked) {
        s_ir_para = "EstoqueTransfereEntreCDs.asp";
	}

    if (f.rb_op_desejada[1].checked) {
        if (f.c_nsu.value == "") {
            alert("Selecione uma transferência!!!");
            return;
        }
        s_ir_para = "EstoqueTransfereEntreCDsConsulta.asp?transf_selecionada=" + f.c_nsu.value;
	}

    f.action = s_ir_para;
	dCONFIRMA.style.visibility="hidden";
	window.status = "Aguarde ...";
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

<link href="<%=URL_FILE__E_CSS%>" Rel="stylesheet" Type="text/css">
<link href="<%=URL_FILE__JQUERY_UI_CSS%>" rel="stylesheet" type="text/css">


<body>
<center>

<form id="fFILTRO" name="fFILTRO" method="post" action="RelSolicitacaoColetasExec.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<% if qtde_nfe_emitente = 1 then %>
<input type="hidden" name="c_nfe_emitente" id="c_nfe_emitente" value="<%=Cstr(id_nfe_emitente_selecionado)%>" />
<% end if %>


<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
    <td align="right" valign="bottom"><span class="PEDIDO">Transferência de Produtos Entre CD's</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>
<br>

<!--  PARÂMETROS  -->
<table class="Qx" cellspacing="0" cellpadding="2">
<tr>
<td class="MT" align="left">
<span class="PLTe">OPERAÇÃO DESEJADA</span>
<br />
	<table cellSpacing="0" style="margin-left:8px;margin-right:8px;">
	<tr>
		<td align="left"><input type="radio" tabindex="-1" id="rb_op_desejada" name="rb_op_desejada"
				value="<%=COD_TRANSF_CADASTRA%>" /><span class="C" style="cursor:default" onclick="fFILTRO.rb_op_desejada[0].click();">Nova Transferência</span>
		</td>
	</tr>
	<tr>
		<td align="left"><input type="radio" tabindex="-1" id="rb_op_desejada" name="rb_op_desejada"
				value="<%=COD_TRANSF_CONSULTA%>" /><span class="C" style="cursor:default" onclick="fFILTRO.rb_op_desejada[1].click();">Consultar Transferências</span>
		</td>
	</tr>
	</table>
</td>
</tr>
<!--  LISTA DE NSU's DAS TRANSFERÊNCIAS  -->
<tr class="TR_LISTA_NSU">
	<td class="ME MD PLTe" nowrap align="left" valign="bottom">&nbsp;Transferências</td>
</tr>
<tr class="TR_LISTA_NSU" bgcolor="#FFFFFF" nowrap>
	<td class="ME MB MD" style="padding-left:10px;" align="left">
		<table cellpadding="0" cellspacing="0" style="margin:1px 3px 6px 10px;">
		    <tr>
		    <td>
			    <select id="c_nsu" name="c_nsu" class="LST" size="5" style="width:200px" multiple>
			    <% =t_transf_nsu_monta_itens_select() %>
			    </select>
		    </td>
		    <td style="width:1px;"></td>
		    <td align="left" valign="top">
			    <a name="bLimparNsu" id="bLimparNsu" href="javascript:limpaCampoSelectNSU()" title="limpa o filtro 'NSU'">
						    <img src="../botao/botao_x_red.gif" style="vertical-align:bottom;margin-bottom:1px;" width="20" height="20" border="0"></a>
		    </td>
		    </tr>
		</table>
	</td>
</tr>
</table>

<!-- ************   SEPARADOR   ************ -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc" align="left">&nbsp;</td></tr>
</table>
<br>


<table width="649" cellSpacing="0">
<tr>
	<td align="left"><a name="bVOLTAR" id="bVOLTAR" href="<%=strUrlBotaoVoltar%>" title="volta para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
	<td align="right"><div name="dCONFIRMA" id="dCONFIRMA"><a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fFILTROConfirma(fFILTRO)" title="executa a consulta">
		<img src="../botao/confirmar.gif" width="176" height="55" border="0"></a></div>
	</td>
</tr>
</table>
</form>

</center>
</body>
</html>

<%
'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>
