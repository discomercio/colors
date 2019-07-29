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
'	  R E L S O L I C I T A C A O C O L E T A S F I L T R O . A S P
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

	Const COD_TIPO_RELATORIO_SOLICITACAO_COLETA = "SOLICITACAO_COLETA"
	Const COD_TIPO_RELATORIO_PRONTO_PARA_ROMANEIO = "PRONTO_PARA_ROMANEIO"
	
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

		if ($("input[name='rb_tipo_relatorio']:checked").val() == '<%=COD_TIPO_RELATORIO_PRONTO_PARA_ROMANEIO%>') {
			$(".TR_DT_ENTREGA").show();
		}
		else {
			$(".TR_DT_ENTREGA").hide();
		}

		$("input[name='rb_tipo_relatorio']").change(function() {
			if ($("input[name='rb_tipo_relatorio']:checked").val() == '<%=COD_TIPO_RELATORIO_PRONTO_PARA_ROMANEIO%>') {
				$(".TR_DT_ENTREGA").show();
			}
			else {
				$(".TR_DT_ENTREGA").hide();
			}
		});
	});
</script>

<script language="JavaScript" type="text/javascript">
function limpaCampoTransp(f) {
	f.c_filtro_transportadora.options[0].selected = true;
}

function fFILTROConfirma( f ) {

	if (f.rb_loja[1].checked) {
		if (converte_numero(f.c_loja.value)==0) {
			alert("Especifique o número da loja!!");
			f.c_loja.focus();
			return;
			}
		}

	if (f.rb_loja[2].checked) {
		if (trim(f.c_loja_de.value)!="") {
			if (converte_numero(f.c_loja_de.value)==0) {
				alert("Número de loja inválido!!");
				f.c_loja_de.focus();
				return;
				}
			}
		if (trim(f.c_loja_ate.value)!="") {
			if (converte_numero(f.c_loja_ate.value)==0) {
				alert("Número de loja inválido!!");
				f.c_loja_ate.focus();
				return;
				}
			}
		if ((trim(f.c_loja_de.value)=="")&&(trim(f.c_loja_ate.value)=="")) {
			alert("Preencha pelo menos um dos campos!!");
			f.c_loja_de.focus();
			return;
			}
		if ((trim(f.c_loja_de.value)!="")&&(trim(f.c_loja_ate.value)!="")) {
			if (converte_numero(f.c_loja_ate.value)<converte_numero(f.c_loja_de.value)) {
				alert("Faixa de lojas inválida!!");
				f.c_loja_ate.focus();
				return;
				}
			}
		}

	if (!f.rb_tipo_relatorio[0].checked && !f.rb_tipo_relatorio[1].checked) {
		alert("Selecione o tipo de relatório!!");
		return;
	}

	if (trim(f.c_nfe_emitente.value) == "") {
		alert("É necessário selecionar um CD!!");
		return;
	}

	if (converte_numero(f.c_nfe_emitente.value) == 0) {
		alert("CD selecionado é inválido!!");
		return;
	}
	
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
	<td align="right" valign="bottom"><span class="PEDIDO">Solicitação de Coletas</span>
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
<span class="PLTe">TIPO DE RELATÓRIO</span>
<br />
	<table cellSpacing="0" style="margin-left:8px;margin-right:8px;">
	<tr>
		<td align="left"><input type="radio" tabindex="-1" id="rb_tipo_relatorio" name="rb_tipo_relatorio"
				value="<%=COD_TIPO_RELATORIO_SOLICITACAO_COLETA%>" /><span class="C" style="cursor:default" onclick="fFILTRO.rb_tipo_relatorio[0].click();">Solicitação de Coleta</span>
		</td>
	</tr>
	<tr>
		<td align="left"><input type="radio" tabindex="-1" id="rb_tipo_relatorio" name="rb_tipo_relatorio"
				value="<%=COD_TIPO_RELATORIO_PRONTO_PARA_ROMANEIO%>" /><span class="C" style="cursor:default" onclick="fFILTRO.rb_tipo_relatorio[1].click();">Pedidos Prontos para Romaneio</span>
		</td>
	</tr>
	</table>
</td>
</tr>
<tr>
<td class="MDBE" align="left">
<span class="PLTe">LOJAS</span>
<br />
	<table cellSpacing="0" cellPadding="5" style="margin-left:4px;margin-right:8px;">
	<tr bgcolor="#FFFFFF"><td align="left">
		<input type="radio" tabindex="-1" id="rb_loja" name="rb_loja"
			value="TODAS" checked><span class="C" style="cursor:default;" 
			onclick="fFILTRO.rb_loja[0].click();">Todas as lojas</span>
		</td></tr>
	<tr bgcolor="#FFFFFF"><td align="left">
		<input type="radio" tabindex="-1" id="rb_loja" name="rb_loja"
			value="UMA"><span class="C" style="cursor:default" 
			onclick="fFILTRO.rb_loja[1].click();">Loja</span>
			<input class="Cc" maxlength="3" style="width:40px;" name="c_loja" id="c_loja" onblur="this.value=normaliza_codigo(this.value, TAM_MIN_LOJA);" onkeypress="if (digitou_enter(true) && tem_info(this.value)) bCONFIRMA.focus(); else fFILTRO.rb_loja[1].click(); filtra_numerico();" onclick="fFILTRO.rb_loja[1].click();">
		</td></tr>
	<tr bgcolor="#FFFFFF"><td align="left">
		<input type="radio" tabindex="-1" id="rb_loja" name="rb_loja"
			value="FAIXA"><span class="C" style="cursor:default" 
			onclick="fFILTRO.rb_loja[2].click();">Lojas</span>
			<input class="Cc" maxlength="3" style="width:40px;" name="c_loja_de" id="c_loja_de" onblur="this.value=normaliza_codigo(this.value, TAM_MIN_LOJA);" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fFILTRO.c_loja_ate.focus(); else fFILTRO.rb_loja[2].click(); filtra_numerico();" onclick="fFILTRO.rb_loja[2].click();">
			<span class="C">a</span>
			<input class="Cc" maxlength="3" style="width:40px;" name="c_loja_ate" id="c_loja_ate" onblur="this.value=normaliza_codigo(this.value, TAM_MIN_LOJA);" onkeypress="fFILTRO.rb_loja[2].click(); if (digitou_enter(true) && tem_info(this.value)) bCONFIRMA.focus(); filtra_numerico();" onclick="fFILTRO.rb_loja[2].click();">
		</td></tr>
	</table>
</td></tr>

<!--  TRANSPORTADORA  -->
<tr>
	<td class="ME MD PLTe" nowrap align="left" valign="bottom">&nbsp;TRANSPORTADORA</td>
</tr>
<tr bgcolor="#FFFFFF" nowrap>
	<td class="ME MB MD" align="left">
		<table cellspacing="0" cellpadding="0" style="margin:1px 10px 6px 10px;">
		<tr>
			<td align="left" valign="middle">
				<select id="c_filtro_transportadora" name="c_filtro_transportadora" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
				<% =transportadora_monta_itens_select(Null) %>
				</select>
			</td>
			<td style="width:10px;"></td>
			<td align="left" valign="middle">
				<a name="bLimparTransp" id="bLimparTransp" href="javascript:limpaCampoTransp(fFILTRO)" title="limpa o filtro 'Transportadora'">
							<img src="../botao/botao_x_red.gif" style="vertical-align:bottom;margin-bottom:1px;" width="20" height="20" border="0"></a>
			</td>
		</tr>
		</table>
	</td>
</tr>
<!--  DATA DE COLETA  -->
<tr class="TR_DT_ENTREGA">
	<td class="ME MD PLTe" nowrap align="left" valign="bottom">&nbsp;DATA COLETA</td>
</tr>
<tr class="TR_DT_ENTREGA" bgcolor="#FFFFFF" nowrap>
	<td class="ME MB MD" style="padding-left:10px;" align="left">
		<input class="Cc" maxlength="10" style="width:70px;margin-bottom:8px;" name="c_filtro_dt_entrega" id="c_filtro_dt_entrega" onblur="if (!isDate(this)) {alert('Data inválida!'); this.focus();}" onkeypress="if (digitou_enter(true)) bCONFIRMA.focus(); filtra_data();" />
	</td>
</tr>
<% if qtde_nfe_emitente > 1 then %>
<tr>
	<td class="MB ME MD" align="left">
	<table class="Qx" cellspacing="0" cellpadding="0">
	<tr bgcolor="#FFFFFF">
		<td align="left" nowrap>
			<span class="PLTe">CD</span>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="left">
			<table style="margin: 4px 8px 4px 8px;" cellspacing="0" cellpadding="0">
				<tr bgcolor="#FFFFFF">
				<td align="left">
					<select id="c_nfe_emitente" name="c_nfe_emitente" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) {this.options[0].selected=true;}" style="margin-left:5px;margin-top:4pt; margin-bottom:4pt;">
						<%=wms_usuario_x_nfe_emitente_monta_itens_select(usuario, "")%>
					</select>
				</td>
				</tr>
			</table>
		</td>
	</tr>
	</table>
	</td>
</tr>
<% end if %>
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
