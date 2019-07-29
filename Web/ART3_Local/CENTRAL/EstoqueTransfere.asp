<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/global.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =============================================
'	  E S T O Q U E T R A N S F E R E . A S P
'     =============================================
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

	dim usuario
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	CONECTA COM O BANCO DE DADOS
	dim cn
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	dim i
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	if (Not operacao_permitida(OP_CEN_TRANSF_MOV_ESTOQUE_PERFIL_BASICO, s_lista_operacoes_permitidas)) And _
	   (Not operacao_permitida(OP_CEN_TRANSF_MOV_ESTOQUE_PERFIL_AVANCADO, s_lista_operacoes_permitidas)) then 
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	dim intIdx
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
<script src="<%=URL_FILE__JQUERY_MY_GLOBAL%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">

$(function () {
	// Trata o problema em que os campos do formulário são limpos após retornar à esta página c/ o history.back() pela 2ª vez quando ocorre erro de consistência
	if (trim(fOP.c_FormFieldValues.value) != "") {
		stringToForm(fOP.c_FormFieldValues.value, $('#fOP'));
	}
});

function AtivaCampoLoja() {
var f;
	f=fOP;
	f.c_loja.readOnly=false;
	f.c_loja.tabIndex=0;
}

function DesativaCampoLoja() {
var f;
	f=fOP;
	f.c_loja.value=''; 
	f.c_loja.readOnly=true;
	f.c_loja.tabIndex=-1;
}

function AtivaCampoPedido() {
var f;
	f=fOP;
	f.c_pedido.readOnly=false;
	f.c_pedido.tabIndex=0;
}

function DesativaCampoPedido() {
var f;
	f=fOP;
	f.c_pedido.value=''; 
	f.c_pedido.readOnly=true;
	f.c_pedido.tabIndex=-1;
}

function AtivaCampoNumOS() {
var f;
	f=fOP;
	f.c_num_OS.readOnly=false;
	f.c_num_OS.tabIndex=0;
}

function DesativaCampoNumOS() {
var f;
	f=fOP;
	f.c_num_OS.value=''; 
	f.c_num_OS.readOnly=true;
	f.c_num_OS.tabIndex=-1;
}

function ConfiguraListaProdutosEmissaoOS() {
var f,i;
	f=fOP;
	f.c_fabricante[0].readOnly=false;
	f.c_codigo[0].readOnly=false;
	f.c_qtde[0].readOnly=true;
	f.c_qtde[0].value='1';

	for (i=1; i < f.c_codigo.length; i++) {
		f.c_fabricante[i].value='';
		f.c_fabricante[i].readOnly=true;
		f.c_codigo[i].value='';
		f.c_codigo[i].readOnly=true;
		f.c_qtde[i].value='';
		f.c_qtde[i].readOnly=true;
		}
}

function ConfiguraListaProdutosTodasLinhasAtivas() {
var f,i;
	f=fOP;
	for (i=0; i < f.c_codigo.length; i++) {
		f.c_fabricante[i].readOnly = false;
		f.c_codigo[i].readOnly = false;
		f.c_qtde[i].readOnly = false;
		}
}

function reconfigura_loja( f ) {
var i;
	for (i=0; i<f.rb_tipo.length; i++) {
		if (f.rb_tipo[i].checked) {
			f.rb_tipo[i].click();
			break;
			}
		}
}

function fOPConfirma( f ) {
	var i, b, ha_item, s, strValueSelecionado, intIndexSelecionado, blnGerarOS, blnEncerrarOS, blnEstoqueDevolucao;

	if (trim(f.c_id_nfe_emitente.value) == "") {
		alert("Selecione a empresa!!");
		f.c_id_nfe_emitente.focus();
		return;
	}

	b=false;
	for (i=0; i<f.rb_tipo.length; i++) {
		if (f.rb_tipo[i].checked) {
			b=true;
			strValueSelecionado=f.rb_tipo[i].value;
			intIndexSelecionado=i;
			s=right(f.rb_tipo[i].value,3);
			if ((s==ID_ESTOQUE_SHOW_ROOM)||(s==ID_ESTOQUE_DEVOLUCAO)) {
				if (trim(f.c_loja.value)=="") {
					alert("Especifique o número da loja!!");
					f.c_loja.focus();
					return;
					}
				}
			f.op_selecionada_descricao.value=srb_tipo[i].outerText;
			break;
			}
		}

	if (!b) {
		alert("Indique o tipo de transferência entre estoques a ser efetuada!!");
		return;
		}

//	É uma transferência entre estoques?
	if (strValueSelecionado.indexOf("TRANSF_") > -1) {
		if ((strValueSelecionado.indexOf("_" + ID_ESTOQUE_SHOW_ROOM) > -1) || (strValueSelecionado.indexOf("_" + ID_ESTOQUE_DEVOLUCAO) > -1)) {
			if (trim(f.c_loja.value)=="") {
				alert("Especifique o número da loja!!");
				f.c_loja.focus();
				return;
				}
			}
		}
		
//  Gera nova ordem de serviço
//  ==========================
	blnGerarOS=false;
//  "Estoque de Venda para Estoque de Produtos Danificados": gera nova ordem de serviço
	if (f.rb_tipo[intIndexSelecionado].value=="SAI_DAN") {
		blnGerarOS=true;
		if (trim(f.c_pedido.value)=="") {
			alert("Informe o número do pedido!!");
			f.c_pedido.focus();
			return;
			}
		}

//  Transferência de "Show-Room >> Danificados": gera nova ordem de serviço
	if (f.rb_tipo[intIndexSelecionado].value=="TRANSF_SHR_DAN") {
		blnGerarOS=true;
		if (trim(f.c_loja.value)=="") {
			alert("Especifique o número da loja!!");
			f.c_loja.focus();
			return;
			}
		}

//  Transferência de "Devolução >> Danificados": gera nova ordem de serviço
	if (f.rb_tipo[intIndexSelecionado].value=="TRANSF_DEV_DAN") {
		blnGerarOS=true;
		if (trim(f.c_loja.value)=="") {
			alert("Especifique o número da loja!!");
			f.c_loja.focus();
			return;
			}
		if (trim(f.c_pedido.value)=="") {
			alert("Informe o número do pedido!!");
			f.c_pedido.focus();
			return;
			}
		}

//  Encerra uma ordem de serviço
//  ============================
	blnEncerrarOS=false;
//  "Estoque de Produtos Danificados para Estoque de Venda": encerra uma ordem de serviço
	if (f.rb_tipo[intIndexSelecionado].value=="ENT_DAN") {
		blnEncerrarOS=true;
		if (trim(f.c_num_OS.value)=="") {
			alert("Informe o número da Ordem de Serviço!!");
			f.c_num_OS.focus();
			return;
			}
		}
		
//  Transferência de "Danificados >> Show-Room": encerra uma ordem de serviço
	if (f.rb_tipo[intIndexSelecionado].value=="TRANSF_DAN_SHR") {
		blnEncerrarOS=true;
		if (trim(f.c_loja.value)=="") {
			alert("Especifique o número da loja!!");
			f.c_loja.focus();
			return;
			}
		if (trim(f.c_num_OS.value)=="") {
			alert("Informe o número da Ordem de Serviço!!");
			f.c_num_OS.focus();
			return;
			}
		}

//  Transferência de "Danificados >> Roubo ou Perda Total": encerra uma ordem de serviço
	if (f.rb_tipo[intIndexSelecionado].value=="TRANSF_DAN_ROU") {
		blnEncerrarOS=true;
		if (trim(f.c_num_OS.value)=="") {
			alert("Informe o número da Ordem de Serviço!!");
			f.c_num_OS.focus();
			return;
			}
		}

//  Estoque de Devolução
//  ====================
	blnEstoqueDevolucao=false;
	if (f.rb_tipo[intIndexSelecionado].value=="ENT_DEV") blnEstoqueDevolucao=true;
	if (f.rb_tipo[intIndexSelecionado].value=="TRANSF_DEV_DAN") blnEstoqueDevolucao=true;
	if (f.rb_tipo[intIndexSelecionado].value=="TRANSF_DEV_ROU") blnEstoqueDevolucao=true;
	if (blnEstoqueDevolucao) {
		if (trim(f.c_pedido.value)=="") {
			alert("Informe o número do pedido!!");
			f.c_pedido.focus();
			return;
			}
		}
		
//  Verifica a lista de produtos
//  ============================
	ha_item=false;
	for (i=0; i < f.c_codigo.length; i++) {
		b=false;
		if (trim(f.c_fabricante[i].value)!="") b=true;
		if (trim(f.c_codigo[i].value)!="") b=true;
		if (trim(f.c_qtde[i].value)!="") b=true;
		
		if (b) {
			ha_item=true;
			if (!isEAN(trim(f.c_codigo[i].value))) {
				if (trim(f.c_fabricante[i].value)=="") {
					alert("Informe o fabricante do produto a ser transferido!!");
					f.c_fabricante[i].focus();
					return;
					}
				}
			if (trim(f.c_codigo[i].value)=="") {
				alert("Informe o código do produto a ser transferido!!");
				f.c_codigo[i].focus();
				return;
				}
			if (trim(f.c_qtde[i].value)=="") {
				alert("Informe a quantidade de produtos a ser transferida!!");
				f.c_qtde[i].focus();
				return;
				}
			if (parseInt(f.c_qtde[i].value)<=0) {
				alert("Quantidade inválida!!");
				f.c_qtde[i].focus();
				return;
				}
			}
		}

	if (!ha_item) {
		alert("Não há produtos para transferir!!");
		f.c_fabricante[0].focus();
		return;
		}

	fOP.c_FormFieldValues.value = formToString($("#fOP"));

	f.ckb_spe_descricao.value=sckb_spe.outerText;
	dCONFIRMA.style.visibility="hidden";
	window.status = "Aguarde ...";
		
	if (blnGerarOS) f.action='OrdemServicoNova.asp';
	if (blnEncerrarOS) f.action='OrdemServicoEncerra.asp';
	
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

<style type="text/css">
#rb_tipo {
	margin: 0pt 2pt 0pt 10pt;
	}
#srb_tipo {
	margin: 0pt 10pt 0pt 0pt;
	}
#ckb_spe {
	margin: 0pt 2pt 0pt 10pt;
	}
#sckb_spe {
	margin: 0pt 10pt 0pt 0pt;
	}
</style>


<body onload="reconfigura_loja(fOP); focus();">
<center>

<form id="fOP" name="fOP" method="post" action="EstoqueTransfereConsiste.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type=hidden name="op_selecionada_descricao" id="op_selecionada_descricao" value="">
<input type=hidden name="ckb_spe_descricao" id="ckb_spe_descricao" value="">
<input type="hidden" name="c_FormFieldValues" id="c_FormFieldValues" value="" />
<input type="hidden" name="url_back" id="url_back" value="EstoqueTransfere.asp" />

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Transferência/Movimentação do Estoque</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>
<br>

<!--  TIPO DE TRANSFERÊNCIA  -->
<table class="Qx" cellspacing="0">
	<!--  EMPRESA  -->
	<tr bgcolor="#FFFFFF">
		<td><span style="width:30px;">&nbsp;</span></td>
		<td colspan="3" class="MT" valign="middle" align="center" nowrap style="background:azure;"><span class="PLTc" style="vertical-align:middle;">EMPRESA</span></td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td>&nbsp;</td>
		<td class="MDBE" align="left" nowrap>
			<select id="c_id_nfe_emitente" name="c_id_nfe_emitente" style="margin:6px 10px 6px 5px;min-width:100px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
			<%=wms_apelido_empresa_nfe_emitente_monta_itens_select(Null) %>
			</select>
		</td>
	</tr>

	<!--  PULA LINHA  -->
	<tr bgcolor="#FFFFFF">
		<td><span style="width:30px;">&nbsp;</span></td>
		<td colspan="3">&nbsp;</td>
	</tr>

	<!--  TÍTULO  -->
	<tr bgcolor="#FFFFFF">
	<td><span style="width:30px;">&nbsp;</span></td>
	<td colspan="3" class="MT" valign="middle" align="center" nowrap style="background:azure;"><span class="PLTc" style="vertical-align:middle;"
		>TRANSFERÊNCIA/MOVIMENTAÇÃO DO ESTOQUE</span></td>
	</tr>
	<!--  OPÇÕES  -->
	<% intIdx = -1 %>
	<tr bgcolor="#FFFFFF">
	<td>&nbsp;</td>
	<td class="MDBE" align="left" nowrap>
		<span class="PLTe">Saída</span>
	<% if operacao_permitida(OP_CEN_TRANSF_MOV_ESTOQUE_PERFIL_AVANCADO, s_lista_operacoes_permitidas) then %>
		<br>
		<% intIdx = intIdx+1 %>
		<input type="radio" tabindex="-1" name="rb_tipo" id="rb_tipo_<%=Cstr(intIdx)%>" value="SAI_<%=ID_ESTOQUE_SHOW_ROOM%>" 
			onclick="AtivaCampoLoja();DesativaCampoPedido();DesativaCampoNumOS();ConfiguraListaProdutosTodasLinhasAtivas();fOP.c_loja.focus();"
		><span class="C" id="srb_tipo" name="srb_tipo" style="cursor:default" onclick="fOP.rb_tipo[<%=Cstr(intIdx)%>].click();"
			>Estoque de Venda para Estoque de Show-Room</span>
	<% end if %>
		<br>
		<% intIdx = intIdx+1 %>
		<input type="radio" tabindex="-1" name="rb_tipo" id="rb_tipo_<%=Cstr(intIdx)%>" value="SAI_<%=ID_ESTOQUE_DANIFICADOS%>" 
			onclick="DesativaCampoLoja();DesativaCampoNumOS();ConfiguraListaProdutosEmissaoOS();AtivaCampoPedido();fOP.c_pedido.focus();"
		><span class="C" id="srb_tipo" name="srb_tipo" style="cursor:default" onclick="fOP.rb_tipo[<%=Cstr(intIdx)%>].click();"
			>Estoque de Venda para Estoque de Produtos Danificados</span>
	<% if operacao_permitida(OP_CEN_TRANSF_MOV_ESTOQUE_PERFIL_AVANCADO, s_lista_operacoes_permitidas) then %>
		<br>
		<% intIdx = intIdx+1 %>
		<input type="radio" tabindex="-1" name="rb_tipo" id="rb_tipo_<%=Cstr(intIdx)%>" value="SAI_<%=ID_ESTOQUE_ROUBO%>" 
			onclick="DesativaCampoLoja();DesativaCampoPedido();DesativaCampoNumOS();ConfiguraListaProdutosTodasLinhasAtivas();fOP.c_fabricante[0].focus();"
		><span class="C" id="srb_tipo" name="srb_tipo" style="cursor:default" onclick="fOP.rb_tipo[<%=Cstr(intIdx)%>].click();"
			>Saída do Estoque de Venda devido a Roubo ou Perda Total</span>
	<% end if %>
		</td>
	<tr bgcolor="#FFFFFF">
	<td>&nbsp;</td>
	<td class="MDBE" align="left" nowrap>
		<span class="PLTe">Entrada</span>
		<br>
		<% intIdx = intIdx+1 %>
		<input type="radio" tabindex="-1" name="rb_tipo" id="rb_tipo_<%=Cstr(intIdx)%>" value="ENT_<%=ID_ESTOQUE_DEVOLUCAO%>" 
			onclick="AtivaCampoLoja();AtivaCampoPedido();DesativaCampoNumOS();ConfiguraListaProdutosTodasLinhasAtivas();fOP.c_loja.focus();"
		><span class="C" id="srb_tipo" name="srb_tipo" style="cursor:default" onclick="fOP.rb_tipo[<%=Cstr(intIdx)%>].click();"
			>Estoque de Devolução para Estoque de Venda</span>
	<% if operacao_permitida(OP_CEN_TRANSF_MOV_ESTOQUE_PERFIL_AVANCADO, s_lista_operacoes_permitidas) then %>
		<br>
		<% intIdx = intIdx+1 %>
		<input type="radio" tabindex="-1" name="rb_tipo" id="rb_tipo_<%=Cstr(intIdx)%>" value="ENT_<%=ID_ESTOQUE_SHOW_ROOM%>" 
			onclick="AtivaCampoLoja();DesativaCampoPedido();DesativaCampoNumOS();ConfiguraListaProdutosTodasLinhasAtivas();fOP.c_loja.focus();"
		><span class="C" id="srb_tipo" name="srb_tipo" style="cursor:default" onclick="fOP.rb_tipo[<%=Cstr(intIdx)%>].click();"
			>Estoque de Show-Room para Estoque de Venda</span>
	<% end if %>
		<br>
		<% intIdx = intIdx+1 %>
		<input type="radio" tabindex="-1" name="rb_tipo" id="rb_tipo_<%=Cstr(intIdx)%>" value="ENT_<%=ID_ESTOQUE_DANIFICADOS%>" 
			onclick="DesativaCampoLoja();DesativaCampoPedido();AtivaCampoNumOS();ConfiguraListaProdutosEmissaoOS();fOP.c_num_OS.focus();"
		><span class="C" id="srb_tipo" name="srb_tipo" style="cursor:default" onclick="fOP.rb_tipo[<%=Cstr(intIdx)%>].click();"
			>Estoque de Produtos Danificados para Estoque de Venda</span>
	<% if operacao_permitida(OP_CEN_TRANSF_MOV_ESTOQUE_PERFIL_AVANCADO, s_lista_operacoes_permitidas) then %>
		<br>
		<% intIdx = intIdx+1 %>
		<input type="radio" tabindex="-1" name="rb_tipo" id="rb_tipo_<%=Cstr(intIdx)%>" value="ENT_<%=ID_ESTOQUE_ROUBO%>" 
			onclick="DesativaCampoLoja();DesativaCampoPedido();DesativaCampoNumOS();ConfiguraListaProdutosTodasLinhasAtivas();fOP.c_fabricante[0].focus();"
		><span class="C" id="srb_tipo" name="srb_tipo" style="cursor:default" onclick="fOP.rb_tipo[<%=Cstr(intIdx)%>].click();"
			>Estorno de Roubo ou Perda Total</span>
	<% end if %>
		<br><input type="checkbox" tabindex="-1" id="ckb_spe" name="ckb_spe" value="SPE_ON"
		><span class="C" id="sckb_spe" name="sckb_spe" style="cursor:default;font-weight:normal;font-style:normal;" 
			onclick="fOP.ckb_spe.click();"
			>Não atender aos pedidos que aguardam a chegada de produtos</span>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
	<td>&nbsp;</td>
	<td class="MDBE" align="left" nowrap>
		<span class="PLTe">Transferência</span>
	<% if operacao_permitida(OP_CEN_TRANSF_MOV_ESTOQUE_PERFIL_AVANCADO, s_lista_operacoes_permitidas) then %>
		<br>
		<% intIdx = intIdx+1 %>
		<input type="radio" tabindex="-1" name="rb_tipo" id="rb_tipo_<%=Cstr(intIdx)%>" value="TRANSF_<%=ID_ESTOQUE_DANIFICADOS & "_" & ID_ESTOQUE_SHOW_ROOM%>" 
			onclick="AtivaCampoLoja();AtivaCampoNumOS();DesativaCampoPedido();ConfiguraListaProdutosEmissaoOS();fOP.c_loja.focus();"
		><span class="C" id="srb_tipo" name="srb_tipo" style="cursor:default" onclick="fOP.rb_tipo[<%=Cstr(intIdx)%>].click();"
			>Danificados &nbsp;&nbsp; >> &nbsp;&nbsp; Show-Room</span>
		<br>
		<% intIdx = intIdx+1 %>
		<input type="radio" tabindex="-1" name="rb_tipo" id="rb_tipo_<%=Cstr(intIdx)%>" value="TRANSF_<%=ID_ESTOQUE_SHOW_ROOM & "_" & ID_ESTOQUE_DANIFICADOS%>" 
			onclick="AtivaCampoLoja();DesativaCampoPedido();DesativaCampoNumOS();ConfiguraListaProdutosEmissaoOS();fOP.c_loja.focus();"
		><span class="C" id="srb_tipo" name="srb_tipo" style="cursor:default" onclick="fOP.rb_tipo[<%=Cstr(intIdx)%>].click();"
			>Show-Room &nbsp;&nbsp; >> &nbsp;&nbsp; Danificados</span>
	<% end if %>
		<br>
		<% intIdx = intIdx+1 %>
		<input type="radio" tabindex="-1" name="rb_tipo" id="rb_tipo_<%=Cstr(intIdx)%>" value="TRANSF_<%=ID_ESTOQUE_DEVOLUCAO & "_" & ID_ESTOQUE_DANIFICADOS%>" 
			onclick="AtivaCampoLoja();AtivaCampoPedido();DesativaCampoNumOS();ConfiguraListaProdutosEmissaoOS();fOP.c_loja.focus();"
		><span class="C" id="srb_tipo" name="srb_tipo" style="cursor:default" onclick="fOP.rb_tipo[<%=Cstr(intIdx)%>].click();"
			>Devolução &nbsp;&nbsp; >> &nbsp;&nbsp; Danificados</span>
	<% if operacao_permitida(OP_CEN_TRANSF_MOV_ESTOQUE_PERFIL_AVANCADO, s_lista_operacoes_permitidas) then %>
		<br>
		<% intIdx = intIdx+1 %>
		<input type="radio" tabindex="-1" name="rb_tipo" id="rb_tipo_<%=Cstr(intIdx)%>" value="TRANSF_<%=ID_ESTOQUE_DANIFICADOS & "_" & ID_ESTOQUE_ROUBO%>" 
			onclick="AtivaCampoNumOS();DesativaCampoLoja();DesativaCampoPedido();ConfiguraListaProdutosEmissaoOS();fOP.c_num_OS.focus();"
		><span class="C" id="srb_tipo" name="srb_tipo" style="cursor:default" onclick="fOP.rb_tipo[<%=Cstr(intIdx)%>].click();"
			>Danificados &nbsp;&nbsp; >> &nbsp;&nbsp; Roubo ou Perda Total</span>
		<br>
		<% intIdx = intIdx+1 %>
		<input type="radio" tabindex="-1" name="rb_tipo" id="rb_tipo_<%=Cstr(intIdx)%>" value="TRANSF_<%=ID_ESTOQUE_SHOW_ROOM & "_" & ID_ESTOQUE_ROUBO%>" 
			onclick="AtivaCampoLoja();DesativaCampoPedido();DesativaCampoNumOS();ConfiguraListaProdutosTodasLinhasAtivas();fOP.c_loja.focus();"
		><span class="C" id="srb_tipo" name="srb_tipo" style="cursor:default" onclick="fOP.rb_tipo[<%=Cstr(intIdx)%>].click();"
			>Show-Room &nbsp;&nbsp; >> &nbsp;&nbsp; Roubo ou Perda Total</span>
		<br>
		<% intIdx = intIdx+1 %>
		<input type="radio" tabindex="-1" name="rb_tipo" id="rb_tipo_<%=Cstr(intIdx)%>" value="TRANSF_<%=ID_ESTOQUE_DEVOLUCAO & "_" & ID_ESTOQUE_ROUBO%>" 
			onclick="AtivaCampoLoja();AtivaCampoPedido();DesativaCampoNumOS();ConfiguraListaProdutosTodasLinhasAtivas();fOP.c_loja.focus();"
		><span class="C" id="srb_tipo" name="srb_tipo" style="cursor:default" onclick="fOP.rb_tipo[<%=Cstr(intIdx)%>].click();"
			>Devolução &nbsp;&nbsp; >> &nbsp;&nbsp; Roubo ou Perda Total</span>
	<% end if %>
		</td>
	
	</tr>
	<tr bgcolor="#FFFFFF">
	<td>&nbsp;</td>
	<td class="MDBE" align="left" nowrap><span class="PLTe">Loja</span>
		<br><input name="c_loja" id="c_loja" readonly class="PLLe" maxlength="3" style="width:60px;margin-left:4px;"
				onkeypress="if (digitou_enter(true)&&tem_info(this.value)) $(this).hUtil('focusNext'); filtra_numerico();"
				onblur="this.value=normaliza_codigo(this.value,TAM_MIN_LOJA);">
	</td>
	</tr>
	<tr bgcolor="#FFFFFF">
	<td>&nbsp;</td>
	<td class="MDBE" align="left" nowrap><span class="PLTe">Pedido</span>
		<br><input name="c_pedido" id="c_pedido" readonly class="PLLe" maxlength="9" style="width:90px;margin-left:4px;"
				onkeypress="if (digitou_enter(true)&&tem_info(this.value)) $(this).hUtil('focusNext'); filtra_pedido();"
				onblur="if (normaliza_num_pedido(this.value)!='') this.value=normaliza_num_pedido(this.value);">
	</td>
	</tr>
	<tr bgcolor="#FFFFFF">
	<td>&nbsp;</td>
	<td class="MDBE" align="left" nowrap><span class="PLTe">Nº Ordem de Serviço</span>
		<br><input name="c_num_OS" id="c_num_OS" readonly class="PLLe" maxlength="12" style="width:90px;margin-left:4px;"
				onkeypress="if (digitou_enter(true)&&tem_info(this.value)) $(this).hUtil('focusNext'); filtra_numerico();">
	</td>
	</tr>
</table>
<br><br>

<!--  R E L A Ç Ã O   D E   P R O D U T O S  -->
<table class="Qx" cellspacing="0">
	<!--  TÍTULO DA TABELA  -->
	<tr bgcolor="#FFFFFF">
	<td>&nbsp;</td>
	<td colspan="3" class="MT" valign="middle" align="center" nowrap style="background:azure;"><span class="PLTc" style="vertical-align:middle;"
		>PRODUTOS A MOVIMENTAR</span></td>
	</tr>
	<!--  TÍTULO DAS COLUNAS  -->
	<tr bgcolor="#FFFFFF">
	<td>&nbsp;</td>
	<td class="MDBE" align="left"><span class="PLTe">Fabricante&nbsp;</span></td>
	<td class="MDB" align="left"><span class="PLTe">Produto</span></td>
	<td class="MDB" align="right"><span class="PLTd">Qtde</span></td>
	</tr>
<% for i=1 to MAX_PRODUTOS_TRANSFERENCIA_ESTOQUE %>
	<tr>
	<td align="left">
		<input name="c_linha" id="c_linha_<%=Cstr(i)%>" readonly tabindex=-1 class="PLLe" maxlength="2" style="width:30px;text-align:right;color:#808080;" 
			value="<%=Cstr(i) & ". " %>"></td>
	<td class="MDBE" align="left">
		<input name="c_fabricante" id="c_fabricante_<%=Cstr(i)%>" readonly class="PLLe" maxlength="4" style="width:60px;" 
			onkeypress="if (digitou_enter(true)) fOP.c_codigo[<%=Cstr(i-1)%>].focus(); filtra_fabricante();" 
			onblur="this.value=normaliza_codigo(this.value,TAM_MIN_FABRICANTE);">
	</td>
	<td class="MDB" align="left">
		<input name="c_codigo" id="c_codigo_<%=Cstr(i)%>" readonly class="PLLe" maxlength="13" style="width:100px;" 
			onkeypress="if (digitou_enter(true)&&(tem_info(this.value)||(<%=Cstr(i)%>!=1))) if (trim(this.value)=='') bCONFIRMA.focus(); else {if (!fOP.c_qtde[<%=Cstr(i-1)%>].readOnly) fOP.c_qtde[<%=Cstr(i-1)%>].focus(); else bCONFIRMA.focus();} filtra_produto();" 
			onblur="this.value=normaliza_produto(this.value);">
	</td>
	<td class="MDB" align="right">
		<input name="c_qtde" id="c_qtde_<%=Cstr(i)%>" readonly class="PLLd" maxlength="4" style="width:35px;" 
			onkeypress="if (digitou_enter(true)) {if (<%=Cstr(i)%>==fOP.c_qtde.length) bCONFIRMA.focus(); else fOP.c_fabricante[<%=Cstr(i)%>].focus();} filtra_numerico();">
	</td>
	</tr>
<% next %>
</table>


<!-- ************   SEPARADOR   ************ -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>
<br>


<table width="649" cellspacing="0">
<tr>
	<td align="left"><a name="bCANCELA" id="bCANCELA" href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="cancela a operação">
		<img src="../botao/cancelar.gif" width="176" height="55" border="0"></a></td>
	<td align="right"><div name="dCONFIRMA" id="dCONFIRMA"><a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fOPConfirma(fOP)" title="vai para a página de confirmação">
		<img src="../botao/proximo.gif" width="176" height="55" border="0"></a></div>
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
