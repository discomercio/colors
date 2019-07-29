<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp"        -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%

'     =========================================
'	  M E N U C A D A S T R O . A S P
'     =========================================
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



' _____________________________________________________________________________________________
'
'						I N I C I A L I Z A     P Á G I N A     A S P
' _____________________________________________________________________________________________

	On Error GoTo 0
	Err.Clear

'	OBTEM USUÁRIO
	dim s, usuario, usuario_nome
	usuario = trim(Session("usuario_atual"))
	usuario_nome = Trim(Session("usuario_nome_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

	dim idx
	
'	CONECTA COM O BANCO DE DADOS
	dim cn
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

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
	<title>CENTRAL ADMINISTRATIVA</title>
	</head>

<script language="JavaScript" type="text/javascript">
window.focus();
</script>

<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
function fOPConcluir( f ){
var s, strCep, iop;

	iop=-1;
	s="";

 // CADASTRO DE PERFIS
	iop++;
	if (f.rb_op[iop].checked) {
		s="Perfil.asp"; 
		}
 
 // CADASTRO DE USUÁRIOS
	iop++;
	if (f.rb_op[iop].checked) {
		s="usuario.asp"; 
		}

 // CADASTRO DE ORÇAMENTISTAS/INDICADORES
	iop++;
	if (f.rb_op[iop].checked) {
		s="MenuOrcamentistaEIndicador.asp"; 
		}

 // CADASTRO DE EQUIPE DE VENDAS
	iop++;
	if (f.rb_op[iop].checked) {
		s="MenuEquipeVendas.asp"; 
		}

 // CADASTRO DE LOJAS
	iop++;
	if (f.rb_op[iop].checked) {
		s="loja.asp";
		}

 // CADASTRO DE GRUPO DE LOJAS
	iop++;
	if (f.rb_op[iop].checked) {
		s="grupolojas.asp";
		}

 // CADASTRO DE FABRICANTES
	iop++;
	if (f.rb_op[iop].checked) {
		s="fabricante.asp";
		}
 
 // CADASTRO DE TRANSPORTADORAS
	iop++;
	if (f.rb_op[iop].checked) {
		s="transportadora.asp";
		}
 
 // CADASTRO DE VEÍCULOS DE MÍDIA
	iop++;
	if (f.rb_op[iop].checked) {
		s="midia.asp";
		}
 
 // CADASTRO DE MENSAGENS DE ALERTA PARA PRODUTOS
	iop++;
	if (f.rb_op[iop].checked) {
		s="MensagemAlertaProduto.asp";
		}

 // CADASTRO DE PRODUTO COMPOSTO
	iop++;
	if (f.rb_op[iop].checked) {
		s = "ECProdutoCompostoMenu.asp";
	}
	
 // CADASTRO DO QUADRO DE AVISOS
	iop++;
	if (f.rb_op[iop].checked) {
		s="quadroaviso.asp";
		}

 // CADASTRO DE CEP
	iop++;
	if (f.rb_op[iop].checked) {
		strCep=retorna_so_digitos(trim(f.c_cep.value));
		if (strCep.length==0) {
			alert("Informe o CEP!!");
			f.c_cep.focus();
			return;
			}
		if (strCep.length!=8) {
			alert("CEP com tamanho inválido!!");
			f.c_cep.focus();
			return;
			}
		s="CepEdita.asp";
		}

 // OPÇÕES PARA "FORMA COMO CONHECEU A BONSHOP" (CADASTRO DE INDICADORES)
		iop++;
		if (f.rb_op[iop].checked) {
			s = "CadIndicadorOpcoesFormaComoConheceu.asp";
		}

 // OPÇÕES DE PAGAMENTO PARA CARTÃO DE CRÉDITO
	iop++;
	if (f.rb_op[iop].checked) {
		s="VisanetOpcoesPagto.asp";
		}

 // CADASTRA O VALOR P/ APROVAÇÃO AUTOMÁTICA DA ANÁLISE DE CRÉDITO
	iop++;
	if (f.rb_op[iop].checked) {
		s="AnaliseCreditoCadAprovAuto.asp";
		}

 // CADASTRA O PERCENTUAL LIMITE DE RA ANTES DE APLICAR DESÁGIO
	iop++;
	if (f.rb_op[iop].checked) {
		s="PercLimiteRASemDesagio.asp";
		}

 // CADASTRA O PERCENTUAL MÁXIMO DE COMISSÃO (ANTERIORMENTE CHAMADO DE RT)
	iop++;
	if (f.rb_op[iop].checked) {
		s="EditaParamGlobalPerc.asp";
		f.TituloParametroGlobal.value="Percentual Máximo de Comissão";
		f.IdParametro.value=f.ID_PARAM_PERC_MAX_RT.value;
		}

 // CADASTRA O PERCENTUAL MÁXIMO DE DESCONTO SEM ZERAR A COMISSÃO (ANTERIORMENTE CHAMADO DE RT)
	iop++;
	if (f.rb_op[iop].checked) {
		s="CadPercMaxDescSemZerarRT.asp";
		}

 // CADASTRA PERCENTUAL MÁXIMO DE COMISSÃO E DESCONTO POR LOJA
		iop++;
		if (f.rb_op[iop].checked) {
			s = "CadPercMaxComissaoEDescPorLoja.asp";
		}

 // CADASTRA O PERCENTUAL MÁXIMO DA SENHA DE DESCONTO SUPERIOR P/ CADASTRAMENTO NA LOJA
	iop++;
	if (f.rb_op[iop].checked) {
		s="CadPercMaxDescCadLoja.asp";
		}

 // CADASTRA O PERÍODO DE CONSULTA EM FILTROS LIMITADO POR PERFIL DE ACESSO
	iop++;
	if (f.rb_op[iop].checked) {
		s="EditaParamGlobalNumInteiro.asp";
		f.TituloParametroGlobal.value="Período de Consulta em Filtros Limitado por Perfil de Acesso";
		f.TituloUnidadeMedida.value="dias";
		f.IdParametro.value=f.ID_PARAM_MAX_DIAS_DT_INICIAL_FILTRO_PERIODO.value;
		}

 // CADASTRA O PERCENTUAL MÁXIMO DO VALOR DO PEDIDO QUE A RA PODE ATINGIR
	iop++;
	if (f.rb_op[iop].checked) {
		s="EditaParamGlobalPerc.asp";
		f.TituloParametroGlobal.value="Limitar a RA a um Percentual do Valor do Pedido";
		f.IdParametro.value=f.ID_PARAM_PercVlPedidoLimiteRA.value;
		}

 // TABELA DE COMISSÃO DO VENDEDOR
	iop++;
	if (f.rb_op[iop].checked) {
		s="TabelaComissaoVendedorFiltro.asp";
		}

 // TABELA DE CUSTO FINANCEIRO POR FORNECEDOR
	iop++;
	if (f.rb_op[iop].checked) {
		s="TabelaCustoFinanceiroFornecedorFiltro.asp";
		}

// Multi CD: Cadastro de Regras do Consumo do Estoque
	iop++;
	if (f.rb_op[iop].checked) {
		s = "MultiCDRegraMenu.asp";
	}

// Multi CD: Associação do Produto com a Regra
	iop++;
	if (f.rb_op[iop].checked) {
		s = "MultiCDAssocProdRegraMenu.asp";
	}

 // FINANCEIRO: CADASTRO DE CONTA CORRENTE
	iop++;
	if (f.rb_op[iop].checked) {
		s="FinCadContaCorrenteMenu.asp";
		}

 // FINANCEIRO: CADASTRO DE EMPRESA DO PLANO DE CONTAS
	iop++;
	if (f.rb_op[iop].checked) {
		s="FinCadPlanoContasEmpresaMenu.asp";
		}

 // FINANCEIRO: CADASTRO DE GRUPO DO PLANO DE CONTAS
	iop++;
	if (f.rb_op[iop].checked) {
		s="FinCadPlanoContasGrupoMenu.asp";
		}

 // FINANCEIRO: CADASTRO DE CONTA DO PLANO DE CONTAS
	iop++;
	if (f.rb_op[iop].checked) {
		s="FinCadPlanoContasContaMenu.asp";
		}

 // FINANCEIRO: BOLETO - PARÂMETROS DO CEDENTE
	iop++;
	if (f.rb_op[iop].checked) {
		s="FinCadBoletoCedenteParametrosMenu.asp";
		}

 // FINANCEIRO: UNIDADES DE NEGÓCIO
		iop++;
		if (f.rb_op[iop].checked) {
			s = "FinCadUnidadeNegocioMenu.asp";
		}

 // FINANCEIRO: RATEIO ENTRE AS UNIDADES DE NEGÓCIO
		iop++;
		if (f.rb_op[iop].checked) {
			s = "FinCadUnidadeNegocioRateioMenu.asp";
		}

	if (s=="") {
		alert("Escolha uma das funções!!");
		return false;
		}

	window.status = "Aguarde ...";
	f.action=s;
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


<body>
<!--  MENU SUPERIOR -->

<table width="100%" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">

<tr>
	<td align="right" valign="bottom"><p class="PEDIDO">CENTRAL&nbsp;&nbsp;ADMINISTRATIVA<br>
	<%	s = usuario_nome
		if s = "" then s = usuario
		s = x_saudacao & ", " & s
		s = "<span class='Cd'>" & s & "</span><br>"
	%>
	<%=s%>
	<span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="senha.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="altera a senha atual do usuário" class="LAlteraSenha">altera senha</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></p></td>
	</tr>

</table>

<br>
<center>


<!--  ***********************************************************************************************  -->
<!--  C A D A S T R O S                         												       -->
<!--  ***********************************************************************************************  -->
<form method="post" id="fOP" name="fOP" onsubmit="if (!fOPConcluir(fOP)) return false">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>

<!--  CAMPOS P/ UTILIZAR FORM GENÉRICO DE CADASTRAMENTO DE PARÂMETROS GLOBAIS  -->
<input type="hidden" name='TituloParametroGlobal' id="TituloParametroGlobal" value=''>
<input type="hidden" name='IdParametro' id="IdParametro" value=''>
<input type="hidden" name='TituloUnidadeMedida' id="TituloUnidadeMedida" value=''>
<input type="hidden" name='ID_PARAM_PERC_MAX_RT' id="ID_PARAM_PERC_MAX_RT" value='<%=ID_PARAM_PERC_MAX_RT%>'>
<input type="hidden" name='ID_PARAM_MAX_DIAS_DT_INICIAL_FILTRO_PERIODO' id="ID_PARAM_MAX_DIAS_DT_INICIAL_FILTRO_PERIODO" value='<%=ID_PARAM_MAX_DIAS_DT_INICIAL_FILTRO_PERIODO%>'>
<input type="hidden" name='ID_PARAM_PercVlPedidoLimiteRA' id="ID_PARAM_PercVlPedidoLimiteRA" value='<%=ID_PARAM_PercVlPedidoLimiteRA%>'>

<span class="T">CADASTROS</span>
<div class="QFn" align="CENTER" style="width:640px;">
<table class="TFn">
	<tr>
		<td align="left" nowrap>
			<%	idx = 0 %>
			
			<%	idx=idx+1
				if operacao_permitida(OP_CEN_CADASTRO_PERFIL, s_lista_operacoes_permitidas) then s="" else s=" disabled" %>
			<input type="radio" id="rb_op" name="rb_op" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fOP.rb_op[<%=Cstr(idx-1)%>].click(); if (fOP.rb_op[<%=Cstr(idx-1)%>].checked) fOP.bEXECUTAR.click();"
				>Perfil</span><br>

			<%	idx=idx+1
				if operacao_permitida(OP_CEN_CADASTRO_USUARIOS, s_lista_operacoes_permitidas) then s="" else s=" disabled" %>
			<input type="radio" id="rb_op" name="rb_op" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fOP.rb_op[<%=Cstr(idx-1)%>].click(); if (fOP.rb_op[<%=Cstr(idx-1)%>].checked) fOP.bEXECUTAR.click();"
				>Usuários</span><br>
			
			<%	idx=idx+1
				if (operacao_permitida(OP_CEN_CADASTRO_ORCAMENTISTAS_E_INDICADORES, s_lista_operacoes_permitidas) Or operacao_permitida(OP_CEN_GER_LIST_CAD_INDICADORES, s_lista_operacoes_permitidas)) then s="" else s=" disabled" %>
			<input type="radio" id="rb_op" name="rb_op" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fOP.rb_op[<%=Cstr(idx-1)%>].click(); if (fOP.rb_op[<%=Cstr(idx-1)%>].checked) fOP.bEXECUTAR.click();"
				>Orçamentistas / Indicadores</span><br>

			<%	idx=idx+1
				if operacao_permitida(OP_CEN_CADASTRO_EQUIPE_VENDAS, s_lista_operacoes_permitidas) then s="" else s=" disabled" %>
			<input type="radio" id="rb_op" name="rb_op" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fOP.rb_op[<%=Cstr(idx-1)%>].click(); if (fOP.rb_op[<%=Cstr(idx-1)%>].checked) fOP.bEXECUTAR.click();"
				>Equipe de Vendas</span><br>
			
			<%	idx=idx+1
				if operacao_permitida(OP_CEN_CADASTRO_LOJAS, s_lista_operacoes_permitidas) then s="" else s=" disabled" %>
			<input type="radio" id="rb_op" name="rb_op" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fOP.rb_op[<%=Cstr(idx-1)%>].click(); if (fOP.rb_op[<%=Cstr(idx-1)%>].checked) fOP.bEXECUTAR.click();"
				>Lojas</span><br>
			
			<%	idx=idx+1
				if operacao_permitida(OP_CEN_CADASTRO_GRUPO_LOJAS, s_lista_operacoes_permitidas) then s="" else s=" disabled" %>
			<input type="radio" id="rb_op" name="rb_op" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fOP.rb_op[<%=Cstr(idx-1)%>].click(); if (fOP.rb_op[<%=Cstr(idx-1)%>].checked) fOP.bEXECUTAR.click();"
				>Grupo de Lojas</span><br>
			
			<%	idx=idx+1
				if operacao_permitida(OP_CEN_CADASTRO_FABRICANTES, s_lista_operacoes_permitidas) then s="" else s=" disabled" %>
			<input type="radio" id="rb_op" name="rb_op" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fOP.rb_op[<%=Cstr(idx-1)%>].click(); if (fOP.rb_op[<%=Cstr(idx-1)%>].checked) fOP.bEXECUTAR.click();"
				>Fabricantes</span><br>
			
			<%	idx=idx+1
				if operacao_permitida(OP_CEN_CADASTRO_TRANSPORTADORAS, s_lista_operacoes_permitidas) then s="" else s=" disabled" %>
			<input type="radio" id="rb_op" name="rb_op" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fOP.rb_op[<%=Cstr(idx-1)%>].click(); if (fOP.rb_op[<%=Cstr(idx-1)%>].checked) fOP.bEXECUTAR.click();"
				>Transportadoras</span><br>
			
			<%	idx=idx+1
				if operacao_permitida(OP_CEN_CADASTRO_VEICULOS_MIDIA, s_lista_operacoes_permitidas) then s="" else s=" disabled" %>
			<input type="radio" id="rb_op" name="rb_op" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fOP.rb_op[<%=Cstr(idx-1)%>].click(); if (fOP.rb_op[<%=Cstr(idx-1)%>].checked) fOP.bEXECUTAR.click();"
				>Veículos de Mídia</span><br>

			<%	idx=idx+1
				if operacao_permitida(OP_CEN_CAD_MENSAGEM_ALERTA_PRODUTOS, s_lista_operacoes_permitidas) then s="" else s=" disabled" %>
			<input type="radio" id="rb_op" name="rb_op" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fOP.rb_op[<%=Cstr(idx-1)%>].click(); if (fOP.rb_op[<%=Cstr(idx-1)%>].checked) fOP.bEXECUTAR.click();"
				>Mensagens de Alerta para Produtos</span><br>
			
			<%	idx=idx+1
				if operacao_permitida(OP_CEN_CAD_EC_PRODUTO_COMPOSTO, s_lista_operacoes_permitidas) then s="" else s=" disabled" %>
			<input type="radio" id="rb_op" name="rb_op" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fOP.rb_op[<%=Cstr(idx-1)%>].click(); if (fOP.rb_op[<%=Cstr(idx-1)%>].checked) fOP.bEXECUTAR.click();"
				>E-Commerce: Cadastro de Produto Composto</span><br>
			
			<%	idx=idx+1
				if operacao_permitida(OP_CEN_CADASTRO_AVISOS, s_lista_operacoes_permitidas) then s="" else s=" disabled" %>
			<input type="radio" id="rb_op" name="rb_op" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fOP.rb_op[<%=Cstr(idx-1)%>].click(); if (fOP.rb_op[<%=Cstr(idx-1)%>].checked) fOP.bEXECUTAR.click();"
				>Quadro de Avisos</span><br>
			
			<%	idx=idx+1
				if operacao_permitida(OP_CEN_CAD_CEP, s_lista_operacoes_permitidas) then s="" else s=" disabled" %>
			<input type="radio" id="rb_op" name="rb_op" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fOP.rb_op[<%=Cstr(idx-1)%>].click(); if (fOP.rb_op[<%=Cstr(idx-1)%>].checked) fOP.bEXECUTAR.click();"
				>CEP</span> &nbsp; <input id="c_cep" name="c_cep" <%=s%> maxlength="9" size="11" onkeypress="if (digitou_enter(true)) {fOP.bEXECUTAR.click();} filtra_cep();fOP.rb_op[<%=Cstr(idx-1)%>].checked=true;" onblur="if (cep_ok(this.value)) this.value=cep_formata(this.value);" onpaste="fOP.rb_op[<%=Cstr(idx-1)%>].checked=true;"><br>

			<%	idx=idx+1
				if operacao_permitida(OP_CEN_CAD_INDICADOR_OPCOES_FORMA_COMO_CONHECEU, s_lista_operacoes_permitidas) then s="" else s=" disabled" %>
			<input type="radio" id="rb_op" name="rb_op" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fOP.rb_op[<%=Cstr(idx-1)%>].click(); if (fOP.rb_op[<%=Cstr(idx-1)%>].checked) fOP.bEXECUTAR.click();"
				>Opções para "Forma como conheceu a Bonshop" (Cadastro de Indicadores)</span><br>

			<%	idx=idx+1
				if operacao_permitida(OP_CEN_OPCOES_PAGTO_VISANET, s_lista_operacoes_permitidas) then s="" else s=" disabled" %>
			<input type="radio" id="rb_op" name="rb_op" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fOP.rb_op[<%=Cstr(idx-1)%>].click(); if (fOP.rb_op[<%=Cstr(idx-1)%>].checked) fOP.bEXECUTAR.click();"
				>Opções de Pagamento para Cartão de Crédito</span><br>

			<%	idx=idx+1
				if operacao_permitida(OP_CEN_CAD_VL_APROV_AUTO_ANALISE_CREDITO, s_lista_operacoes_permitidas) then s="" else s=" disabled" %>
			<input type="radio" id="rb_op" name="rb_op" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fOP.rb_op[<%=Cstr(idx-1)%>].click(); if (fOP.rb_op[<%=Cstr(idx-1)%>].checked) fOP.bEXECUTAR.click();"
				>Análise de Crédito</span><br>

			<%	idx=idx+1
				if operacao_permitida(OP_CEN_CAD_PERC_LIMITE_RA_SEM_DESAGIO, s_lista_operacoes_permitidas) then s="" else s=" disabled" %>
			<input type="radio" id="rb_op" name="rb_op" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fOP.rb_op[<%=Cstr(idx-1)%>].click(); if (fOP.rb_op[<%=Cstr(idx-1)%>].checked) fOP.bEXECUTAR.click();"
				>Percentual Limite de RA Sem Deságio</span><br>

			<%	idx=idx+1
				if operacao_permitida(OP_CEN_CAD_PERC_MAX_RT, s_lista_operacoes_permitidas) then s="" else s=" disabled" %>
			<!-- DESATIVADO POR TER SIDO SUBSTITUÍDO POR 'Percentual Máximo de Comissão e Desconto por Loja' -->
			<div style="visibility:hidden;position:absolute;">
			<input type="radio" id="rb_op" name="rb_op" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fOP.rb_op[<%=Cstr(idx-1)%>].click(); if (fOP.rb_op[<%=Cstr(idx-1)%>].checked) fOP.bEXECUTAR.click();"
				>Percentual Máximo de Comissão</span><br>
			</div>

			<%	idx=idx+1
				if operacao_permitida(OP_CEN_CAD_PERC_MAX_DESC_SEM_ZERAR_RT, s_lista_operacoes_permitidas) then s="" else s=" disabled" %>
			<!-- DESATIVADO POR TER SIDO SUBSTITUÍDO POR 'Percentual Máximo de Comissão e Desconto por Loja' -->
			<div style="visibility:hidden;position:absolute;">
			<input type="radio" id="rb_op" name="rb_op" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fOP.rb_op[<%=Cstr(idx-1)%>].click(); if (fOP.rb_op[<%=Cstr(idx-1)%>].checked) fOP.bEXECUTAR.click();"
				>Desconto Máximo Aceito Sem Zerar a Comissão</span><br>
			</div>

			<%	idx=idx+1
				if operacao_permitida(OP_CEN_CAD_PERC_MAX_COMISSAO_E_DESC_POR_LOJA, s_lista_operacoes_permitidas) then s="" else s=" disabled" %>
			<input type="radio" id="rb_op" name="rb_op" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fOP.rb_op[<%=Cstr(idx-1)%>].click(); if (fOP.rb_op[<%=Cstr(idx-1)%>].checked) fOP.bEXECUTAR.click();"
				>Percentual Máximo de Comissão e Desconto por Loja</span><br>

			<%	idx=idx+1
				if operacao_permitida(OP_CEN_CAD_PARAMETROS_GLOBAIS, s_lista_operacoes_permitidas) then s="" else s=" disabled" %>
			<input type="radio" id="rb_op" name="rb_op" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fOP.rb_op[<%=Cstr(idx-1)%>].click(); if (fOP.rb_op[<%=Cstr(idx-1)%>].checked) fOP.bEXECUTAR.click();"
				>Percentual Máximo da Senha de Desconto p/ Cadastramento na Loja</span><br>

			<%	idx=idx+1
				if operacao_permitida(OP_CEN_CAD_PARAMETROS_GLOBAIS, s_lista_operacoes_permitidas) then s="" else s=" disabled" %>
			<input type="radio" id="rb_op" name="rb_op" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fOP.rb_op[<%=Cstr(idx-1)%>].click(); if (fOP.rb_op[<%=Cstr(idx-1)%>].checked) fOP.bEXECUTAR.click();"
				>Período de Consulta em Filtros Limitado por Perfil de Acesso</span><br>

			<%	idx=idx+1
				if operacao_permitida(OP_CEN_CAD_PARAMETROS_GLOBAIS, s_lista_operacoes_permitidas) then s="" else s=" disabled" %>
			<input type="radio" id="rb_op" name="rb_op" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fOP.rb_op[<%=Cstr(idx-1)%>].click(); if (fOP.rb_op[<%=Cstr(idx-1)%>].checked) fOP.bEXECUTAR.click();"
				>Limitar a RA a um Percentual do Valor do Pedido</span><br>

			<%	idx=idx+1
				if operacao_permitida(OP_CEN_CAD_TABELA_COMISSAO_VENDEDOR, s_lista_operacoes_permitidas) then s="" else s=" disabled" %>
			<input type="radio" id="rb_op" name="rb_op" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fOP.rb_op[<%=Cstr(idx-1)%>].click(); if (fOP.rb_op[<%=Cstr(idx-1)%>].checked) fOP.bEXECUTAR.click();"
				>Tabela de Comissão do Vendedor</span><br>

			<%	idx=idx+1
				if operacao_permitida(OP_CEN_CAD_TABELA_CUSTO_FINANCEIRO_FORNECEDOR, s_lista_operacoes_permitidas) then s="" else s=" disabled" %>
			<input type="radio" id="rb_op" name="rb_op" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fOP.rb_op[<%=Cstr(idx-1)%>].click(); if (fOP.rb_op[<%=Cstr(idx-1)%>].checked) fOP.bEXECUTAR.click();"
				>Tabela de Custo Financeiro por Fornecedor</span><br>

			<%	idx=idx+1
				if operacao_permitida(OP_CEN_MULTI_CD_CADASTRO_REGRAS_CONSUMO_ESTOQUE, s_lista_operacoes_permitidas) then s="" else s=" disabled" %>
			<input type="radio" id="rb_op" name="rb_op" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fOP.rb_op[<%=Cstr(idx-1)%>].click(); if (fOP.rb_op[<%=Cstr(idx-1)%>].checked) fOP.bEXECUTAR.click();"
				>Multi CD: Cadastro de Regras do Consumo do Estoque</span><br>

			<%	idx=idx+1
				if operacao_permitida(OP_CEN_MULTI_CD_ASSOCIACAO_PRODUTO_REGRA, s_lista_operacoes_permitidas) then s="" else s=" disabled" %>
			<input type="radio" id="rb_op" name="rb_op" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fOP.rb_op[<%=Cstr(idx-1)%>].click(); if (fOP.rb_op[<%=Cstr(idx-1)%>].checked) fOP.bEXECUTAR.click();"
				>Multi CD: Associação do Produto com a Regra</span><br>

			<%	idx=idx+1
				if operacao_permitida(OP_CEN_FIN_CADASTRO_PLANO_CONTAS, s_lista_operacoes_permitidas) then s="" else s=" disabled" %>
			<input type="radio" id="rb_op" name="rb_op" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fOP.rb_op[<%=Cstr(idx-1)%>].click(); if (fOP.rb_op[<%=Cstr(idx-1)%>].checked) fOP.bEXECUTAR.click();"
				>Financeiro: Cadastro de Conta Corrente</span><br>

			<%	idx=idx+1
				if operacao_permitida(OP_CEN_FIN_CADASTRO_PLANO_CONTAS, s_lista_operacoes_permitidas) then s="" else s=" disabled" %>
			<input type="radio" id="rb_op" name="rb_op" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fOP.rb_op[<%=Cstr(idx-1)%>].click(); if (fOP.rb_op[<%=Cstr(idx-1)%>].checked) fOP.bEXECUTAR.click();"
				>Financeiro: Plano de Contas - Empresa</span><br>

			<%	idx=idx+1
				if operacao_permitida(OP_CEN_FIN_CADASTRO_PLANO_CONTAS, s_lista_operacoes_permitidas) then s="" else s=" disabled" %>
			<input type="radio" id="rb_op" name="rb_op" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fOP.rb_op[<%=Cstr(idx-1)%>].click(); if (fOP.rb_op[<%=Cstr(idx-1)%>].checked) fOP.bEXECUTAR.click();"
				>Financeiro: Plano de Contas - Grupo</span><br>

			<%	idx=idx+1
				if operacao_permitida(OP_CEN_FIN_CADASTRO_PLANO_CONTAS, s_lista_operacoes_permitidas) then s="" else s=" disabled" %>
			<input type="radio" id="rb_op" name="rb_op" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fOP.rb_op[<%=Cstr(idx-1)%>].click(); if (fOP.rb_op[<%=Cstr(idx-1)%>].checked) fOP.bEXECUTAR.click();"
				>Financeiro: Plano de Contas - Conta</span><br>

			<%	idx=idx+1
				if operacao_permitida(OP_CEN_FIN_CADASTRO_PLANO_CONTAS, s_lista_operacoes_permitidas) then s="" else s=" disabled" %>
			<input type="radio" id="rb_op" name="rb_op" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fOP.rb_op[<%=Cstr(idx-1)%>].click(); if (fOP.rb_op[<%=Cstr(idx-1)%>].checked) fOP.bEXECUTAR.click();"
				>Financeiro: Boleto - Parâmetros do Cedente</span><br>

			<%	idx=idx+1
				if operacao_permitida(OP_CEN_FIN_CADASTRO_PLANO_CONTAS, s_lista_operacoes_permitidas) then s="" else s=" disabled" %>
			<input type="radio" id="rb_op" name="rb_op" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fOP.rb_op[<%=Cstr(idx-1)%>].click(); if (fOP.rb_op[<%=Cstr(idx-1)%>].checked) fOP.bEXECUTAR.click();"
				>Financeiro: Unidades de Negócio</span><br>

			<%	idx=idx+1
				if operacao_permitida(OP_CEN_FIN_CADASTRO_PLANO_CONTAS, s_lista_operacoes_permitidas) then s="" else s=" disabled" %>
			<input type="radio" id="rb_op" name="rb_op" value="<%=Cstr(idx)%>" class="CBOX" <%=s%>><span class="rbLink" onclick="fOP.rb_op[<%=Cstr(idx-1)%>].click(); if (fOP.rb_op[<%=Cstr(idx-1)%>].checked) fOP.bEXECUTAR.click();"
				>Financeiro: Rateio entre Unidades de Negócio</span><br>

			</td>
		</tr>
	</table>

	<span class="R" style="margin: 4 10 0 10">&nbsp;</span>
	<input name="bEXECUTAR" type="SUBMIT" class="Botao" value="EXECUTAR" title="executa">
	<p class="R" style="margin: 0 10 0 10">&nbsp;</p>

</div>
</form>


<p class="TracoBottom"></p>

<table cellspacing="0">
<tr>
	<td align="center"><a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="volta para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
</tr>
</table>

</center>

</body>
</html>

<%
'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>
