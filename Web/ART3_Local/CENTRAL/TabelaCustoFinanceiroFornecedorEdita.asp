<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =============================================
'	  TabelaCustoFinanceiroFornecedorEdita.asp
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

	dim i
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	if Not operacao_permitida(OP_CEN_CAD_TABELA_CUSTO_FINANCEIRO_FORNECEDOR, s_lista_operacoes_permitidas) then 
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	dim c_fabricante, ckb_clonar_tabela, c_fabricante_a_clonar
	c_fabricante = Trim(Request("c_fabricante"))
	ckb_clonar_tabela = Trim(Request("ckb_clonar_tabela"))
	c_fabricante_a_clonar = Trim(Request("c_fabricante_a_clonar"))
	
	dim alerta
	alerta = ""

	if c_fabricante = "" then
		alerta = "Informe um fornecedor."
		end if

	if alerta = "" then
		if ckb_clonar_tabela = "S" then
			if c_fabricante_a_clonar = "" then
				alerta = "Não foi informado o fornecedor que deve ser usado como base para a clonagem da tabela de custo financeiro"
				end if
			end if
		end if
		
	dim strComEntradaQtdeParcelas, strComEntradaCoeficiente, strSemEntradaQtdeParcelas, strSemEntradaCoeficiente
	dim strSql
	dim cn, rs
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	dim blnCadastrado

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

<script language="JavaScript" type="text/javascript">
function ComEntradaMoverParaCima( f, intIndice ) {
var qtde_parcelas_aux, coeficiente_aux;
	if (intIndice==0) return;
	qtde_parcelas_aux=f.c_com_entrada_qtde_parcelas[intIndice-1].value;
	coeficiente_aux=f.c_com_entrada_coeficiente[intIndice-1].value;
	f.c_com_entrada_qtde_parcelas[intIndice-1].value=f.c_com_entrada_qtde_parcelas[intIndice].value;
	f.c_com_entrada_coeficiente[intIndice-1].value=f.c_com_entrada_coeficiente[intIndice].value;
	f.c_com_entrada_qtde_parcelas[intIndice].value=qtde_parcelas_aux;
	f.c_com_entrada_coeficiente[intIndice].value=coeficiente_aux;
}

function ComEntradaMoverParaBaixo( f, intIndice ) {
var qtde_parcelas_aux, coeficiente_aux;
	if (intIndice==(f.c_com_entrada_qtde_parcelas.length-1)) return;
	qtde_parcelas_aux=f.c_com_entrada_qtde_parcelas[intIndice+1].value;
	coeficiente_aux=f.c_com_entrada_coeficiente[intIndice+1].value;
	f.c_com_entrada_qtde_parcelas[intIndice+1].value=f.c_com_entrada_qtde_parcelas[intIndice].value;
	f.c_com_entrada_coeficiente[intIndice+1].value=f.c_com_entrada_coeficiente[intIndice].value;
	f.c_com_entrada_qtde_parcelas[intIndice].value=qtde_parcelas_aux;
	f.c_com_entrada_coeficiente[intIndice].value=coeficiente_aux;
}

function ComEntradaIncluiNovaLinha( f, intIndice ) {
var i, blnUltLinhaTemDados;
	if (intIndice==(f.c_com_entrada_qtde_parcelas.length-1)) return;
	blnUltLinhaTemDados=false;
	if (trim(f.c_com_entrada_qtde_parcelas[f.c_com_entrada_qtde_parcelas.length-1].value)!="") blnUltLinhaTemDados=true;
	if (trim(f.c_com_entrada_coeficiente[f.c_com_entrada_coeficiente.length-1].value)!="") blnUltLinhaTemDados=true;
	if (blnUltLinhaTemDados) {
		if (!confirm("Os dados da última linha serão perdidos!!\nContinua?")) return;
		}
	for (i=(f.c_com_entrada_qtde_parcelas.length-1); i>intIndice; i--) {
		f.c_com_entrada_qtde_parcelas[i].value=f.c_com_entrada_qtde_parcelas[i-1].value;
		f.c_com_entrada_coeficiente[i].value=f.c_com_entrada_coeficiente[i-1].value;
		}
	f.c_com_entrada_qtde_parcelas[intIndice].value="";
	f.c_com_entrada_coeficiente[intIndice].value="";
}

function ComEntradaRemoveLinha( f, intIndice ) {
var i;
	if (!confirm("Exclui esta linha?")) return;
	for (i=intIndice; i < (f.c_com_entrada_qtde_parcelas.length-1); i++) {
		f.c_com_entrada_qtde_parcelas[i].value=f.c_com_entrada_qtde_parcelas[i+1].value;
		f.c_com_entrada_coeficiente[i].value=f.c_com_entrada_coeficiente[i+1].value;
		}
	f.c_com_entrada_qtde_parcelas[f.c_com_entrada_qtde_parcelas.length-1].value="";
	f.c_com_entrada_coeficiente[f.c_com_entrada_coeficiente.length-1].value="";
}

function SemEntradaMoverParaCima( f, intIndice ) {
var qtde_parcelas_aux, coeficiente_aux;
	if (intIndice==0) return;
	qtde_parcelas_aux=f.c_sem_entrada_qtde_parcelas[intIndice-1].value;
	coeficiente_aux=f.c_sem_entrada_coeficiente[intIndice-1].value;
	f.c_sem_entrada_qtde_parcelas[intIndice-1].value=f.c_sem_entrada_qtde_parcelas[intIndice].value;
	f.c_sem_entrada_coeficiente[intIndice-1].value=f.c_sem_entrada_coeficiente[intIndice].value;
	f.c_sem_entrada_qtde_parcelas[intIndice].value=qtde_parcelas_aux;
	f.c_sem_entrada_coeficiente[intIndice].value=coeficiente_aux;
}

function SemEntradaMoverParaBaixo( f, intIndice ) {
var qtde_parcelas_aux, coeficiente_aux;
	if (intIndice==(f.c_sem_entrada_qtde_parcelas.length-1)) return;
	qtde_parcelas_aux=f.c_sem_entrada_qtde_parcelas[intIndice+1].value;
	coeficiente_aux=f.c_sem_entrada_coeficiente[intIndice+1].value;
	f.c_sem_entrada_qtde_parcelas[intIndice+1].value=f.c_sem_entrada_qtde_parcelas[intIndice].value;
	f.c_sem_entrada_coeficiente[intIndice+1].value=f.c_sem_entrada_coeficiente[intIndice].value;
	f.c_sem_entrada_qtde_parcelas[intIndice].value=qtde_parcelas_aux;
	f.c_sem_entrada_coeficiente[intIndice].value=coeficiente_aux;
}

function SemEntradaIncluiNovaLinha( f, intIndice ) {
var i, blnUltLinhaTemDados;
	if (intIndice==(f.c_sem_entrada_qtde_parcelas.length-1)) return;
	blnUltLinhaTemDados=false;
	if (trim(f.c_sem_entrada_qtde_parcelas[f.c_sem_entrada_qtde_parcelas.length-1].value)!="") blnUltLinhaTemDados=true;
	if (trim(f.c_sem_entrada_coeficiente[f.c_sem_entrada_coeficiente.length-1].value)!="") blnUltLinhaTemDados=true;
	if (blnUltLinhaTemDados) {
		if (!confirm("Os dados da última linha serão perdidos!!\nContinua?")) return;
		}
	for (i=(f.c_sem_entrada_qtde_parcelas.length-1); i>intIndice; i--) {
		f.c_sem_entrada_qtde_parcelas[i].value=f.c_sem_entrada_qtde_parcelas[i-1].value;
		f.c_sem_entrada_coeficiente[i].value=f.c_sem_entrada_coeficiente[i-1].value;
		}
	f.c_sem_entrada_qtde_parcelas[intIndice].value="";
	f.c_sem_entrada_coeficiente[intIndice].value="";
}

function SemEntradaRemoveLinha( f, intIndice ) {
var i;
	if (!confirm("Exclui esta linha?")) return;
	for (i=intIndice; i < (f.c_sem_entrada_qtde_parcelas.length-1); i++) {
		f.c_sem_entrada_qtde_parcelas[i].value=f.c_sem_entrada_qtde_parcelas[i+1].value;
		f.c_sem_entrada_coeficiente[i].value=f.c_sem_entrada_coeficiente[i+1].value;
		}
	f.c_sem_entrada_qtde_parcelas[f.c_sem_entrada_qtde_parcelas.length-1].value="";
	f.c_sem_entrada_coeficiente[f.c_sem_entrada_coeficiente.length-1].value="";
}

function fCadRemove( f ) {
	if (!confirm("Exclui esta tabela de custo financeiro?")) return;
	dREMOVE.style.visibility="hidden";
	window.status = "Aguarde ...";
	f.operacao_selecionada.value=OP_EXCLUI;
	f.submit();
}

function fCadConfirma( f ) {
var i, b, ha_item;
	ha_item=false;
	for (i=0; i < f.c_com_entrada_qtde_parcelas.length; i++) {
		b=false;
		if (trim(f.c_com_entrada_coeficiente[i].value)!="") b=true;
		if (b) {
			ha_item=true;
			if (converte_numero(f.c_com_entrada_qtde_parcelas[i].value)<=0) {
				alert("Quantidade de parcelas inválida!!");
				f.c_com_entrada_qtde_parcelas[i].focus();
				return;
				}
			if (converte_numero(f.c_com_entrada_coeficiente[i].value)<=0) {
				alert("Coeficiente inválido!!");
				f.c_com_entrada_coeficiente[i].focus();
				return;
				}
			}
		}

	if (!ha_item) {
		alert("Nenhuma opção de parcelamento com entrada foi informado!!");
		f.c_com_entrada_coeficiente[0].focus();
		return;
		}

	ha_item=false;
	for (i=0; i < f.c_sem_entrada_qtde_parcelas.length; i++) {
		b=false;
		if (trim(f.c_sem_entrada_coeficiente[i].value)!="") b=true;
		if (b) {
			ha_item=true;
			if (converte_numero(f.c_sem_entrada_qtde_parcelas[i].value)<=0) {
				alert("Quantidade de parcelas inválida!!");
				f.c_sem_entrada_qtde_parcelas[i].focus();
				return;
				}
			if (converte_numero(f.c_sem_entrada_coeficiente[i].value)<=0) {
				alert("Coeficiente inválido!!");
				f.c_sem_entrada_coeficiente[i].focus();
				return;
				}
			}
		}

	if (!ha_item) {
		alert("Nenhuma opção de parcelamento sem entrada foi informado!!");
		f.c_sem_entrada_coeficiente[0].focus();
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
<body onload="if (trim(fCad.c_sem_entrada_coeficiente[0].value)=='') fCad.c_sem_entrada_coeficiente[0].focus();">
<center>

<form id="fCad" name="fCad" method="post" action="TabelaCustoFinanceiroFornecedorConfirma.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="c_fabricante" id="c_fabricante" value="<%=c_fabricante%>">

<% 
'	TABELA CADASTRADA?
	strSql = "SELECT * FROM t_PERCENTUAL_CUSTO_FINANCEIRO_FORNECEDOR WHERE (fabricante = '" & c_fabricante & "')"
	set rs = cn.Execute(strSql)
	blnCadastrado = False
	if Not rs.Eof then blnCadastrado = True
%>
<% if blnCadastrado then %>
<input type="hidden" name="operacao_selecionada" id="operacao_selecionada" value="<%=OP_CONSULTA%>">
<% else %>
<input type="hidden" name="operacao_selecionada" id="operacao_selecionada" value="<%=OP_INCLUI%>">
<% end if %>

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Tabela de Custo Financeiro por Fornecedor</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>
<br>

<!--  TABELA DE CUSTO FINANCEIRO  -->
<table class="Qx" cellSpacing="0">
	<!--  FORNECEDOR  -->
	<tr bgColor="#FFFFFF">
		<td colspan="3" align="center" class="MT"><span class="PLTc">FORNECEDOR</span>
		<br><p class="Cc"><%=c_fabricante & " - " & fabricante_descricao(c_fabricante)%></p>
		</td>
	</tr>

	<tr bgColor="#FFFFFF">
	<td colspan="3">&nbsp;</td>
	</tr>
	
	<tr>

		<td>
		<% 
		'	TABELA SEM ENTRADA
			strSql = "SELECT * FROM t_PERCENTUAL_CUSTO_FINANCEIRO_FORNECEDOR WHERE (fabricante = '" & c_fabricante & "') AND (tipo_parcelamento = '" & COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__SEM_ENTRADA & "') ORDER BY qtde_parcelas"
			set rs = cn.Execute(strSql)
			blnCadastrado = False
			if Not rs.Eof then blnCadastrado = True
			
			if Not blnCadastrado then
				if ckb_clonar_tabela = "S" then
					strSql = "SELECT * FROM t_PERCENTUAL_CUSTO_FINANCEIRO_FORNECEDOR WHERE (fabricante = '" & c_fabricante_a_clonar & "') AND (tipo_parcelamento = '" & COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__SEM_ENTRADA & "') ORDER BY qtde_parcelas"
					set rs = cn.Execute(strSql)
					end if
				end if
		%>

		<table cellSpacing="0">
			<!--  TABELA SEM ENTRADA  -->
			<tr bgColor="#FFFFFF">
			<td>&nbsp;</td>
			<td colspan="2" class="ME MD MC"><p class="PLTc">SEM ENTRADA</p></td>
			<td colspan="4">&nbsp;</td>
			</tr>
			<tr bgColor="#FFFFFF">
			<td>&nbsp;</td>
			<td class="MDBE MC"><p class="PLTd">Parcelas</p></td>
			<td class="MDB MC"><p class="PLTd">Coeficiente</p></td>
			<td colspan="4">&nbsp;</td>
			</tr>
		<% for i=1 to MAX_LINHAS_TABELA_CUSTO_FINANCEIRO_FORNECEDOR %>
			<%
				if Not rs.Eof then
					strSemEntradaQtdeParcelas = Cstr(rs("qtde_parcelas"))
					strSemEntradaCoeficiente = formata_coeficiente_custo_financ_fornecedor(rs("coeficiente"))
				else
					strSemEntradaQtdeParcelas = Cstr(i)
					strSemEntradaCoeficiente = ""
					end if
			%>
			<tr>
			<td>
				<input name="c_linha" id="c_linha" readonly tabindex=-1 class="PLLe" maxlength="2" style="width:30px;text-align:right;color:#808080;" 
					value="<%="0 + " %>"></td>
			<td class="MDBE" align="right">
				<input name="c_sem_entrada_qtde_parcelas" id="c_sem_entrada_qtde_parcelas" class="PLLc" maxlength="2" style="width:50px;" 
					value="<%=strSemEntradaQtdeParcelas%>"
					readonly tabindex=-1
					onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fCad.c_sem_entrada_coeficiente[<%=Cstr(i-1)%>].focus(); filtra_numerico();" 
					onblur="this.value=formata_inteiro(this.value);"></td>
			<td class="MDB" align="right">
				<input name="c_sem_entrada_coeficiente" id="c_sem_entrada_coeficiente" class="PLLd" maxlength="12" style="width:70px;" 
					value="<%=strSemEntradaCoeficiente%>"
					onkeypress="if (digitou_enter(true)&&tem_info(this.value)) {if (<%=Cstr(i)%>==fCad.c_sem_entrada_coeficiente.length) fCad.c_com_entrada[0].focus(); else fCad.c_sem_entrada_coeficiente[<%=Cstr(i)%>].focus();} filtra_coeficiente_custo_financ_fornecedor();"
					onblur="this.value=formata_coeficiente_custo_financ_fornecedor(this.value);"></td>
			<td>
				<% if i = 1 then %>
					&nbsp;
				<% else %>
				<a name="bSetaCima" id="bSetaCima" href="javascript:SemEntradaMoverParaCima(fCad,<%=Cstr(i-1)%>)" title="move para cima"
					tabindex=-1>
					<img src="../botao/SetaCima.gif" style="vertical-align:bottom;margin-left:4px;margin-bottom:1px;" border="0"></a>
				<% end if %>
				</td>
			<td>
				<% if i = MAX_LINHAS_TABELA_CUSTO_FINANCEIRO_FORNECEDOR then %>
					&nbsp;
				<% else %>
				<a name="bSetaBaixo" id="bSetaBaixo" href="javascript:SemEntradaMoverParaBaixo(fCad,<%=Cstr(i-1)%>)" title="move para baixo"
					tabindex=-1>
					<img src="../botao/SetaBaixo.gif" style="vertical-align:bottom;margin-left:4px;margin-bottom:1px;" border="0"></a>
				<% end if %>
				</td>
			<td>
				<% if i = MAX_LINHAS_TABELA_CUSTO_FINANCEIRO_FORNECEDOR then %>
					&nbsp;
				<% else %>
				<a name="bNovaLinha" id="bNovaLinha" href="javascript:SemEntradaIncluiNovaLinha(fCad,<%=Cstr(i-1)%>)" title="inclui uma nova linha"
					tabindex=-1>
					<img src="../botao/Adicionar.gif" style="vertical-align:bottom;margin-left:4px;margin-bottom:1px;" border="0"></a>
				<% end if %>
				</td>
			<td>
				<a name="bRemoveLinha" id="bRemoveLinha" href="javascript:SemEntradaRemoveLinha(fCad,<%=Cstr(i-1)%>)" title="remove a linha"
					tabindex=-1>
					<img src="../botao/botao_x_red.gif" style="vertical-align:bottom;margin-left:4px;margin-bottom:1px;" border="0"></a>
				</td>
			</tr>
			<% if Not rs.Eof then rs.MoveNext %>
		<% next %>
		</table>
		</td>

		<!--  **********  COLUNA SEPARADORA  **********  -->
		<td><span style="width:20px;">&nbsp;</span></td>
		
		<td>
		<% 
		'	TABELA COM ENTRADA
			strSql = "SELECT * FROM t_PERCENTUAL_CUSTO_FINANCEIRO_FORNECEDOR WHERE (fabricante = '" & c_fabricante & "') AND (tipo_parcelamento = '" & COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__COM_ENTRADA & "') ORDER BY qtde_parcelas"
			set rs = cn.Execute(strSql)
			blnCadastrado = False
			if Not rs.Eof then blnCadastrado = True
			
			if Not blnCadastrado then
				if ckb_clonar_tabela = "S" then
					strSql = "SELECT * FROM t_PERCENTUAL_CUSTO_FINANCEIRO_FORNECEDOR WHERE (fabricante = '" & c_fabricante_a_clonar & "') AND (tipo_parcelamento = '" & COD_CUSTO_FINANC_FORNEC_TIPO_PARCELAMENTO__COM_ENTRADA & "') ORDER BY qtde_parcelas"
					set rs = cn.Execute(strSql)
					end if
				end if
		%>
		
		<table cellSpacing="0">
			<!--  TABELA COM ENTRADA  -->
			<tr bgColor="#FFFFFF">
			<td>&nbsp;</td>
			<td colspan="2" class="ME MD MC"><p class="PLTc">COM ENTRADA</p></td>
			<td colspan="4">&nbsp;</td>
			</tr>
			<tr bgColor="#FFFFFF">
			<td>&nbsp;</td>
			<td class="MDBE MC"><p class="PLTd">Parcelas</p></td>
			<td class="MDB MC"><p class="PLTd">Coeficiente</p></td>
			<td colspan="4">&nbsp;</td>
			</tr>
		<% for i=1 to MAX_LINHAS_TABELA_CUSTO_FINANCEIRO_FORNECEDOR %>
			<%
				if Not rs.Eof then
					strComEntradaQtdeParcelas = Cstr(rs("qtde_parcelas"))
					strComEntradaCoeficiente = formata_coeficiente_custo_financ_fornecedor(rs("coeficiente"))
				else
					strComEntradaQtdeParcelas = Cstr(i)
					strComEntradaCoeficiente = ""
					end if
			%>
			<tr>
			<td>
				<input name="c_linha" id="c_linha" readonly tabindex=-1 class="PLLe" maxlength="2" style="width:30px;text-align:right;color:#808080;" 
					value="<%="1 + " %>"></td>
			<td class="MDBE" align="right">
				<input name="c_com_entrada_qtde_parcelas" id="c_com_entrada_qtde_parcelas" class="PLLc" maxlength="2" style="width:50px;" 
					value="<%=strComEntradaQtdeParcelas%>"
					readonly tabindex=-1
					onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fCad.c_com_entrada_coeficiente[<%=Cstr(i-1)%>].focus(); filtra_numerico();" 
					onblur="this.value=formata_inteiro(this.value);"></td>
			<td class="MDB" align="right">
				<input name="c_com_entrada_coeficiente" id="c_com_entrada_coeficiente" class="PLLd" maxlength="12" style="width:70px;" 
					value="<%=strComEntradaCoeficiente%>"
					onkeypress="if (digitou_enter(true)&&tem_info(this.value)) {if (<%=Cstr(i)%>==fCad.c_com_entrada_coeficiente.length) bCONFIRMA.focus(); else fCad.c_com_entrada_coeficiente[<%=Cstr(i)%>].focus();} filtra_coeficiente_custo_financ_fornecedor();"
					onblur="this.value=formata_coeficiente_custo_financ_fornecedor(this.value);"></td>
			<td>
				<% if i = 1 then %>
					&nbsp;
				<% else %>
				<a name="bSetaCima" id="bSetaCima" href="javascript:ComEntradaMoverParaCima(fCad,<%=Cstr(i-1)%>)" title="move para cima"
					tabindex=-1>
					<img src="../botao/SetaCima.gif" style="vertical-align:bottom;margin-left:4px;margin-bottom:1px;" border="0"></a>
				<% end if %>
				</td>
			<td>
				<% if i = MAX_LINHAS_TABELA_CUSTO_FINANCEIRO_FORNECEDOR then %>
					&nbsp;
				<% else %>
				<a name="bSetaBaixo" id="bSetaBaixo" href="javascript:ComEntradaMoverParaBaixo(fCad,<%=Cstr(i-1)%>)" title="move para baixo"
					tabindex=-1>
					<img src="../botao/SetaBaixo.gif" style="vertical-align:bottom;margin-left:4px;margin-bottom:1px;" border="0"></a>
				<% end if %>
				</td>
			<td>
				<% if i = MAX_LINHAS_TABELA_CUSTO_FINANCEIRO_FORNECEDOR then %>
					&nbsp;
				<% else %>
				<a name="bNovaLinha" id="bNovaLinha" href="javascript:ComEntradaIncluiNovaLinha(fCad,<%=Cstr(i-1)%>)" title="inclui uma nova linha"
					tabindex=-1>
					<img src="../botao/Adicionar.gif" style="vertical-align:bottom;margin-left:4px;margin-bottom:1px;" border="0"></a>
				<% end if %>
				</td>
			<td>
				<a name="bRemoveLinha" id="bRemoveLinha" href="javascript:ComEntradaRemoveLinha(fCad,<%=Cstr(i-1)%>)" title="remove a linha"
					tabindex=-1>
					<img src="../botao/botao_x_red.gif" style="vertical-align:bottom;margin-left:4px;margin-bottom:1px;" border="0"></a>
				</td>
			</tr>
			<% if Not rs.Eof then rs.MoveNext %>
		<% next %>
		</table>
		</td>
		
	</tr>
</table>


<!-- ************   SEPARADOR   ************ -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>
<br>


<table width="649" cellSpacing="0">
<tr>
	<td><a href="TabelaCustoFinanceiroFornecedorFiltro.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="volta para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
	<%if blnCadastrado then%>
	<td align="center"><div name="dREMOVE" id="dREMOVE"><a name="bREMOVE" id="bREMOVE" href="javascript:fCadRemove(fCad)" title="exclui do banco de dados">
		<img src="../botao/remover.gif" width="176" height="55" border="0"></a></div>
	</td>
	<%end if%>
	<td align="right"><div name="dCONFIRMA" id="dCONFIRMA"><a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fCadConfirma(fCad)" title="grava os dados">
		<img src="../botao/confirmar.gif" width="176" height="55" border="0"></a></div>
	</td>
</tr>
</table>
</form>

</center>
</body>
<%end if%>

</html>


<%
'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>