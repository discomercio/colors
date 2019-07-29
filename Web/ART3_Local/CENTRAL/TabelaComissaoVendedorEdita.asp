<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =============================================
'	  TabelaComissaoVendedorEdita.asp
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
	if Not operacao_permitida(OP_CEN_CAD_TABELA_COMISSAO_VENDEDOR, s_lista_operacoes_permitidas) then 
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	dim c_id_perfil, ckb_clonar_tabela, c_perfil_a_clonar
	c_id_perfil = Trim(Request("c_id_perfil"))
	ckb_clonar_tabela = Trim(Request("ckb_clonar_tabela"))
	c_perfil_a_clonar = Trim(Request("c_perfil_a_clonar"))
	
	dim alerta
	alerta = ""

	if c_id_perfil = "" then
		alerta = "Informe um perfil."
		end if

	if alerta = "" then
		if ckb_clonar_tabela = "S" then
			if c_perfil_a_clonar = "" then
				alerta = "Não foi informado o perfil que deve ser usado como base para a clonagem da tabela de comissão"
				end if
			end if
		end if
		
	dim strPercDesconto, strPercComissao, strSql	
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
function MoverParaCima( f, intIndice ) {
var perc_desconto_aux, perc_comissao_aux;
	if (intIndice==0) return;
	perc_desconto_aux=f.c_perc_desconto[intIndice-1].value;
	perc_comissao_aux=f.c_perc_comissao[intIndice-1].value;
	f.c_perc_desconto[intIndice-1].value=f.c_perc_desconto[intIndice].value;
	f.c_perc_comissao[intIndice-1].value=f.c_perc_comissao[intIndice].value;
	f.c_perc_desconto[intIndice].value=perc_desconto_aux;
	f.c_perc_comissao[intIndice].value=perc_comissao_aux;
}

function MoverParaBaixo( f, intIndice ) {
var perc_desconto_aux, perc_comissao_aux;
	if (intIndice==(f.c_perc_desconto.length-1)) return;
	perc_desconto_aux=f.c_perc_desconto[intIndice+1].value;
	perc_comissao_aux=f.c_perc_comissao[intIndice+1].value;
	f.c_perc_desconto[intIndice+1].value=f.c_perc_desconto[intIndice].value;
	f.c_perc_comissao[intIndice+1].value=f.c_perc_comissao[intIndice].value;
	f.c_perc_desconto[intIndice].value=perc_desconto_aux;
	f.c_perc_comissao[intIndice].value=perc_comissao_aux;
}

function IncluiNovaLinha( f, intIndice ) {
var i, blnUltLinhaTemDados;
	if (intIndice==(f.c_perc_desconto.length-1)) return;
	blnUltLinhaTemDados=false;
	if (trim(f.c_perc_desconto[f.c_perc_desconto.length-1].value)!="") blnUltLinhaTemDados=true;
	if (trim(f.c_perc_comissao[f.c_perc_comissao.length-1].value)!="") blnUltLinhaTemDados=true;
	if (blnUltLinhaTemDados) {
		if (!confirm("Os dados da última linha serão perdidos!!\nContinua?")) return;
		}
	for (i=(f.c_perc_desconto.length-1); i>intIndice; i--) {
		f.c_perc_desconto[i].value=f.c_perc_desconto[i-1].value;
		f.c_perc_comissao[i].value=f.c_perc_comissao[i-1].value;
		}
	f.c_perc_desconto[intIndice].value="";
	f.c_perc_comissao[intIndice].value="";
}

function RemoveLinha( f, intIndice ) {
var i;
	if (!confirm("Exclui esta linha?")) return;
	for (i=intIndice; i < (f.c_perc_desconto.length-1); i++) {
		f.c_perc_desconto[i].value=f.c_perc_desconto[i+1].value;
		f.c_perc_comissao[i].value=f.c_perc_comissao[i+1].value;
		}
	f.c_perc_desconto[f.c_perc_desconto.length-1].value="";
	f.c_perc_comissao[f.c_perc_comissao.length-1].value="";
}

function fCadRemove( f ) {
	if (!confirm("Exclui esta tabela de comissão do vendedor?")) return;
	dREMOVE.style.visibility="hidden";
	window.status = "Aguarde ...";
	f.operacao_selecionada.value=OP_EXCLUI;
	f.submit();
}

function fCadConfirma( f ) {
var i, b, ha_item;
	ha_item=false;
	for (i=0; i < f.c_perc_desconto.length; i++) {
		b=false;
		if (trim(f.c_perc_desconto[i].value)!="") b=true;
		if (trim(f.c_perc_comissao[i].value)!="") b=true;
		if (b) {
			ha_item=true;
			if ((converte_numero(f.c_perc_desconto[i].value)<0)||(converte_numero(f.c_perc_desconto[i].value)>100)) {
				alert("Percentual de desconto inválido!!");
				f.c_perc_desconto[i].focus();
				return;
				}
			if ((converte_numero(f.c_perc_comissao[i].value)<0)||(converte_numero(f.c_perc_comissao[i].value)>100)) {
				alert("Percentual de comissão inválido!!");
				f.c_perc_comissao[i].focus();
				return;
				}
			}
		}

	if (!ha_item) {
		alert("Nenhum percentual de comissão foi informado!!");
		f.c_perc_desconto[0].focus();
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
<body onload="if (trim(fCad.c_perc_desconto[0].value)=='') fCad.c_perc_desconto[0].focus();">
<center>

<% 
	strSql = "SELECT * FROM t_PERCENTUAL_COMISSAO_VENDEDOR WHERE id_perfil = '" & c_id_perfil & "' ORDER BY perc_desconto"
	set rs = cn.Execute(strSql)
	blnCadastrado = False
	if Not rs.Eof then blnCadastrado = True
	
	if Not blnCadastrado then
		if ckb_clonar_tabela = "S" then
			strSql = "SELECT * FROM t_PERCENTUAL_COMISSAO_VENDEDOR WHERE id_perfil = '" & c_perfil_a_clonar & "' ORDER BY perc_desconto"
			set rs = cn.Execute(strSql)
			end if
		end if
%>

<form id="fCad" name="fCad" method="post" action="TabelaComissaoVendedorConfirma.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="c_id_perfil" id="c_id_perfil" value="<%=c_id_perfil%>">

<% if blnCadastrado then %>
<input type="hidden" name="operacao_selecionada" id="operacao_selecionada" value="<%=OP_CONSULTA%>">
<% else %>
<input type="hidden" name="operacao_selecionada" id="operacao_selecionada" value="<%=OP_INCLUI%>">
<% end if %>


<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Tabela de Comissão do Vendedor</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>
<br>

<!--  TABELA DE COMISSÃO DOS VENDEDORES  -->
<table class="Qx" cellSpacing="0">
	<!--  PERFIL  -->
	<tr bgColor="#FFFFFF">
		<td>&nbsp;</td>
		<td colspan="2" class="MT"><span class="PLTe">PERFIL</span>
		<br><p class="C"><%=x_perfil_apelido(c_id_perfil)%></p>
		</td>
		<td colspan="4">&nbsp;</td>
	</tr>

	<tr bgColor="#FFFFFF">
	<td colspan="3">&nbsp;</td>
	<td colspan="4">&nbsp;</td>
	</tr>
	
	<!--  TÍTULO DAS COLUNAS  -->
	<tr bgColor="#FFFFFF">
	<td>&nbsp;</td>
	<td class="MDBE MC"><p class="PLTd">Desc (%)</p></td>
	<td class="MDB MC"><p class="PLTd">Comissão (%)</p></td>
	<td colspan="4">&nbsp;</td>
	</tr>
<% for i=1 to MAX_LINHAS_TABELA_COMISSAO_VENDEDOR %>
	<%
		if Not rs.Eof then
			strPercDesconto = formata_perc_desc(rs("perc_desconto"))
			strPercComissao = formata_perc_desc(rs("perc_comissao"))
		else
			strPercDesconto = ""
			strPercComissao = ""
			end if
	%>
	<tr>
	<td>
		<input name="c_linha" id="c_linha" readonly tabindex=-1 class="PLLe" maxlength="2" style="width:30px;text-align:right;color:#808080;" 
			value="<%=Cstr(i) & ". " %>"></td>
	<td class="MDBE">
		<input name="c_perc_desconto" id="c_perc_desconto" class="PLLd" maxlength="5" style="width:80px;" 
			value="<%=strPercDesconto%>"
			onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fCad.c_perc_comissao[<%=Cstr(i-1)%>].focus(); filtra_percentual();" 
			onblur="this.value=formata_perc_desc(this.value);"></td>
	<td class="MDB" align="right">
		<input name="c_perc_comissao" id="c_perc_comissao" class="PLLd" maxlength="5" style="width:80px;" 
			value="<%=strPercComissao%>"
			onkeypress="if (digitou_enter(true)&&tem_info(this.value)) {if (<%=Cstr(i)%>==fCad.c_perc_comissao.length) bCONFIRMA.focus(); else fCad.c_perc_desconto[<%=Cstr(i)%>].focus();} filtra_percentual();"
			onblur="this.value=formata_perc_desc(this.value);"></td>
	<td>
		<% if i = 1 then %>
			&nbsp;
		<% else %>
		<a name="bSetaCima" id="bSetaCima" href="javascript:MoverParaCima(fCad,<%=Cstr(i-1)%>)" title="move para cima">
			<img src="../botao/SetaCima.gif" style="vertical-align:bottom;margin-left:4px;margin-bottom:1px;" border="0"></a>
		<% end if %>
		</td>
	<td>
		<% if i = MAX_LINHAS_TABELA_COMISSAO_VENDEDOR then %>
			&nbsp;
		<% else %>
		<a name="bSetaBaixo" id="bSetaBaixo" href="javascript:MoverParaBaixo(fCad,<%=Cstr(i-1)%>)" title="move para baixo">
			<img src="../botao/SetaBaixo.gif" style="vertical-align:bottom;margin-left:4px;margin-bottom:1px;" border="0"></a>
		<% end if %>
		</td>
	<td>
		<% if i = MAX_LINHAS_TABELA_COMISSAO_VENDEDOR then %>
			&nbsp;
		<% else %>
		<a name="bNovaLinha" id="bNovaLinha" href="javascript:IncluiNovaLinha(fCad,<%=Cstr(i-1)%>)" title="inclui uma nova linha">
			<img src="../botao/Adicionar.gif" style="vertical-align:bottom;margin-left:4px;margin-bottom:1px;" border="0"></a>
		<% end if %>
		</td>
	<td>
		<a name="bRemoveLinha" id="bRemoveLinha" href="javascript:RemoveLinha(fCad,<%=Cstr(i-1)%>)" title="remove a linha">
			<img src="../botao/botao_x_red.gif" style="vertical-align:bottom;margin-left:4px;margin-bottom:1px;" border="0"></a>
		</td>
	</tr>
	<% if Not rs.Eof then rs.MoveNext %>
<% next %>
</table>


<!-- ************   SEPARADOR   ************ -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>
<br>


<table width="649" cellSpacing="0">
<tr>
	<td><a href="TabelaComissaoVendedorFiltro.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="volta para a página anterior">
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