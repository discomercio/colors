<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/Global.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ===============================================
'	  S E N H A D E S C S U P . A S P
'     ===============================================
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

	dim s, s_aux, cliente_selecionado, intCounter
	cliente_selecionado=Trim(request("cliente_selecionado"))
	if cliente_selecionado = "" then Response.Redirect("aviso.asp?id=" & ERR_CLIENTE_NAO_ESPECIFICADO)

'	CONECTA COM O BANCO DE DADOS
	dim cn
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	
	dim max_qtde_itens
	max_qtde_itens = obtem_parametro_SenhaDescontoSuperior_MaxQtdeItens

	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	if Not operacao_permitida(OP_CEN_CADASTRA_SENHA_DESCONTO, s_lista_operacoes_permitidas) then 
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	dim r_cliente
	set r_cliente = New cl_CLIENTE
	if Not x_cliente_bd(cliente_selecionado, r_cliente) then Response.Redirect("aviso.asp?id=" & ERR_CLIENTE_NAO_CADASTRADO)

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
function fCANCConfirma( f ) {
var b;
	b=window.confirm("Confirma o cancelamento das senhas de autorização concedidas para este cliente?");
	if (!b) return;

	dREMOVE.style.visibility="hidden";
	window.status = "Aguarde ...";
	f.submit();
}

function fFILTROConfirma( f ) {
var desc, i, b, ha_item;

	if (trim(f.c_supervisor_autorizador.value)=="") {
		alert("Informe quem está autorizando o desconto!!");
		f.c_supervisor_autorizador.focus();
		return;
		}

	if (trim(f.c_loja.value)=="") {
		alert("Especifique a loja!!");
		f.c_loja.focus();
		return;
		}

	ha_item=false;
	for (i=0; i < f.c_fabricante.length; i++) {
		b=false;
		if (trim(f.c_fabricante[i].value)!="") b=true;
		if (trim(f.c_produto[i].value)!="") b=true;
		if (converte_numero(f.c_desc_max_senha[i].value)!=0) b=true;
		
		if (b) {
			ha_item=true;
			if (trim(f.c_fabricante[i].value)=="") {
				alert("Especifique o código do fabricante!!");
				f.c_fabricante[i].focus();
				return;
				}
			if (trim(f.c_produto[i].value)=="") {
				alert("Especifique o código do produto!!");
				f.c_produto[i].focus();
				return;
				}
			desc=converte_numero(f.c_desc_max_senha[i].value);
			if ((desc<=0)||(desc>100)) {
				alert("Percentual de desconto inválido!!");
				f.c_desc_max_senha[i].focus();
				return;
				}
			}
		}
		
	if (!ha_item) {
		alert("Nenhum desconto foi informado!!");
		f.c_fabricante[0].focus();
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
<link href="<%=URL_FILE__EPRINTER_CSS%>" Rel="stylesheet" Type="text/css" media="print">

<style type="text/css">
#c_supervisor_autorizador {
	margin: 4pt 4pt 4pt 10pt;
	vertical-align: top;
	}
</style>

<body onload="fFILTRO.c_loja.focus();">
<center>

<form id="fCANC" name="fCANC" method="post" action="SenhaDescSupCancela.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="cliente_selecionado" id="cliente_selecionado" value="<%=cliente_selecionado%>">
</form>

<form id="fFILTRO" name="fFILTRO" method="post" action="SenhaDescSupConsiste.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="cliente_selecionado" id="cliente_selecionado" value="<%=cliente_selecionado%>">

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Senha para Autorização de Desconto Superior</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>
<br>

<!--  DESCRIÇÃO DA OPERAÇÃO -->
<table width="649" cellPadding="0" CellSpacing="0">
<tr><td><p class="Expl">OBSERVAÇÕES</p></td></tr>
<tr><td>
	<p class="Expl">Para cancelar as senhas de autorização concedidas para este cliente, clique no botão "Remover" (não é necessário preencher nenhum campo).
	<br>A autorização para desconto superior permanece válida apenas pelos <%=Cstr(TIMEOUT_DESCONTO_EM_MIN)%> minutos seguintes ao seu cadastramento.</p>
	</td>
</tr>
</table>
<br>

<!--  CLIENTE  -->
<table class="Qx" cellSpacing="0" style="width:450px;">
	<tr bgColor="#FFFFFF">
		<%	s = cnpj_cpf_formata(r_cliente.cnpj_cpf)
			if s="" then s="&nbsp;"
		%>
		<td class="MT"><span class="PLTe">CLIENTE</span>
		<br><p class="C"><%=s%></p>
		</td>
	</tr>
	<tr bgColor="#FFFFFF">
		<%	s = iniciais_em_maiusculas(r_cliente.nome)
			if s="" then s="&nbsp;"
		%>
		<td class="MDBE"><span class="PLTe">NOME</span>
		<br><p class="C"><%=s%></p>
		</td>
	</tr>
	<tr bgColor="#FFFFFF">
		<%	s = ""
			with r_cliente
				if .endereco <> "" then
					s = iniciais_em_maiusculas(.endereco)
					s_aux=Trim(.endereco_numero)
					if s_aux<>"" then s=s & ", " & s_aux
					s_aux=Trim(.endereco_complemento)
					if s_aux<>"" then s=s & " " & s_aux
					s_aux=iniciais_em_maiusculas(.bairro)
					if s_aux<>"" then s=s & " - " & s_aux
					s_aux=iniciais_em_maiusculas(.cidade)
					if s_aux<>"" then s=s & " - " & s_aux
					s_aux=.uf
					if s_aux<>"" then s=s & " - " & s_aux
					s_aux=cep_formata(.cep)
					if s_aux<>"" then s=s & " - " & s_aux
					end if
				end with
			if s="" then s="&nbsp;"
		%>
		<td class="MDBE"><span class="PLTe">ENDEREÇO</span>
		<p class="C"><%=s%></p>
		</td>
	</tr>
	<tr bgColor="#FFFFFF">
		<%	s = r_cliente.obs_crediticias
			if s="" then s="&nbsp;"
		%>
		<td class="MDBE"><span class="PLTe">OBSERVAÇÕES CREDITÍCIAS</span>
		<br><p class="C" style="color:red;"><%=s%></p>
		</td>
	</tr>
<!--  PULA LINHA  -->
	<tr bgColor="#FFFFFF">
		<td>&nbsp;</td>
	</tr>
<!--  AUTORIZADO POR  -->
	<tr bgColor="#FFFFFF">
		<td class="MT"><span class="PLTe">AUTORIZADO POR</span>
		<br>
		<select id="c_supervisor_autorizador" name="c_supervisor_autorizador" style="width:490px;" onchange="fFILTRO.c_loja.focus();">
		  <% =autorizador_desconto_monta_itens_select("") %>
		</select>
		</td>
	</tr>
<!--  LOJA  -->
	<tr bgColor="#FFFFFF">
		<td class="MDBE"><span class="PLTe">LOJA</span>
		<br>
			<input maxlength="3" class="PLLe" style="width:250px;" name="c_loja" id="c_loja" onblur="this.value=normaliza_codigo(this.value,TAM_MIN_LOJA);" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fFILTRO.c_fabricante[0].focus(); filtra_numerico();">
		</td>
	</tr>
</table>


<!--  PULA LINHA  -->
<br><br>

<table class="Qx" cellSpacing="0">
	<tr bgColor="#FFFFFF">
	<td class="MB" valign="bottom"><p class="PLTe">Fabr</p></td>
	<td class="MB" valign="bottom"><p class="PLTe">Produto</p></td>
	<td class="MB" valign="bottom"><p class="PLTd">Desc Máx<br>Autorizado (%)</p></td>
	</tr>
<%
	for intCounter=1 to max_qtde_itens
%>

	<tr>
	<!--  FABRICANTE  -->
	<td class="MDBE" NOWRAP>
		<input maxlength="4" class="PLLe" style="width:50px;" name="c_fabricante" id="c_fabricante" onblur="this.value=normaliza_codigo(this.value,TAM_MIN_FABRICANTE);" onkeypress="if (digitou_enter(true)&&(tem_info(this.value)||(<%=Cstr(intCounter)%>!=1))) if (trim(this.value)=='') bCONFIRMA.focus(); else fFILTRO.c_produto[<%=Cstr(intCounter-1)%>].focus(); filtra_fabricante();">
		</td>

	<!--  PRODUTO  -->
	<td class="MDB" NOWRAP>
		<input maxlength="8" class="PLLe" style="width:100px;" name="c_produto" id="c_produto" onblur="this.value=ucase(normaliza_codigo(this.value,TAM_MIN_PRODUTO));" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) fFILTRO.c_desc_max_senha[<%=Cstr(intCounter-1)%>].focus(); filtra_produto();">
		</td>

	<!--  DESCONTO MÁXIMO  -->
	<td class="MDB" NOWRAP>
		<input maxlength="5" class="PLLd" style="color:green;font-size:11pt;width:100px;" name="c_desc_max_senha" id="c_desc_max_senha" onblur="this.value=formata_perc_desc(this.value); if ((converte_numero(this.value)>100)||(converte_numero(this.value)<0)) {alert('Percentual inválido!!');this.focus();}" onkeypress="if (digitou_enter(true)&&tem_info(this.value)) {if (<%=Cstr(intCounter)%>==fFILTRO.c_desc_max_senha.length) bCONFIRMA.focus(); else fFILTRO.c_fabricante[<%=Cstr(intCounter)%>].focus();} filtra_percentual();">
		</td>
	</tr>
<% next %>
</table>


<!-- ************   SEPARADOR   ************ -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>
<br>


<table class="notPrint" width="649" cellSpacing="0">
<tr>
	<td><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()" title="volta para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
	<td align="center"><div name="dREMOVE" id="dREMOVE"><a name="bREMOVE" id="bREMOVE" href="javascript:fCANCConfirma(fCANC)" title="cancela as senhas de autorização para desconto superior concedidas para este cliente">
		<img src="../botao/remover.gif" width="176" height="55" border="0"></a></div>
	</td>
	<td align="right"><div name="dCONFIRMA" id="dCONFIRMA"><a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fFILTROConfirma(fFILTRO)" title="vai para a página de confirmação">
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
