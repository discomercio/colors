<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/Global.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =============================================
'	  EstoqueTransfereEntreCDs.asp
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
	if Not operacao_permitida(OP_CEN_TRANSF_MOV_ESTOQUE_PERFIL_AVANCADO, s_lista_operacoes_permitidas) then
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	dim max_qtde_itens
	max_qtde_itens = obtem_parametro_TransfProdutosEntreCDs_MaxQtdeItens

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
<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>
<%  dim strScript
	strScript = _
		"<script language='JavaScript' type='text/javascript'>" & chr(13) & _
		"var MAX_TAM_T_ESTOQUE_CAMPO_OBS = " & Cstr(MAX_TAM_T_ESTOQUE_CAMPO_OBS) & ";" & chr(13) & _
		"</script>" & chr(13)
	Response.Write strScript
%>

<script language="JavaScript" type="text/javascript">
    $(function () {
        $(".TdFab").hide();
        document.getElementById("img_collapse_fabricante").src = document.getElementById("img_collapse_fabricante").src.replace("CollapseLeft_20px.png", "CollapseRight_20px.png");
    });

    function toggleColFabricante() {
        if ($(".TdFab").first().is(":visible")) {
            $(".TdFab").hide();
            document.getElementById("img_collapse_fabricante").src = document.getElementById("img_collapse_fabricante").src.replace("CollapseLeft_20px.png", "CollapseRight_20px.png");
        }
        else {
            $(".TdFab").show();
            document.getElementById("img_collapse_fabricante").src = document.getElementById("img_collapse_fabricante").src.replace("CollapseRight_20px.png", "CollapseLeft_20px.png");
        }
    }

    function calcula_tamanho_restante() {
        var f, s;
        f = fOP;
        s = "" + f.c_obs.value;
        f.c_tamanho_restante.value = MAX_TAM_T_ESTOQUE_CAMPO_OBS - s.length;
    }

    function cancela_onpaste() {
        return false;
    }

    function fOpConfirma(f) {
        var i, b, ha_item;

        if (trim(f.c_nfe_emitente_origem.value) == "") {
            alert("Selecione o CD de origem da transferência!!");
            f.c_nfe_emitente_origem.focus();
            return;
        }

        if (trim(f.c_nfe_emitente_destino.value) == "") {
            alert("Selecione o CD de destino da transferência!!");
            f.c_nfe_emitente_destino.focus();
            return;
        }

        if (trim(f.c_nfe_emitente_origem.value) == trim(f.c_nfe_emitente_destino.value)) {
            alert("O CD de destino da transferência não pode ser o mesmo da origem!!");
            f.c_nfe_emitente_destino.focus();
            return;
        }

        ha_item = false;
        for (i = 0; i < f.c_produto.length; i++) {
            b = false;
            if (trim(f.c_fabricante[i].value) != "") b = true;
            if (trim(f.c_produto[i].value) != "") b = true;
            if (trim(f.c_qtde[i].value) != "") b = true;

            if (b) {
                ha_item = true;
                if (trim(f.c_produto[i].value) == "") {
                    alert("Informe o código do produto!!");
                    f.c_produto[i].focus();
                    return;
                }
                if (trim(f.c_qtde[i].value) == "") {
                    alert("Informe a quantidade!!");
                    f.c_qtde[i].focus();
                    return;
                }
                if (parseInt(f.c_qtde[i].value) <= 0) {
                    alert("Quantidade inválida!!");
                    f.c_qtde[i].select();
                    f.c_qtde[i].focus();
                    return;
                }
            }
        }

        if (!ha_item) {
            alert("Não há produtos na lista!!");
            f.c_produto[0].focus();
            return;
        }

        //s = "" + f.c_obs.value;
        //if (s.length > MAX_TAM_T_ESTOQUE_CAMPO_OBS) {
        //	alert('Conteúdo de "Observações" excede em ' + (s.length-MAX_TAM_T_ESTOQUE_CAMPO_OBS) + ' caracteres o tamanho máximo de ' + MAX_TAM_T_ESTOQUE_CAMPO_OBS + '!!');
        //	f.c_obs.focus();
        //	return;
        //	}

        dCONFIRMA.style.visibility = "hidden";
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

<link href="<%=URL_FILE__E_CSS%>" rel="stylesheet" type="text/css">


<body>
<center>

<form id="fOP" name="fOP" method="post" action="EstoqueTransfereEntreCDsConsiste.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Transferência de Produtos Entre CD's</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>
<br>

<!--  INFORMAÇÕES SOBRE A TRANSFERÊNCIA E PRODUTOS A SEREM TRANSFERIDOS  -->
<table class="Qx" cellspacing="0" cellpadding="0">
	<!--  TÍTULO  -->
	<tr bgcolor="#FFFFFF">
	<td><span style="width:30px;">&nbsp;</span></td>
	<td class="MT" valign="middle" align="center" nowrap style="background:azure;padding:5px;"><span class="PLTc" style="vertical-align:middle;"
		>TRANSFERÊNCIA</span></td>
	</tr>
	<!--  CD DE ORIGEM  -->
	<tr bgcolor="#FFFFFF">
	<td>&nbsp;</td>
	<td class="MDBE">
		<table cellspacing="0" cellpadding="0" border="0" width="100%">
			<tr>
				<td class="MD" width="50%" style="padding:3px;">
					<span class="PLTe">CD ORIGEM</span>
					<br />
					<select name="c_nfe_emitente_origem" id="c_nfe_emitente_origem" class="C" style="margin:4px 8px 10px 8px;min-width:100px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
					<% =wms_apelido_empresa_nfe_emitente_monta_itens_select(Null) %>
					</select>
				</td>
				<td width="50%">
					<span class="PLTe" style="padding:3px;">CD DESTINO</span>
					<br />
					<select name="c_nfe_emitente_destino" id="c_nfe_emitente_destino" class="C" style="margin:4px 8px 10px 8px;min-width:100px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
					<% =wms_apelido_empresa_nfe_emitente_monta_itens_select(Null) %>
					</select>
				</td>
			</tr>
		</table>
	</td>
	</tr>
	<!--  OBS  -->
<!--	<tr bgcolor="#FFFFFF">
	<td>&nbsp;</td>
	<td class="MDBE" style="padding-top:3px;">
		<table cellspacing="0" cellpadding="0" border="0" width="100%">
			<tr>
				<td width="50%">
					<span class="PLTe">Observações</span>
				</td>
				<td width="50%" align="right">
					<span class="PLLd">Tamanho restante:</span><input name="c_tamanho_restante" id="c_tamanho_restante" tabindex=-1 readonly class="TA" style="width:35px;text-align:right;" value='<%=Cstr(MAX_TAM_T_ESTOQUE_CAMPO_OBS)%>' />
				</td>
			</tr>
			<tr>
				<td colspan="2">
					<textarea name="c_obs" id="c_obs" class="PLLe" rows="<%=Cstr(MAX_LINHAS_ESTOQUE_OBS)%>"
							style="width:642px;margin-left:2pt;" onkeypress="limita_tamanho(this,MAX_TAM_T_ESTOQUE_CAMPO_OBS);" onblur="this.value=trim(this.value);calcula_tamanho_restante();"
							onkeyup="calcula_tamanho_restante();"
					></textarea>
				</td>
			</tr>
		</table>
	</td>
	</tr>-->
</table>
<br /><br />

<!--  R E L A Ç Ã O   D E   P R O D U T O S  -->
<table class="Qx" cellspacing="0" cellpadding="2">
	<tr bgColor="#FFFFFF">
	<td align="right" style="font-weight:bold;cursor:pointer;"><a href="javascript:toggleColFabricante();"><img id="img_collapse_fabricante" src="../IMAGEM/CollapseRight_20px.png" /></a></td>
	<td class="MB TdFab" align="left"><span class="PLTe">Fabricante</span></td>
	<td class="MB" align="left"><span class="PLTe">Produto</span></td>
	<td class="MB" align="left"><span class="PLTd">Qtde</span></td>
	</tr>
<% for i=1 to max_qtde_itens %>
	<tr>
	<td class="MD" align="left">
		<input name="c_linha" id="c_linha" readonly tabindex=-1 class="PLLe" maxlength="2" style="width:30px;text-align:right;color:#808080;" 
			value="<%=Cstr(i) & ". " %>"></td>
	<td class="MDB TdFab" align="left">
		<input name="c_fabricante" id="c_fabricante_<%=Cstr(i)%>" class="PLLe" maxlength="4" style="width:60px;" 
			onkeypress="if (digitou_enter(true)) {fOP.c_produto[<%=Cstr(i-1)%>].focus();} filtra_fabricante();"
			onblur="this.value=normaliza_fabricante(this.value);"></td>
	<td class="MDB" align="left">
		<input name="c_produto" id="c_produto_<%=Cstr(i)%>" class="PLLe" maxlength="13" style="width:100px;" 
			onkeypress="if (digitou_enter(true)&&(tem_info(this.value)||(<%=Cstr(i)%>!=1))) if (trim(this.value)=='') bCONFIRMA.focus(); else fOP.c_qtde[<%=Cstr(i-1)%>].focus(); filtra_produto();" 
			onblur="this.value=normaliza_produto(this.value);"></td>
	<td class="MDB" align="right">
		<input name="c_qtde" id="c_qtde_<%=Cstr(i)%>" class="PLLd" maxlength="4" style="width:35px;" 
			onkeypress="if (digitou_enter(true)) {if (<%=Cstr(i)%>==fOP.c_qtde.length) {bCONFIRMA.focus();} else {if ($('#c_fabricante_<%=Cstr(i+1)%>').is(':visible')) fOP.c_fabricante[<%=Cstr(i)%>].focus(); else fOP.c_produto[<%=Cstr(i)%>].focus();}} filtra_numerico();"></td>
	</tr>
<% next %>
</table>


<!-- ************   SEPARADOR   ************ -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>
<br>


<table width="649" cellSpacing="0">
<tr>
	<td><a name="bCANCELA" id="bCANCELA" href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="cancela a operação">
		<img src="../botao/cancelar.gif" width="176" height="55" border="0"></a></td>
	<td align="right"><div name="dCONFIRMA" id="dCONFIRMA"><a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fOpConfirma(fOP)" title="vai para a página de confirmação">
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