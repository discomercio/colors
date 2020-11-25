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
'	  P E D I D O O C O R R E N C I A N O V A . A S P
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

	dim usuario, pedido_selecionado
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO)

	pedido_selecionado = Trim(request("pedido_selecionado"))
	if (pedido_selecionado = "") then Response.Redirect("aviso.asp?id=" & ERR_PEDIDO_NAO_ESPECIFICADO)

    dim url_origem
    url_origem = Trim(Request("url_origem"))

    dim c_id_cliente
    c_id_cliente = Trim(Request("c_id_cliente"))
	
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

	if Not operacao_permitida(OP_CEN_OCORRENCIAS_EM_PEDIDOS_CADASTRAMENTO, s_lista_operacoes_permitidas) then
		Response.Redirect("aviso.asp?id=" & ERR_NIVEL_ACESSO_INSUFICIENTE)
		end if

	dim cn
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

    dim ddd_1, ddd_2, tel_1, tel_2, contato
    ddd_1 = ""
    ddd_2 = ""
    tel_1 = ""
    tel_2 = ""
    contato = ""

	dim blnUsarMemorizacaoCompletaEnderecos
	blnUsarMemorizacaoCompletaEnderecos = isActivatedFlagPedidoUsarMemorizacaoCompletaEnderecos

' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________

' ____________________________________
' OBTEM_DDD_TELEFONE_CLIENTE
function obtem_ddd_telefone_cliente(ByRef ddd_1, ByRef ddd_2, ByRef tel_1, ByRef tel_2, ByRef contato)
dim s, r
dim blnTelefoneLidoDoPedido, r_pedido, msg_erro

    'vamos ler os dados do pedido?
    blnTelefoneLidoDoPedido = false
    if blnUsarMemorizacaoCompletaEnderecos then
    	if Not le_pedido(pedido_selecionado, r_pedido, msg_erro) then exit function
        if r_pedido.st_memorizacao_completa_enderecos <> 0 then
            blnTelefoneLidoDoPedido = true
            contato = r_pedido.endereco_nome

            if r_pedido.endereco_ddd_cel<>"" then
                ddd_1 = r_pedido.endereco_ddd_cel
                tel_1 = r_pedido.endereco_tel_cel
            end if
            if r_pedido.endereco_ddd_res<>"" then
                if ddd_1 = "" then
                    ddd_1 = r_pedido.endereco_ddd_res
                    tel_1 = r_pedido.endereco_tel_res
                else
                    ddd_2 = r_pedido.endereco_ddd_res
                    tel_2 = r_pedido.endereco_tel_res
                end if
            end if
            if r_pedido.endereco_ddd_com<>"" And ddd_2 = "" then
                if ddd_1 = "" then
                    ddd_1 = r_pedido.endereco_ddd_com
                    tel_1 = r_pedido.endereco_tel_com
                else
                    ddd_2 = r_pedido.endereco_ddd_com
                    tel_2 = r_pedido.endereco_tel_com
                end if
            end if
            if r_pedido.endereco_ddd_com_2<>"" And ddd_2 = "" then
                if ddd_1 = "" then
                    ddd_1 = r_pedido.endereco_ddd_com_2
                    tel_1 = r_pedido.endereco_tel_com_2
                else
                    ddd_2 = r_pedido.endereco_ddd_com_2
                    tel_2 = r_pedido.endereco_tel_com_2
                end if
            end if

        end if
    end if

    'formato de pedido antigo, vamos ler do cliente
    if not blnTelefoneLidoDoPedido then
        s = "SELECT * FROM t_CLIENTE WHERE (id = '" & c_id_cliente & "')"

        set r = cn.Execute(s)

        contato = Trim("" & r("nome"))

        if Trim("" & r("ddd_cel"))<>"" then
            ddd_1 = r("ddd_cel")
            tel_1 = r("tel_cel")
        end if
        if Trim("" & r("ddd_res"))<>"" then
            if ddd_1 = "" then
                ddd_1 = r("ddd_res")
                tel_1 = r("tel_res")
            else
                ddd_2 = r("ddd_res")
                tel_2 = r("tel_res")
            end if
        end if
        if Trim("" & r("ddd_com"))<>"" And ddd_2 = "" then
            if ddd_1 = "" then
                ddd_1 = r("ddd_com")
                tel_1 = r("tel_com")
            else
                ddd_2 = r("ddd_com")
                tel_2 = r("tel_com")
            end if
        end if
        if Trim("" & r("ddd_com_2"))<>"" And ddd_2 = "" then
            if ddd_1 = "" then
                ddd_1 = r("ddd_com_2")
                tel_1 = r("tel_com_2")
            else
                ddd_2 = r("ddd_com_2")
                tel_2 = r("tel_com_2")
            end if
        end if

        r.Close
	    set r=nothing
    end if

end function

' ____________________________________
' MOTIVO_OCORRENCIA_MONTA_ITENS_SELECT

function motivo_ocorrencia_monta_itens_select()
dim x, r, strResp, strSql
    strSql = "SELECT * FROM t_CODIGO_DESCRICAO" & _
                " WHERE grupo='" & GRUPO_T_CODIGO_DESCRICAO__OCORRENCIAS_EM_PEDIDOS__MOTIVO_ABERTURA & "' AND st_inativo=0" & _
                " ORDER BY ordenacao"

    set r = cn.Execute(strSql)
	strResp = "<option value='' selected>&nbsp;</option>"
	do while Not r.EOF 
        x = r("codigo")
        strResp = strResp & "<option"
	    strResp = strResp & " value='" & x & "'>"
        strResp = strResp & r("descricao")
        strResp = strResp & "</option>"
		r.MoveNext        
    loop

    motivo_ocorrencia_monta_itens_select = strResp
	r.Close
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



<% if c_id_cliente <> "" then obtem_ddd_telefone_cliente ddd_1, ddd_2, tel_1, tel_2, contato %>
<html>


<head>
	<title>CENTRAL<%=MontaNumPedidoExibicaoTitleBrowser(pedido_selecionado)%></title>
	</head>



<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>
<script src="<%=URL_FILE__AJAX_JS%>" Language="JavaScript" Type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>

<script type="text/javascript">
    function autoCompletaTelefone(f) {
        f.c_contato.value = "<%=contato%>";
        f.c_ddd_1.value = "<%=ddd_1%>";
        f.c_tel_1.value = "<%=tel_1%>";
        f.c_ddd_2.value = "<%=ddd_2%>";
        f.c_tel_2.value = "<%=tel_2%>";
    }
</script>

<script language="JavaScript" type="text/javascript">
    $(function () {
        $("#table_descricao").hide();
    })

    function calcula_tamanho_restante() {
        var f, s;
        f = fPED;
        s = "" + fPED.c_texto.value;
        f.c_tamanho_restante.value = MAX_TAM_DESCRICAO_OCORRENCIAS_EM_PEDIDOS - s.length;
    }

    function fPEDOcorrenciaNovaConfirma(f) {
        var s, blnFlag;

        //  Contato
        if (trim(f.c_contato.value) == "") {
            alert("Informe o nome da pessoa para contato!!");
            f.c_contato.focus();
            return;
        }

        //  Telefone 1
        blnFlag = false;
        if (trim(f.c_ddd_1.value) != "") blnFlag = true;
        if (trim(f.c_tel_1.value) != "") blnFlag = true;
        if (blnFlag) {
            if (trim(f.c_ddd_1.value) == "") {
                alert("Informe o DDD do telefone!!");
                f.c_ddd_1.focus();
                return;
            }
            if (trim(f.c_tel_1.value) == "") {
                alert("Informe o número do telefone!!");
                f.c_tel_1.focus();
                return;
            }
        }

        //  Telefone 2
        blnFlag = false;
        if (trim(f.c_ddd_2.value) != "") blnFlag = true;
        if (trim(f.c_tel_2.value) != "") blnFlag = true;
        if (blnFlag) {
            if (trim(f.c_ddd_2.value) == "") {
                alert("Informe o DDD do telefone!!");
                f.c_ddd_2.focus();
                return;
            }
            if (trim(f.c_tel_2.value) == "") {
                alert("Informe o número do telefone!!");
                f.c_tel_2.focus();
                return;
            }
        }

        //  Foi informado pelo menos 1 telefone?
        if ((trim(f.c_tel_1.value) == "") && (trim(f.c_tel_2.value) == "")) {
            alert("É necessário informar pelo menos um número de telefone para contato!!");
            f.c_ddd_1.focus();
            return;
        }

        //  Texto descrevendo a ocorrência
        s = "" + f.c_texto.value;
        if (f.motivo_ocorrencia.value == 001) {
            if (s.length == 0) {
                alert('É necessário escrever o texto descrevendo a ocorrência!!');
                f.c_texto.focus();
                return;
            }
        }
        if (s.length > MAX_TAM_DESCRICAO_OCORRENCIAS_EM_PEDIDOS) {
            alert('Conteúdo do texto excede em ' + (s.length - MAX_TAM_DESCRICAO_OCORRENCIAS_EM_PEDIDOS) + ' caracteres o tamanho máximo de ' + MAX_TAM_DESCRICAO_OCORRENCIAS_EM_PEDIDOS + '!!');
            f.c_texto.focus();
            return;
        }

        dCONFIRMA.style.visibility = "hidden";
        window.status = "Aguarde ...";
        f.submit();
    }
</script>

<script type="text/javascript">
    function selecionaMotivoOcorrencia(codigo) {
        if (codigo == 001) {
            $("#table_descricao").show();
            return;
        }
        else {
            $("#table_descricao").hide();
            $("#c_texto").empty();
            return;
        }
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

<body onload="fPED.c_contato.focus();">
<center>

<form id="fPED" name="fPED" method="post" action="PedidoOcorrenciaNovaConfirma.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="pedido_selecionado" id="pedido_selecionado" value='<%=pedido_selecionado%>'>
<input type="hidden" name="url_origem" id="url_origem" value="<%=url_origem%>" />

<!--  I D E N T I F I C A Ç Ã O   D O   P E D I D O -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">

<tr>
	<td align="left" valign="bottom"><p class="PEDIDO">Ocorrência</p></td>
	<td align="right" valign="bottom"><p class="PEDIDO" style="font-size:14pt;">Pedido <%=pedido_selecionado%></p></td>
</tr>
</table>
<br>

<table>
<tr>
	<td>
	<table class="Q" style="width:649px;" cellSpacing="0">
		<tr>
			<td colspan="2">
				<p class="Rf">CONTATO</p>
				<p class="C">
				<input name="c_contato" id="c_contato" class="TA" maxlength="60"  style="width:280px;" value="" onkeypress="if (digitou_enter(true)) fPED.c_ddd_1.focus(); filtra_nome_identificador();">
				<a href="javascript:autoCompletaTelefone(fPED);">
				<img src="../IMAGEM/copia_20x20.png" title="Clique para preencher os telefones conforme o contido no cadastro do cliente" /></a>
                </p>
			</td>
		</tr>
		<tr>
			<td class="MC MD" width="20%">
				<p class="Rf">DDD</p>
				<p class="C">
				<input name="c_ddd_1" id="c_ddd_1" class="TA" maxlength="2" size="4" value="" onkeypress="if (digitou_enter(true)) fPED.c_tel_1.focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inválido!!');this.focus();}">
				</p>
			</td>
			<td class="MC">
				<p class="R">TELEFONE</p>
				<p class="C">
				<input name="c_tel_1" id="c_tel_1" class="TA" maxlength="9" value="" onkeypress="if (digitou_enter(true)) fPED.c_ddd_2.focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('Telefone inválido!!');this.focus();} else this.value=telefone_formata(this.value);">
				</p>
			</td>
		</tr>
		<tr>
			<td class="MC MD" width="20%">
				<p class="Rf">DDD</p>
				<p class="C">
				<input name="c_ddd_2" id="c_ddd_2" class="TA" maxlength="2" size="4" value="" onkeypress="if (digitou_enter(true)) fPED.c_tel_2.focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inválido!!');this.focus();}">
				</p>
			</td>
			<td class="MC">
				<p class="R">TELEFONE</p>
				<p class="C">
				<input name="c_tel_2" id="c_tel_2" class="TA" maxlength="9" value="" onkeypress="if (digitou_enter(true)) fPED.c_texto.focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('Telefone inválido!!');this.focus();} else this.value=telefone_formata(this.value);">
				</p>
			</td>
		</tr>
	</table>
	</td>
</tr>
<tr>
	<td>&nbsp;</td>
</tr>
<tr>
	<td align="left" valign="bottom" style="border: 1pt solid #C0C0C0;">
        <p class="Rf">MOTIVO DA ABERTURA DA OCORRÊNCIA</p>
        <p class="C">
            <select name="motivo_ocorrencia" onchange="selecionaMotivoOcorrencia(this.value);">
                <%=motivo_ocorrencia_monta_itens_select%>
            </select>
        </p>  
	</td>
</tr>
<tr>
	<td>
	    <table id="table_descricao" class="Q" style="width:649px;" cellSpacing="0">
		    <tr>
			    <td>
                    <p class="Rf">DESCRIÇÃO DA OCORRÊNCIA</p>

			    </td>
                <td align="right" valign="bottom">
		            <span class="PLLd">Tamanho restante:</span><input name="c_tamanho_restante" id="c_tamanho_restante" tabindex=-1 readonly class="TA" style="width:35px;text-align:right;" value='<%=Cstr(MAX_TAM_DESCRICAO_OCORRENCIAS_EM_PEDIDOS)%>' />
	            </td>
            </tr>
            <tr>
                <td colspan="2">
				    <textarea name="c_texto" id="c_texto" class="PLLe" rows="<%=Cstr(MAX_LINHAS_DESCRICAO_OCORRENCIAS_EM_PEDIDOS)%>" 
					    style="width:642px;margin-left:2pt;" onkeypress="limita_tamanho(this,MAX_TAM_DESCRICAO_OCORRENCIAS_EM_PEDIDOS);" onblur="this.value=trim(this.value);calcula_tamanho_restante();"
					    onkeyup="calcula_tamanho_restante();"
					    ></textarea>
			    </td>
		    </tr>
	    </table>
	</td>
</tr>
</table>


<!-- ************   SEPARADOR   ************ -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>
<br>


<table class="notPrint" width="649" cellSpacing="0">
<tr>
	<td><a name="bCANCELA" id="bCANCELA" href="javascript:history.back()" title="cancela o cadastramento da nova ocorrência">
		<img src="../botao/cancelar.gif" width="176" height="55" border="0"></a></td>
	<td align="right"><div name="dCONFIRMA" id="dCONFIRMA">
		<a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fPEDOcorrenciaNovaConfirma(fPED)" title="grava a nova ocorrência">
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