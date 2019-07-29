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
'	  P E D I D O C H A M A D O N O V O . A S P
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

	dim usuario, loja, pedido_selecionado
	usuario = Trim(Session("usuario_atual"))
	loja = Trim(Session("loja_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 
	If (loja = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

	pedido_selecionado = Trim(request("pedido_selecionado"))
	if (pedido_selecionado = "") then Response.Redirect("aviso.asp?id=" & ERR_PEDIDO_NAO_ESPECIFICADO)

    dim url_origem
    url_origem = Trim(Request("url_origem"))

    dim c_id_cliente
    c_id_cliente = Trim(Request("c_id_cliente"))
	
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

	if Not operacao_permitida(OP_LJA_PEDIDO_CHAMADO_CADASTRAMENTO, s_lista_operacoes_permitidas) then
		Response.Redirect("aviso.asp?id=" & ERR_NIVEL_ACESSO_INSUFICIENTE)
		end if

	dim cn
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

    dim nivel_acesso_chamado
	nivel_acesso_chamado = Session("nivel_acesso_chamado")
	if Trim(nivel_acesso_chamado) = "" then
		nivel_acesso_chamado = obtem_nivel_acesso_chamado_pedido(cn, usuario)
		Session("nivel_acesso_chamado") = nivel_acesso_chamado
		end if

    dim ddd_1, ddd_2, tel_1, tel_2, contato
    ddd_1 = ""
    ddd_2 = ""
    tel_1 = ""
    tel_2 = ""

' _____________________________________________________________________________________________
'
'									F  U  N  �  �  E  S 
' _____________________________________________________________________________________________

' ____________________________________
' OBTEM_DDD_TELEFONE_CLIENTE
function obtem_ddd_telefone_cliente(ByRef ddd_1, ByRef ddd_2, ByRef tel_1, ByRef tel_2, ByRef contato)
dim s, r
    
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

end function

' ____________________________________
' DEPTO_PEDIDO_CHAMADO_MONTA_ITENS_SELECT

function depto_pedido_chamado_monta_itens_select()
dim x, r, strResp, strSql
    strSql = "SELECT * FROM t_PEDIDO_CHAMADO_DEPTO" & _
                " WHERE st_inativo=0" & _
                " ORDER BY descricao"

    set r = cn.Execute(strSql)
	strResp = "<option value='' selected>&nbsp;</option>"
	do while Not r.EOF 
        x = r("id")
        strResp = strResp & "<option"
	    strResp = strResp & " value='" & x & "'>"
        strResp = strResp & r("descricao")
        strResp = strResp & "</option>"
		r.MoveNext        
    loop

    depto_pedido_chamado_monta_itens_select = strResp
	r.Close
	set r=nothing
end function

' _____________________________________
' LISTA MONTA
'
sub lista_monta
dim r, x, s_sql, fabricante_a, nome_fabricante, n_reg, msg_erro

	x = "<script language='JavaScript'>" & chr(13) & _
		"Pd = new Array();" & chr(13) & _
		"Pd[0] = new oPd('','','');" & chr(13)

	s_sql = "SELECT * FROM t_CODIGO_DESCRICAO WHERE grupo='" & GRUPO_T_CODIGO_DESCRICAO__CHAMADOS_EM_PEDIDOS__MOTIVO_ABERTURA & "'"
	
	n_reg = 0

	if Not cria_recordset_otimista(r, msg_erro) then
		Response.Write msg_erro
		exit sub
		end if
	
	r.Open s_sql, cn
	do while Not r.Eof
		n_reg = n_reg + 1
	
	 '> MONTA LINHA
		x = x & "Pd[Pd.length]=new oPd('" & Trim("" & r("codigo")) & "'" & _
				",'" & Trim("" & r("descricao")) & "'" & _
				",'" & Trim("" & r("codigo_pai")) & "'" & _
				");" & chr(13)

		r.movenext
		loop

	x = x & "</script>" & chr(13)
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


<% obtem_ddd_telefone_cliente ddd_1, ddd_2, tel_1, tel_2, contato %>
<html>


<head>
	<title>LOJA<%=MontaNumPedidoExibicaoTitleBrowser(pedido_selecionado)%></title>
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
    
    function oPd(codigo, descricao, codigo_pai) {
        this.codigo = codigo;
        this.descricao = descricao;
        this.codigo_pai = codigo_pai;
    }

    function calcula_tamanho_restante() {
        var f, s;
        f = fPED;
        s = "" + fPED.c_texto.value;
        f.c_tamanho_restante.value = MAX_TAM_DESCRICAO_CHAMADO_EM_PEDIDOS - s.length;
    }

    function fPEDChamadoNovoConfirma(f) {
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
                alert("Informe o n�mero do telefone!!");
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
                alert("Informe o n�mero do telefone!!");
                f.c_tel_2.focus();
                return;
            }
        }

        //  Foi informado pelo menos 1 telefone?
        if ((trim(f.c_tel_1.value) == "") && (trim(f.c_tel_2.value) == "")) {
            alert("� necess�rio informar pelo menos um n�mero de telefone para contato!!");
            f.c_ddd_1.focus();
            return;
        }

        //  Texto descrevendo o chamado
        s = "" + f.c_texto.value;
            if (s.length == 0) {
                alert('� necess�rio escrever o texto descrevendo o chamado!!');
                f.c_texto.focus();
                return;
            }

        if (s.length > MAX_TAM_DESCRICAO_CHAMADO_EM_PEDIDOS) {
            alert('Conte�do do texto excede em ' + (s.length - MAX_TAM_DESCRICAO_CHAMADO_EM_PEDIDOS) + ' caracteres o tamanho m�ximo de ' + MAX_TAM_DESCRICAO_CHAMADO_EM_PEDIDOS + '!!');
            f.c_texto.focus();
            return;
        }

        dCONFIRMA.style.visibility = "hidden";
        window.status = "Aguarde ...";
        f.submit();
    }
</script>

<% lista_monta %>

<script type="text/javascript">
    function montaDropdownListMotivosAbertura(depto) {
        var i;
        var select = document.getElementById("motivo_chamado");
        select.innerHTML = "";
        select.options[select.options.length] = new Option("", "");
        i = 0;

        for (i = 0; i < Pd.length; i++) {
            if (Pd[i].codigo_pai == depto) {
                select.options[select.options.length] = new Option(Pd[i].descricao, Pd[i].codigo);
            }
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

<form id="fPED" name="fPED" method="post" action="PedidoChamadoNovoConfirma.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="pedido_selecionado" id="pedido_selecionado" value='<%=pedido_selecionado%>'>
<input type="hidden" name="url_origem" id="url_origem" value="<%=url_origem%>" />

<!--  I D E N T I F I C A � � O   D O   P E D I D O -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">

<tr>
	<td align="left" valign="bottom"><p class="PEDIDO">Abertura de Chamado</p></td>
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
				<input name="c_contato" id="c_contato" class="TA" maxlength="40"  style="width:280px;" value="" onkeypress="if (digitou_enter(true)) fPED.c_ddd_1.focus(); filtra_nome_identificador();">
                <a href="javascript:autoCompletaTelefone(fPED);">
				<img src="../IMAGEM/copia_20x20.png" title="Clique para preencher os telefones conforme o contido no cadastro do cliente" /></a>
				</p>
			</td>
		</tr>
		<tr>
			<td class="MC MD" width="20%">
				<p class="Rf">DDD</p>
				<p class="C">
				<input name="c_ddd_1" id="c_ddd_1" class="TA" maxlength="2" size="4" value="" onkeypress="if (digitou_enter(true)) fPED.c_tel_1.focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inv�lido!!');this.focus();}">
				</p>
			</td>
			<td class="MC">
				<p class="R">TELEFONE</p>
				<p class="C">
				<input name="c_tel_1" id="c_tel_1" class="TA" maxlength="9" value="" onkeypress="if (digitou_enter(true)) fPED.c_ddd_2.focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('Telefone inv�lido!!');this.focus();} else this.value=telefone_formata(this.value);">				
                </p>
			</td>
		</tr>
		<tr>
			<td class="MC MD" width="20%">
				<p class="Rf">DDD</p>
				<p class="C">
				<input name="c_ddd_2" id="c_ddd_2" class="TA" maxlength="2" size="4" value="" onkeypress="if (digitou_enter(true)) fPED.c_tel_2.focus(); filtra_numerico();" onblur="if (!ddd_ok(this.value)) {alert('DDD inv�lido!!');this.focus();}">
				</p>
			</td>
			<td class="MC">
				<p class="R">TELEFONE</p>
				<p class="C">
				<input name="c_tel_2" id="c_tel_2" class="TA" maxlength="9" value="" onkeypress="if (digitou_enter(true)) fPED.c_texto.focus(); filtra_numerico();" onblur="if (!telefone_ok(this.value)) {alert('Telefone inv�lido!!');this.focus();} else this.value=telefone_formata(this.value);">
				</p>
			</td>
		</tr>
        <tr>
	    <td align="left" valign="bottom" colspan="2" class="MC">
        <p class="Rf">DEPARTAMENTO RESPONS�VEL</p>
        <p class="C">
            <select name="c_depto" onchange="montaDropdownListMotivosAbertura(this.value)"  style="width: 280px">
                <%=depto_pedido_chamado_monta_itens_select%>
            </select>
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
        <p class="Rf">MOTIVO DA ABERTURA DO CHAMADO</p>
        <p class="C">
            <select name="motivo_chamado" id="motivo_chamado" style="width: 500px">
            </select>
        </p>  
	</td>
</tr>
<tr>
	<td>
	    <table id="table_descricao" class="Q" style="width:649px;margin:0" cellSpacing="0">
		    <tr>
			    <td>
                    <p class="Rf">DESCRI��O DO CHAMADO</p>

			    </td>
                <td align="right" valign="bottom">
		            <span class="PLLd">Tamanho restante:</span><input name="c_tamanho_restante" id="c_tamanho_restante" tabindex=-1 readonly class="TA" style="width:35px;text-align:right;" value='<%=Cstr(MAX_TAM_DESCRICAO_CHAMADO_EM_PEDIDOS)%>' />
	            </td>
            </tr>
            <tr>
                <td colspan="2">
				    <textarea name="c_texto" id="c_texto" class="PLLe" rows="<%=Cstr(MAX_LINHAS_DESCRICAO_CHAMADOS_EM_PEDIDOS)%>" 
					    style="width:642px;margin-left:2pt;" onkeypress="limita_tamanho(this,MAX_TAM_DESCRICAO_CHAMADO_EM_PEDIDOS);" onblur="this.value=trim(this.value);calcula_tamanho_restante();"
					    onkeyup="calcula_tamanho_restante();"
					    ></textarea>
			    </td>
		    </tr>
	    </table>
	</td>
</tr>
    <tr>
	<td align="left" valign="bottom" style="border: 1pt solid #C0C0C0;">
        <p class="Rf">N�VEL DE ACESSO AO CHAMADO</p>
        <p class="C">
            <select name="c_nivel_acesso_chamado" style="width: 200px">
                <% =nivel_acesso_chamado_pedido_monta_itens_select(Null, nivel_acesso_chamado, True) %>
            </select>
        </p>  
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
	<td><a name="bCANCELA" id="bCANCELA" href="javascript:history.back()" title="cancela a abertura do chamado">
		<img src="../botao/cancelar.gif" width="176" height="55" border="0"></a></td>
	<td align="right"><div name="dCONFIRMA" id="dCONFIRMA">
		<a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fPEDChamadoNovoConfirma(fPED)" title="grava a abertura do chamado">
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