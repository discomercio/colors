<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp"        -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ========================================================
'	  RelFarolResumidoFiltro.asp
'     ========================================================
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

	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	if Not operacao_permitida(OP_CEN_REL_FAROL_RESUMIDO, s_lista_operacoes_permitidas) then
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if
	
'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	
	dim strDtHojeYYYYMMDD, strDtHojeDDMMYYYY
	strDtHojeYYYYMMDD = formata_data_yyyymmdd(Date)
	strDtHojeDDMMYYYY = formata_data(Date)





' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________

' ____________________________________________________________________________
' FABRICANTE MONTA ITENS SELECT
'
function fabricante_monta_itens_select(byval id_default)
dim x, r, strResp, ha_default, strSql, i
dim v
	id_default = Trim("" & id_default)
	v = split(id_default, ", ")
	ha_default=False
	strSql = "SELECT DISTINCT" & _
				" Coalesce(fabricante,'') AS fabricante" & _
			" FROM t_PRODUTO" & _
			" WHERE" & _
				" (Coalesce(fabricante,'') <> '')" & _
			" ORDER BY" & _
				" Coalesce(fabricante,'')"
	set r = cn.Execute(strSql)
	strResp = ""
  
	do while Not r.eof 
	    
		x = Trim("" & r("fabricante"))
		strResp = strResp & "<option "
            for i=LBound(v) to UBound(v) 
		        if (id_default<>"") And (v(i)=x) then
		            strResp = strResp & "selected"
		         end if
		   	 next
		strResp = strResp & " value='" & x & "'>"
		strResp = strResp & Trim("" & r("fabricante")) & "&nbsp;&nbsp;"
		strResp = strResp & "</option>" & chr(13)
		r.MoveNext	
 	loop

	fabricante_monta_itens_select = strResp
	r.close
	set r=nothing
end function



' ____________________________________________________________________________
' GRUPO MONTA ITENS SELECT
'
function grupo_monta_itens_select(byval id_default)
dim x, r, strResp, ha_default, strSql, v, i, sDescricao
	id_default = Trim("" & id_default)
	v = split(id_default, ", ")
	ha_default=False
	strSql = "SELECT DISTINCT" & _
				" tP.grupo," & _
				" tPG.descricao" & _
			" FROM t_PRODUTO tP" & _
				" LEFT JOIN t_PRODUTO_GRUPO tPG ON (tP.grupo = tPG.codigo)" & _
			" WHERE" & _
				" (LEN(Coalesce(tP.grupo,'')) > 0)" & _
			" ORDER BY" & _
				" tP.grupo"
	set r = cn.Execute(strSql)
	strResp = ""
	do while Not r.eof 
		x = Trim("" & r("grupo"))
		sDescricao = Trim("" & r("descricao"))
		strResp = strResp & "<option "
		for i=LBound(v) to UBound(v) 
			if (id_default<>"") And (v(i)=x) then
				strResp = strResp & "selected"
				end if
			next
		strResp = strResp & " value='" & x & "'>"
		strResp = strResp & Trim("" & r("grupo"))
		if sDescricao <> "" then strResp = strResp & " &nbsp;(" & sDescricao & ")"
		strResp = strResp & "</option>" & chr(13)
		r.MoveNext	
 	loop
		
	grupo_monta_itens_select = strResp
	r.close
	set r=nothing
end function

'----------------------------------------------------------------------------------------------
' SUBGRUPO MONTA ITENS SELECT
function subgrupo_monta_itens_select(byval id_default)
dim x, r, strSql, strResp, ha_default, v, i, sDescricao
	id_default = Trim("" & id_default)
	v = split(id_default, ", ")
	ha_default=False
	strSql = "SELECT DISTINCT tP.subgrupo, tPS.descricao FROM t_PRODUTO tP LEFT JOIN t_PRODUTO_SUBGRUPO tPS ON (tP.subgrupo = tPS.codigo) WHERE LEN(Coalesce(tP.subgrupo,'')) > 0 ORDER by tP.subgrupo"
	set r = cn.Execute(strSql)
	strResp = ""
	do while Not r.eof 
		x = UCase(Trim("" & r("subgrupo")))
		sDescricao = Trim("" & r("descricao"))
		strResp = strResp & "<option "
		for i=LBound(v) to UBound(v) 
			if (id_default<>"") And (v(i)=x) then
				strResp = strResp & "selected"
				end if
			next
		strResp = strResp & " VALUE='" & x & "'>"
		strResp = strResp & x
		if sDescricao <> "" then strResp = strResp & " &nbsp;(" & sDescricao & ")"
		strResp = strResp & "</OPTION>" & chr(13)
		r.MoveNext
		loop
	
	subgrupo_monta_itens_select = strResp
	r.close
	set r=nothing
end function

' ____________________________________________________________________________
' POTENCIA BTU MONTA ITENS SELECT
'
function potencia_BTU_monta_itens_select(byval id_default)
dim x, r, strResp, ha_default, strSql
	id_default = Trim("" & id_default)
	ha_default=False
	strSql = "SELECT DISTINCT" & _
				" potencia_BTU" & _
			" FROM t_PRODUTO" & _
			" WHERE" & _
				" (potencia_BTU <> 0)" & _
			" ORDER BY" & _
				" potencia_BTU"
	set r = cn.Execute(strSql)
	strResp = ""
	do while Not r.eof 
		x = Trim("" & r("potencia_BTU"))
		if (id_default<>"") And (id_default=x) then
			strResp = strResp & "<option selected"
			ha_default=True
		else
			strResp = strResp & "<option"
			end if
		strResp = strResp & " value='" & x & "'>"
		strResp = strResp & formata_inteiro(r("potencia_BTU")) & "&nbsp;&nbsp;"
		strResp = strResp & "</option>" & chr(13)
		r.MoveNext
		loop

	strResp = "<option selected value=''>&nbsp;</option>" & chr(13) & strResp
		
	potencia_BTU_monta_itens_select = strResp
	r.close
	set r=nothing
end function



' ____________________________________________________________________________
' CICLO MONTA ITENS SELECT
'
function ciclo_monta_itens_select(byval id_default)
dim x, r, strResp, ha_default, strSql
	id_default = Trim("" & id_default)
	ha_default=False
	strSql = "SELECT DISTINCT" & _
				" Coalesce(ciclo,'') AS ciclo" & _
			" FROM t_PRODUTO" & _
			" WHERE" & _
				" (Coalesce(ciclo,'') <> '')" & _
			" ORDER BY" & _
				" Coalesce(ciclo,'')"
	set r = cn.Execute(strSql)
	strResp = ""
	do while Not r.eof 
		x = UCase(Trim("" & r("ciclo")))
		if (id_default<>"") And (id_default=x) then
			strResp = strResp & "<option selected"
			ha_default=True
		else
			strResp = strResp & "<option"
			end if
		strResp = strResp & " value='" & x & "'>"
		strResp = strResp & UCase(Trim("" & r("ciclo"))) & "&nbsp;&nbsp;"
		strResp = strResp & "</option>" & chr(13)
		r.MoveNext
		loop

	strResp = "<option selected value=''>&nbsp;</option>" & chr(13) & strResp
		
	ciclo_monta_itens_select = strResp
	r.close
	set r=nothing
end function



' ____________________________________________________________________________
' POSICAO MERCADO MONTA ITENS SELECT
'
function posicao_mercado_monta_itens_select(byval id_default)
dim x, r, strResp, ha_default, strSql
	id_default = Trim("" & id_default)
	ha_default=False
	strSql = "SELECT DISTINCT" & _
				" Coalesce(posicao_mercado,'') AS posicao_mercado" & _
			" FROM t_PRODUTO" & _
			" WHERE" & _
				" (Coalesce(posicao_mercado,'') <> '')" & _
			" ORDER BY" & _
				" Coalesce(posicao_mercado,'')"
	set r = cn.Execute(strSql)
	strResp = ""
	do while Not r.eof 
		x = UCase(Trim("" & r("posicao_mercado")))
		if (id_default<>"") And (id_default=x) then
			strResp = strResp & "<option selected"
			ha_default=True
		else
			strResp = strResp & "<option"
			end if
		strResp = strResp & " value='" & x & "'>"
		strResp = strResp & UCase(Trim("" & r("posicao_mercado"))) & "&nbsp;&nbsp;"
		strResp = strResp & "</option>" & chr(13)
		r.MoveNext
		loop

	strResp = "<option selected value=''>&nbsp;</option>" & chr(13) & strResp
		
	posicao_mercado_monta_itens_select = strResp
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




<% if False then 'APENAS P/ HABILITAR O INTELLISENSE DURANTE O DESENVOLVIMENTO!! %>
<script src="../Global/jquery.js" language="JavaScript" type="text/javascript"></script>
<% end if %>

<script src="<%=URL_FILE__JQUERY%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_I18N%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__AJAX_JS%>" language="JavaScript" type="text/javascript"></script>

<script type="text/javascript">
	$(function() {
		$("#c_dt_periodo_inicio").hUtilUI('datepicker_filtro_inicial');
		$("#c_dt_periodo_termino").hUtilUI('datepicker_filtro_final');

		$("#divMsgAguardeObtendoDados").css('filter', 'alpha(opacity=50)');

        $("#c_fabricante").change(function () {
            $("#spnCounterFabricante").text($("#c_fabricante :selected").length);
        });

        $("#c_grupo").change(function () {
            $("#spnCounterGrupo").text($("#c_grupo :selected").length);
        });

        $("#c_subgrupo").change(function () {
            $("#spnCounterSubgrupo").text($("#c_subgrupo :selected").length);
        });

        $("#spnCounterFabricante").text($("#c_fabricante :selected").length);
        $("#spnCounterGrupo").text($("#c_grupo :selected").length);
        $("#spnCounterSubgrupo").text($("#c_subgrupo :selected").length);
	});
</script>

<script language="JavaScript" type="text/javascript">
function limpaCampoSelect(c) {
	c.options[0].selected = true;
}
function limpaCampoSelectFabricante() {
    $("#c_fabricante").children().prop('selected', false);
    $("#spnCounterFabricante").text($("#c_fabricante :selected").length);
}
function limpaCampoSelectGrupo() {
    $("#c_grupo").children().prop('selected', false);
    $("#spnCounterGrupo").text($("#c_grupo :selected").length);
}
function limpaCampoSelectSubgrupo() {
    $("#c_subgrupo").children().prop('selected', false);
    $("#spnCounterSubgrupo").text($("#c_subgrupo :selected").length);
}
function filtra_percentual_crescimento() {
	var letra;
	letra = String.fromCharCode(window.event.keyCode);
	if (((letra < "0") || (letra > "9")) && (letra != "-") && (letra != ".") && (letra != ",")) window.event.keyCode = 0;
}

function fFILTROConfirma( f ) {
    var s_de, s_ate, s_hoje;

    if (trim(f.c_dt_periodo_inicio.value) == "") {
        alert("Informe a data de início do período de vendas!!");
        f.c_dt_periodo_inicio.focus();
        return;
    }

    if (!isDate(f.c_dt_periodo_inicio)) {
        alert("Data inválida!!");
        f.c_dt_periodo_inicio.focus();
        return;
    }

    if (trim(f.c_dt_periodo_termino.value) == "") {
        alert("Informe a data de término do período de vendas!!");
        f.c_dt_periodo_termino.focus();
        return;
    }

    if (!isDate(f.c_dt_periodo_termino)) {
        alert("Data inválida!!");
        f.c_dt_periodo_termino.focus();
        return;
    }

    s_de = trim(f.c_dt_periodo_inicio.value);
    s_ate = trim(f.c_dt_periodo_termino.value);
    if ((s_de != "") && (s_ate != "")) {
        s_de = retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(s_de));
        s_ate = retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(s_ate));
        if (s_de > s_ate) {
            alert("Data de término é menor que a data de início!!");
            f.c_dt_periodo_termino.focus();
            return;
        }
    }

    s_ate = trim(f.c_dt_periodo_termino.value);
    if (s_ate != "") {
        s_ate = retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(s_ate));
        s_hoje = retorna_so_digitos(f.c_DtHojeYYYYMMDD.value);
        if (s_ate > s_hoje) {
            alert("Data de término não pode ser uma data futura!!");
            f.c_dt_periodo_termino.focus();
            return;
        }
    }

	window.status = "Aguarde ...";
	fFILTRO.action = "RelFarolResumidoExec.asp";
	f.submit();
}

</script>

<script type="text/javascript">
    function geraArquivoXLSv2(f) {
    var serverVariableUrl, strUrl, xmlHttp;
    var i, dt_inicio, dt_termino, fabricante, grupo, subgrupo, valorVisao;
    var s_de, s_ate, s_hoje, lojaAux;

    if (trim(f.c_dt_periodo_inicio.value) == "") {
        alert("Informe a data de início do período de vendas!!");
        f.c_dt_periodo_inicio.focus();
        return;
    }

    if (!isDate(f.c_dt_periodo_inicio)) {
        alert("Data inválida!!");
        f.c_dt_periodo_inicio.focus();
        return;
    }

    if (trim(f.c_dt_periodo_termino.value) == "") {
        alert("Informe a data de término do período de vendas!!");
        f.c_dt_periodo_termino.focus();
        return;
    }

    if (!isDate(f.c_dt_periodo_termino)) {
        alert("Data inválida!!");
        f.c_dt_periodo_termino.focus();
        return;
    }

    s_de = trim(f.c_dt_periodo_inicio.value);
    s_ate = trim(f.c_dt_periodo_termino.value);
    if ((s_de != "") && (s_ate != "")) {
        s_de = retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(s_de));
        s_ate = retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(s_ate));
        if (s_de > s_ate) {
            alert("Data de término é menor que a data de início!!");
            f.c_dt_periodo_termino.focus();
            return;
        }
    }

    s_ate = trim(f.c_dt_periodo_termino.value);
    if (s_ate != "") {
        s_ate = retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(s_ate));
        s_hoje = retorna_so_digitos(f.c_DtHojeYYYYMMDD.value);
        if (s_ate > s_hoje) {
            alert("Data de término não pode ser uma data futura!!");
            f.c_dt_periodo_termino.focus();
            return;
        }
    }

    var visao = document.getElementsByName('rb_visao');
       i = 0;
       for(i; i<visao.length; i++){
          if(visao[i].checked){
             visaoValor = visao[i].value;
             break;
          }
        }
    
    fabricante = "";
	grupo = "";
    subgrupo = "";
    dt_inicio = f.c_dt_periodo_inicio.value;
    dt_termino = f.c_dt_periodo_termino.value;

    for (i = 0; i < f.c_fabricante.length; i++) {
        if (f.c_fabricante[i].selected == true) {
            if (fabricante != "") fabricante += "_";
            fabricante += f.c_fabricante[i].value;
        }
    }
    for (i = 0; i < f.c_grupo.length; i++) {
        if (f.c_grupo[i].selected == true) {
            if (grupo != "") grupo += "|";
            grupo += f.c_grupo[i].value;
        }
    }

    for (i = 0; i < f.c_subgrupo.length; i++) {
        if (f.c_subgrupo[i].selected == true) {
            if (subgrupo != "") subgrupo += "|";
            subgrupo += f.c_subgrupo[i].value;
        }
    }

    serverVariableUrl = '<%=Request.ServerVariables("URL")%>';
    serverVariableUrl = serverVariableUrl.toUpperCase();
    serverVariableUrl = serverVariableUrl.substring(0, serverVariableUrl.indexOf("CENTRAL"));

    xmlhttp = GetXmlHttpObject();
    if (xmlhttp == null) {
        alert("O browser NÃO possui suporte ao AJAX!!");
        return;
    }
    
    window.status = "Aguarde, gerando arquivo ...";
    divMsgAguardeObtendoDados.style.visibility = "";

    lojaAux = f.c_loja.value + "";
    while (lojaAux.indexOf("\n") != -1) {
        lojaAux = lojaAux.replace("\n", ",");
    }

    strUrl = 'http://<%=Request.ServerVariables("SERVER_NAME")%>:<%=Request.ServerVariables("SERVER_PORT")%>' + serverVariableUrl + 'WebAPI/api/Farol/GetXLSReport/';
    strUrl = strUrl + '?usuario=<%=usuario%>';
    strUrl = strUrl + '&dt_inicio=' + dt_inicio;
    strUrl = strUrl + '&dt_termino=' + dt_termino;
    strUrl = strUrl + '&fabricante=' + fabricante;
	strUrl = strUrl + '&grupo=' + grupo;
    strUrl = strUrl + '&subgrupo=' + subgrupo;
    strUrl = strUrl + '&btu=' + f.c_potencia_BTU.value;
    strUrl = strUrl + '&ciclo=' + f.c_ciclo.value;
    strUrl = strUrl + '&pos_mercado=' + f.c_posicao_mercado.value;
    strUrl = strUrl + '&perc_est_cresc=' + f.c_perc_est_cresc.value;
    strUrl = strUrl + '&loja=' + lojaAux;
    strUrl = strUrl + '&visao=' + visaoValor;
    

    xmlhttp.onreadystatechange = function () {
        var xmlResp;

        if (xmlhttp.readyState == AJAX_REQUEST_IS_COMPLETE) {
            xmlResp = JSON.parse(xmlhttp.responseText);

            if (xmlResp.Status == "OK") {

            	fFILTRO.action = 'http://<%=Request.ServerVariables("SERVER_NAME")%>:<%=Request.ServerVariables("SERVER_PORT")%>' + serverVariableUrl + 'WebAPI/api/Farol/downloadXLS/?fileName=' + xmlResp.fileName;
                fFILTRO.submit();

                window.status = "Concluído";
                divMsgAguardeObtendoDados.style.visibility = "hidden";
            }
            else if (xmlResp.Status == "Falha") {
                window.status = "Concluído";
                divMsgAguardeObtendoDados.style.visibility = "hidden";

                alert("Falha ao gerar o arquivo XLS\n" + xmlResp.Exception);
                return;
            }
            else if (xmlResp.Status == "Vazio") {
                window.status = "Concluído";
                divMsgAguardeObtendoDados.style.visibility = "hidden";

                alert(xmlResp.Exception);
                return;
            }
        }
    }

    xmlhttp.open("POST", strUrl, true);
    xmlhttp.send();

}
</script>

<script type="text/javascript">
    function geraArquivoXLSv3(f) {
    var serverVariableUrl, strUrl, xmlHttp;
    var i, dt_inicio, dt_termino, fabricante, grupo, subgrupo, valorVisao;
    var s_de, s_ate, s_hoje, lojaAux;

    if (trim(f.c_dt_periodo_inicio.value) == "") {
        alert("Informe a data de início do período de vendas!!");
        f.c_dt_periodo_inicio.focus();
        return;
    }

    if (!isDate(f.c_dt_periodo_inicio)) {
        alert("Data inválida!!");
        f.c_dt_periodo_inicio.focus();
        return;
    }

    if (trim(f.c_dt_periodo_termino.value) == "") {
        alert("Informe a data de término do período de vendas!!");
        f.c_dt_periodo_termino.focus();
        return;
    }

    if (!isDate(f.c_dt_periodo_termino)) {
        alert("Data inválida!!");
        f.c_dt_periodo_termino.focus();
        return;
    }

    s_de = trim(f.c_dt_periodo_inicio.value);
    s_ate = trim(f.c_dt_periodo_termino.value);
    if ((s_de != "") && (s_ate != "")) {
        s_de = retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(s_de));
        s_ate = retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(s_ate));
        if (s_de > s_ate) {
            alert("Data de término é menor que a data de início!!");
            f.c_dt_periodo_termino.focus();
            return;
        }
    }

    s_ate = trim(f.c_dt_periodo_termino.value);
    if (s_ate != "") {
        s_ate = retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(s_ate));
        s_hoje = retorna_so_digitos(f.c_DtHojeYYYYMMDD.value);
        if (s_ate > s_hoje) {
            alert("Data de término não pode ser uma data futura!!");
            f.c_dt_periodo_termino.focus();
            return;
        }
    }

    var visao = document.getElementsByName('rb_visao');
       i = 0;
       for(i; i<visao.length; i++){
          if(visao[i].checked){
             visaoValor = visao[i].value;
             break;
          }
        }
    
    fabricante = "";
    grupo = "";
    subgrupo = "";
	dt_inicio = f.c_dt_periodo_inicio.value;
    dt_termino = f.c_dt_periodo_termino.value;

    for (i = 0; i < f.c_fabricante.length; i++) {
        if (f.c_fabricante[i].selected == true) {
            if (fabricante != "") fabricante += "_";
            fabricante += f.c_fabricante[i].value;
        }
    }
    for (i = 0; i < f.c_grupo.length; i++) {
        if (f.c_grupo[i].selected == true) {
            if (grupo != "") grupo += "|";
            grupo += f.c_grupo[i].value;
        }
    }
    for (i = 0; i < f.c_subgrupo.length; i++) {
        if (f.c_subgrupo[i].selected == true) {
            if (subgrupo != "") subgrupo += "|";
            subgrupo += f.c_subgrupo[i].value;
        }
    }

    serverVariableUrl = '<%=Request.ServerVariables("URL")%>';
    serverVariableUrl = serverVariableUrl.toUpperCase();
    serverVariableUrl = serverVariableUrl.substring(0, serverVariableUrl.indexOf("CENTRAL"));

    xmlhttp = GetXmlHttpObject();
    if (xmlhttp == null) {
        alert("O browser NÃO possui suporte ao AJAX!!");
        return;
    }
    
    window.status = "Aguarde, gerando arquivo ...";
    divMsgAguardeObtendoDados.style.visibility = "";

    lojaAux = f.c_loja.value + "";
    while (lojaAux.indexOf("\n") != -1) {
        lojaAux = lojaAux.replace("\n", ",");
    }

    strUrl = 'http://<%=Request.ServerVariables("SERVER_NAME")%>:<%=Request.ServerVariables("SERVER_PORT")%>' + serverVariableUrl + 'WebAPI/api/FarolV3/GetXLSReport/';
    strUrl = strUrl + '?usuario=<%=usuario%>';
    strUrl = strUrl + '&dt_inicio=' + dt_inicio;
    strUrl = strUrl + '&dt_termino=' + dt_termino;
    strUrl = strUrl + '&fabricante=' + fabricante;
	strUrl = strUrl + '&grupo=' + grupo;
    strUrl = strUrl + '&subgrupo=' + subgrupo;
    strUrl = strUrl + '&btu=' + f.c_potencia_BTU.value;
    strUrl = strUrl + '&ciclo=' + f.c_ciclo.value;
    strUrl = strUrl + '&pos_mercado=' + f.c_posicao_mercado.value;
    strUrl = strUrl + '&perc_est_cresc=' + f.c_perc_est_cresc.value;
    strUrl = strUrl + '&loja=' + lojaAux;
    strUrl = strUrl + '&visao=' + visaoValor;
    

    xmlhttp.onreadystatechange = function () {
        var xmlResp;

        if (xmlhttp.readyState == AJAX_REQUEST_IS_COMPLETE) {
            xmlResp = JSON.parse(xmlhttp.responseText);

            if (xmlResp.Status == "OK") {

            	fFILTRO.action = 'http://<%=Request.ServerVariables("SERVER_NAME")%>:<%=Request.ServerVariables("SERVER_PORT")%>' + serverVariableUrl + 'WebAPI/api/FarolV3/downloadXLS/?fileName=' + xmlResp.fileName;
                fFILTRO.submit();

                window.status = "Concluído";
                divMsgAguardeObtendoDados.style.visibility = "hidden";
            }
            else if (xmlResp.Status == "Falha") {
                window.status = "Concluído";
                divMsgAguardeObtendoDados.style.visibility = "hidden";

                alert("Falha ao gerar o arquivo XLS\n" + xmlResp.Exception);
                return;
            }
            else if (xmlResp.Status == "Vazio") {
                window.status = "Concluído";
                divMsgAguardeObtendoDados.style.visibility = "hidden";

                alert(xmlResp.Exception);
                return;
            }
        }
    }

    xmlhttp.open("POST", strUrl, true);
    xmlhttp.send();

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
<link href="<%=URL_FILE__JQUERY_UI_CSS%>" rel="stylesheet" type="text/css">

<style type="text/css">
.LST
{
	margin:6px 6px 6px 6px;
}
</style>
	</head>

<body onload="focus();">
<center>

    <div id="divMsgAguardeObtendoDados" name="divMsgAguardeObtendoDados" style="background-image: url('../Imagem/ajax_loader_gray_256.gif');background-repeat:no-repeat;background-position: center center;position:absolute;bottom:0px;left:0px;width:100%;height:100%;z-index:9;border: 1pt solid #C0C0C0;background-color: black;opacity:0.6;visibility:hidden;vertical-align: middle">

	</div>

<form id="fFILTRO" name="fFILTRO" method="post">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="c_DtHojeYYYYMMDD" id="c_DtHojeYYYYMMDD" value='<%=strDtHojeYYYYMMDD%>'>
<input type="hidden" name="c_DtHojeDDMMYYYY" id="c_DtHojeDDMMYYYY" value='<%=strDtHojeDDMMYYYY%>'>


<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Farol Resumido</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>
<br>

<table class="Qx" cellspacing="0">
	<!--  PERÍODO DE VENDAS  -->
	<tr bgcolor="#FFFFFF">
	<td class="MT" align="left" nowrap>
		<span class="PLTe">PERÍODO DE VENDAS</span>
		<br>
		<table cellspacing="0" cellpadding="0">
			<tr bgcolor="#FFFFFF">
			<td align="left">
				<input class="PLLc" maxlength="10" style="width:70px;" name="c_dt_periodo_inicio" id="c_dt_periodo_inicio" onfocus="this.select();" onblur="if (!isDate(this)) {alert('Data de início inválida!'); this.focus();}" onkeypress="if (digitou_enter(true)) fFILTRO.c_dt_periodo_termino.focus(); filtra_data();"
					value='<%=get_default_valor_texto_bd(usuario, "RelFarolResumidoFiltro|c_dt_periodo_inicio")%>'
					/>&nbsp;<span class="PLLc" style="color:#808080;">&nbsp;&nbsp;&nbsp;até&nbsp;</span>&nbsp;<input class="PLLc" maxlength="10" style="width:70px;" name="c_dt_periodo_termino" id="c_dt_periodo_termino" onfocus="this.select();" onblur="if (!isDate(this)) {alert('Data de término inválida!'); this.focus();}" onkeypress="if (digitou_enter(true)) fFILTRO.c_fabricante.focus(); filtra_data();"
					value='<%=get_default_valor_texto_bd(usuario, "RelFarolResumidoFiltro|c_dt_periodo_termino")%>'
					/>
			</td>
			</tr>
		</table>
	</td>
	</tr>
	<!--  FABRICANTE  -->
	<tr bgcolor="#FFFFFF">
	<td class="ME MD MB" align="left" nowrap>
		<span class="PLTe">FABRICANTE</span>
		<br>
		<table cellpadding="0" cellspacing="0">
		<tr>
		<td>
			<select id="c_fabricante" name="c_fabricante" class="LST" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;" size="10"style="width:100px" multiple>
			<% =fabricante_monta_itens_select(get_default_valor_texto_bd(usuario, "RelFarolResumidoFiltro|c_fabricante")) %>
			</select>
		</td>
		<td style="width:1px;"></td>
		<td align="left" valign="top">
			<a name="bLimparFabricante" id="bLimparFabricante" href="javascript:limpaCampoSelectFabricante()" title="limpa o filtro 'Fabricante'">
						<img src="../botao/botao_x_red.gif" style="vertical-align:bottom;margin-bottom:1px;" width="20" height="20" border="0"></a>
                        <br />
                        (<span class="Lbl" id="spnCounterFabricante"></span>)
		</td>
		</tr>
		</table>
	</td>
	</tr>
	<!-- GRUPO DE PRODUTOS -->
	<tr bgcolor="#FFFFFF">
	<td class="ME MD MB" align="left" nowrap>
		<span class="PLTe">GRUPO DE PRODUTOS</span>
		<br>
		<table cellpadding="0" cellspacing="0">
		<tr>
		<td>
			<select id="c_grupo" name="c_grupo" class="LST" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;" size="10" style="min-width:250px" multiple>
			<% =grupo_monta_itens_select(get_default_valor_texto_bd(usuario, "RelFarolResumidoFiltro|c_grupo")) %>
			</select>
		</td>
		<td style="width:1px;"></td>
		<td align="left" valign="top">
			<a name="bLimparGrupo" id="bLimparGrupo" href="javascript:limpaCampoSelectGrupo()" title="limpa o filtro 'Grupo de Produtos'">
						<img src="../botao/botao_x_red.gif" style="vertical-align:bottom;margin-bottom:1px;" width="20" height="20" border="0"></a>
                        <br />
                        (<span class="Lbl" id="spnCounterGrupo"></span>)
		</td>
		</tr>
		</table>
	</td>
	</tr>
	<!-- SUBGRUPO DE PRODUTOS -->
	<tr bgcolor="#FFFFFF">
	<td class="ME MD MB" align="left" nowrap>
		<span class="PLTe">SUBGRUPO DE PRODUTOS</span>
		<br>
		<table cellpadding="0" cellspacing="0">
		<tr>
		<td>
			<select id="c_subgrupo" name="c_subgrupo" class="LST" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;" size="10" style="min-width:250px" multiple>
			<% =subgrupo_monta_itens_select(get_default_valor_texto_bd(usuario, "RelFarolResumidoFiltro|c_subgrupo")) %>
			</select>
		</td>
		<td style="width:1px;"></td>
		<td align="left" valign="top">
			<a name="bLimparSubgrupo" id="bLimparSubgrupo" href="javascript:limpaCampoSelectSubgrupo()" title="limpa o filtro 'Subgrupo de Produtos'">
						<img src="../botao/botao_x_red.gif" style="vertical-align:bottom;margin-bottom:1px;" width="20" height="20" border="0"></a>
                        <br />
                        (<span class="Lbl" id="spnCounterSubgrupo"></span>)
		</td>
		</tr>
		</table>
	</td>
	</tr>
	<!-- BTU/h -->
	<tr bgcolor="#FFFFFF">
	<td class="ME MD MB" align="left" nowrap>
		<span class="PLTe">BTU/H</span>
		<br>
		<table cellpadding="0" cellspacing="0">
		<tr>
		<td>
			<select id="c_potencia_BTU" name="c_potencia_BTU" class="LST" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
			<% =potencia_BTU_monta_itens_select(get_default_valor_texto_bd(usuario, "RelFarolResumidoFiltro|c_potencia_BTU")) %>
			</select>
		</td>
		<td style="width:1px;"></td>
		<td align="left" valign="middle">
			<a name="bLimparPotenciaBTU" id="bLimparPotenciaBTU" href="javascript:limpaCampoSelect(fFILTRO.c_potencia_BTU)" title="limpa o filtro 'BTU/h'">
						<img src="../botao/botao_x_red.gif" style="vertical-align:bottom;margin-bottom:1px;" width="20" height="20" border="0"></a>
		</td>
		</tr>
		</table>
	</td>
	</tr>
	<!-- CICLO -->
	<tr bgcolor="#FFFFFF">
	<td class="ME MD MB" align="left" nowrap>
		<span class="PLTe">CICLO</span>
		<br>
		<table cellpadding="0" cellspacing="0">
		<tr>
		<td>
			<select id="c_ciclo" name="c_ciclo" class="LST" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
			<% =ciclo_monta_itens_select(get_default_valor_texto_bd(usuario, "RelFarolResumidoFiltro|c_ciclo")) %>
			</select>
		</td>
		<td style="width:1px;"></td>
		<td align="left" valign="middle">
			<a name="bLimparCiclo" id="bLimparCiclo" href="javascript:limpaCampoSelect(fFILTRO.c_ciclo)" title="limpa o filtro 'Ciclo'">
						<img src="../botao/botao_x_red.gif" style="vertical-align:bottom;margin-bottom:1px;" width="20" height="20" border="0"></a>
		</td>
		</tr>
		</table>
	</td>
	</tr>
	<!-- POSIÇÃO MERCADO -->
	<tr bgcolor="#FFFFFF">
	<td class="ME MD MB" align="left" nowrap>
		<span class="PLTe">POSIÇÃO MERCADO</span>
		<br>
		<table cellpadding="0" cellspacing="0">
		<tr>
		<td>
			<select id="c_posicao_mercado" name="c_posicao_mercado" class="LST" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
			<% =posicao_mercado_monta_itens_select(get_default_valor_texto_bd(usuario, "RelFarolResumidoFiltro|c_posicao_mercado")) %>
			</select>
		</td>
		<td style="width:1px;"></td>
		<td align="left" valign="middle">
			<a name="bLimparPosicaoMercado" id="bLimparPosicaoMercado" href="javascript:limpaCampoSelect(fFILTRO.c_posicao_mercado)" title="limpa o filtro 'Posição Mercado'">
						<img src="../botao/botao_x_red.gif" style="vertical-align:bottom;margin-bottom:1px;" width="20" height="20" border="0"></a>
		</td>
		</tr>
		</table>
	</td>
	</tr>

    <!--  LOJA  -->
	<tr bgcolor="#FFFFFF">
	<td class="MDBE" align="left" nowrap><span class="PLTe">LOJA(S)</span>
	<br>
		<textarea class="PLBe" style="width:100px;font-size:9pt;margin-left:7px;margin-bottom:4px;" rows="8" name="c_loja" id="c_loja" onkeypress="if (!digitou_enter(false) && !digitou_char('-')) filtra_numerico();" onblur="this.value=normaliza_lista_lojas(this.value);"><%=substitui_caracteres(get_default_valor_texto_bd(usuario, "RelFarolResumidoFiltro|c_loja"), ",", chr(10))%></textarea>
	</td></tr>
	<!--  PERCENTUAL ESTIMADO DE CRESCIMENTO  -->
	<tr bgcolor="#FFFFFF">
	<td class="ME MD MB" align="left" nowrap>
		<span class="PLTe">% ESTIMADO DE CRESCIMENTO</span>
		<br>
		<input class="PLLd" maxlength="5" style="width:70px;" id="c_perc_est_cresc" name="c_perc_est_cresc"
			value='<%=get_default_valor_texto_bd(usuario, "RelFarolResumidoFiltro|c_perc_est_cresc")%>'
			onfocus="this.select();"
			onkeypress="if (digitou_enter(true)) bCONFIRMA.focus(); filtra_percentual_crescimento();"
			onblur="this.value=formata_numero(this.value,1);"><span class="C" style="margin-left:2px;">%</span>
	</td>
	</tr>

    <!--  VISÃO: SINTÉTICA/ANALÍTICA  -->
    <tr bgcolor="#FFFFFF">
        <td class="MDBE" align="left" nowrap><span class="PLTe">VISÃO</span>
	    <br>
	    <table cellspacing="0" cellpadding="0" style="margin-bottom:4px;">
	    <tr bgcolor="#FFFFFF">
		    <td align="left">
		    <input type="radio" tabindex="-1" id="rb_visao" name="rb_visao"
			    value="ANALITICA" <% if get_default_valor_texto_bd(usuario, "RelFarolResumidoFiltro|rb_visao") = "ANALITICA" then Response.Write " checked"%> /><span class="C" style="cursor:default" 
			    onclick="fFILTRO.rb_visao[0].click();">Analítica</span>
		        </td>
	    </tr>
	    <tr bgcolor="#FFFFFF">
		    <td align="left">
		    <input type="radio" tabindex="-1" id="rb_visao" name="rb_visao"
			    value="SINTETICA" <% if get_default_valor_texto_bd(usuario, "RelFarolResumidoFiltro|rb_visao") = "SINTETICA" then Response.Write " checked"%> /><span class="C" style="cursor:default" 
			    onclick="fFILTRO.rb_visao[1].click();">Sintética</span>
		    </td>
	</tr>       
	</table>
</td>
</tr>

    <!--  RELATÓRIO ANTIGO  -->
	<tr bgcolor="#FFFFFF">
	<td class="ME MD MB" style="padding-bottom:8px;" align="left" nowrap>
		<span class="PLTe">RELATÓRIO FAROL (Versões anteriores)</span>
		<br />
		<a href="javascript:fFILTROConfirma(fFILTRO)" style="color:#000;font-weight:bold;"><div class="Button" style="width:150px;margin-left:40px;padding:3px;color:black;text-align:center;">Gerar relatório (v1)</div></a>
        <br />
		<a href="javascript:geraArquivoXLSv2(fFILTRO);" style="color:#000;font-weight:bold;"><div class="Button" style="width:150px;margin-left:40px;padding:3px;color:black;text-align:center;">Gerar relatório (v2)</div></a>
	</td>
	</tr>
</table>


<!-- ************   SEPARADOR   ************ -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc" align="left">&nbsp;</td></tr>
</table>
<br>

<table width="649" cellspacing="0">
<tr>
	<td align="left"><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()" title="volta para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
	<td align="right"><div name="dCONFIRMA" id="dCONFIRMA"><a name="bCONFIRMA" id="bCONFIRMA" href="javascript:geraArquivoXLSv3(fFILTRO);" title="confirma a operação">
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