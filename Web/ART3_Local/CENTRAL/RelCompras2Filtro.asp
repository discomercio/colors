<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ===============================================
'	  RelCompras2Filtro.asp
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

	dim usuario, s, intIdx
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	CONECTA COM O BANCO DE DADOS
	dim cn
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	if (Not operacao_permitida(OP_CEN_REL_COMPRAS2, s_lista_operacoes_permitidas)) then 
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	dim dtMinDtInicialFiltroPeriodo, intMaxDiasDtInicialFiltroPeriodo
	dim strMinDtInicialFiltroPeriodoYYYYMMDD, strMinDtInicialFiltroPeriodoDDMMYYYY
	if operacao_permitida(OP_CEN_RESTRINGE_DT_INICIAL_FILTRO_PERIODO, s_lista_operacoes_permitidas) then
		intMaxDiasDtInicialFiltroPeriodo = obtem_max_dias_dt_inicial_filtro_periodo()
		dtMinDtInicialFiltroPeriodo = Date - intMaxDiasDtInicialFiltroPeriodo
		strMinDtInicialFiltroPeriodoYYYYMMDD = formata_data_yyyymmdd(dtMinDtInicialFiltroPeriodo)
		strMinDtInicialFiltroPeriodoDDMMYYYY = formata_data(dtMinDtInicialFiltroPeriodo)
	else
		strMinDtInicialFiltroPeriodoYYYYMMDD = ""
		strMinDtInicialFiltroPeriodoDDMMYYYY = ""
		end if

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
dim x, r, strResp, ha_default, strSql, v, i
	id_default = Trim("" & id_default)
	v = split(id_default, ", ")
	ha_default=False
	strSql = "SELECT DISTINCT" & _
				" Coalesce(grupo,'') AS grupo" & _
			" FROM t_PRODUTO" & _
			" WHERE" & _
				" (Coalesce(grupo,'') <> '')" & _
			" ORDER BY" & _
				" Coalesce(grupo,'')"
	set r = cn.Execute(strSql)
	strResp = ""
	do while Not r.eof 
	    
		x = Trim("" & r("grupo"))
		strResp = strResp & "<option "
            for i=LBound(v) to UBound(v) 
		        if (id_default<>"") And (v(i)=x) then
		            strResp = strResp & "selected"
		         end if
		   	 next
		strResp = strResp & " value='" & x & "'>"
		strResp = strResp & Trim("" & r("grupo")) & "&nbsp;&nbsp;"
		strResp = strResp & "</option>" & chr(13)
		r.MoveNext	
 	loop
		
	grupo_monta_itens_select = strResp
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
<script src="<%=URL_FILE__JQUERY_UI_I18N%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__AJAX_JS%>" language="JavaScript" type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
    $(function () {
        $("#c_dt_inicio").hUtilUI('datepicker_filtro_inicial');
        $("#c_dt_termino").hUtilUI('datepicker_filtro_final');

        $("#c_dt_nf_inicio").hUtilUI('datepicker_filtro_inicial');
        $("#c_dt_nf_termino").hUtilUI('datepicker_filtro_final');

        $("#divMsgAguardeObtendoDados").css('filter', 'alpha(opacity=50)');
    });


function fFILTROConsulta( f ) {
var i, b;
var s_de, s_ate;
var strDtRefYYYYMMDD, strDtRefDDMMYYYY;

//  PERÍODO
	if (trim(f.c_dt_inicio.value)=="") {
		alert("Informe a data inicial do período!!");
		f.c_dt_inicio.focus();
		return;
		}
	
	if (trim(f.c_dt_termino.value)=="") {
		alert("Informe a data final do período!!");
		f.c_dt_termino.focus();
		return;
		}
		
	if (trim(f.c_dt_inicio.value)!="") {
		if (!isDate(f.c_dt_inicio)) {
			alert("Data inválida!!");
			f.c_dt_inicio.focus();
			return;
			}
		}

	if (trim(f.c_dt_termino.value)!="") {
		if (!isDate(f.c_dt_termino)) {
			alert("Data inválida!!");
			f.c_dt_termino.focus();
			return;
			}
		}

	s_de = trim(f.c_dt_inicio.value);
	s_ate = trim(f.c_dt_termino.value);
	if ((s_de!="")&&(s_ate!="")) {
		s_de=retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(s_de));
		s_ate=retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(s_ate));
		if (s_de > s_ate) {
			alert("Data de término é menor que a data de início!!");
			f.c_dt_termino.focus();
			return;
			}
		}

//  Período de consulta está restrito por perfil de acesso?
	if (trim(f.c_MinDtInicialFiltroPeriodoYYYYMMDD.value)!="") {
	// PERÍODO
		strDtRefDDMMYYYY = trim(f.c_dt_inicio.value);
		if (trim(strDtRefDDMMYYYY)!="") {
			strDtRefYYYYMMDD = retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(strDtRefDDMMYYYY));
			if (strDtRefYYYYMMDD < f.c_MinDtInicialFiltroPeriodoYYYYMMDD.value) {
				alert("Data inválida para consulta: " + strDtRefDDMMYYYY + "\nO período de consulta não pode compreender datas anteriores a " + f.c_MinDtInicialFiltroPeriodoDDMMYYYY.value + "!!");
				return;
				}
			}
		strDtRefDDMMYYYY = trim(f.c_dt_termino.value);
		if (trim(strDtRefDDMMYYYY)!="") {
			strDtRefYYYYMMDD = retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(strDtRefDDMMYYYY));
			if (strDtRefYYYYMMDD < f.c_MinDtInicialFiltroPeriodoYYYYMMDD.value) {
				alert("Data inválida para consulta: " + strDtRefDDMMYYYY + "\nO período de consulta não pode compreender datas anteriores a " + f.c_MinDtInicialFiltroPeriodoDDMMYYYY.value + "!!");
				return;
				}
			}
		}

	if (trim(f.c_produto.value)!="") {
		if (!isEAN(trim(f.c_produto.value))) {
			if (trim(f.c_fabricante.value)=="") {
				alert("Informe o fabricante do produto!!");
				f.c_fabricante.focus();
				return;
				}
			}
		}

//  DATA NF ENTRADA
    if (trim(f.c_dt_nf_inicio.value) != "") {
        if (!isDate(f.c_dt_nf_inicio)) {
            alert("Data inválida!!");
            f.c_dt_nf_inicio.focus();
            return;
        }
    }

    if (trim(f.c_dt_nf_termino.value) != "") {
        if (!isDate(f.c_dt_nf_termino)) {
            alert("Data inválida!!");
            f.c_dt_nf_termino.focus();
            return;
        }
    }

    s_de = trim(f.c_dt_nf_inicio.value);
    s_ate = trim(f.c_dt_nf_termino.value);
    if ((s_de != "") && (s_ate != "")) {
        s_de = retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(s_de));
        s_ate = retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(s_ate));
        if (s_de > s_ate) {
            alert("Emissão NF Entrada: data de término é menor que a data de início!!");
            f.c_dt_nf_termino.focus();
            return;
        }
    }


	b=false;
	for (i=0; i<f.rb_detalhe.length; i++) {
		if (f.rb_detalhe[i].checked) {
			b=true;
			break;
			}
		}
	if (!b) {
		alert("Selecione o tipo de detalhamento da consulta!!");
		return;
		}

	
	window.status = "Aguarde ...";
	fFILTRO.action = "RelCompras2Exec.asp";
	f.submit();
}
function limpaCampoSelectFabricante() {
    $("#c_fabricante").children().prop('selected', false);
}
function limpaCampoSelectProduto() {
    $("#c_grupo").children().prop('selected', false);
}
function limpaCampoSelect(c) {
    c.options[0].selected = true;
}
</script>
    <script type="text/javascript">
        function geraArquivoXLS(f) {
            var serverVariableUrl, strUrl, xmlHttp;
            var i, dt_inicio, dt_termino, fabricante, grupo, dt_nf_inicio, dt_nf_termino, valorVisao;
            var s_de, s_ate, s_hoje, b;

            if (trim(f.c_dt_inicio.value) == "") {
                alert("Informe a data de início do período de vendas!!");
                f.c_dt_inicio.focus();
                return;
            }

            if (!isDate(f.c_dt_inicio)) {
                alert("Data inválida!!");
                f.c_dt_inicio.focus();
                return;
            }

            if (trim(f.c_dt_termino.value) == "") {
                alert("Informe a data de término do período de vendas!!");
                f.c_dt_termino.focus();
                return;
            }

            if (!isDate(f.c_dt_termino)) {
                alert("Data inválida!!");
                f.c_dt_termino.focus();
                return;
            }

            s_de = trim(f.c_dt_inicio.value);
            s_ate = trim(f.c_dt_termino.value);
            if ((s_de != "") && (s_ate != "")) {
                s_de = retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(s_de));
                s_ate = retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(s_ate));
                if (s_de > s_ate) {
                    alert("Data de término é menor que a data de início!!");
                    f.c_dt_termino.focus();
                    return;
                }
            }

            s_ate = trim(f.c_dt_termino.value);
            if (s_ate != "") {
                s_ate = retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(s_ate));
                s_hoje = retorna_so_digitos(f.c_DtHojeYYYYMMDD.value);
                if (s_ate > s_hoje) {
                    alert("Data de término não pode ser uma data futura!!");
                    f.c_dt_termino.focus();
                    return;
                }
            }

            if (!isDate(f.c_dt_nf_inicio)) {
                alert("Data inválida!!");
                f.c_dt_nf_inicio.focus();
                return;
            }

            if (!isDate(f.c_dt_nf_termino)) {
                alert("Data inválida!!");
                f.c_dt_nf_termino.focus();
                return;
            }

            s_de = trim(f.c_dt_nf_inicio.value);
            s_ate = trim(f.c_dt_nf_termino.value);
            if ((s_de != "") && (s_ate != "")) {
                s_de = retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(s_de));
                s_ate = retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(s_ate));
                if (s_de > s_ate) {
                    alert("Emissão NF Entrada: data de término é menor que a data de início!!");
                    f.c_dt_nf_termino.focus();
                    return;
                }
            }

            var detalhamento = document.getElementsByName('rb_detalhe');
            var detalhamentoValor;
            i = 0;
            for (i; i < detalhamento.length; i++) {
                if (detalhamento[i].checked) {
                    detalhamentoValor = detalhamento[i].value;
                    break;
                }
            }

            fabricante = "";
            grupo = "";
            dt_inicio = f.c_dt_inicio.value;
            dt_termino = f.c_dt_termino.value;
            dt_nf_inicio = f.c_dt_nf_inicio.value;
            dt_nf_termino = f.c_dt_nf_termino.value;

            for (i = 0; i < f.c_fabricante.length; i++) {
                if (f.c_fabricante[i].selected == true) {
                    if (fabricante != "") fabricante += "_";
                    fabricante += f.c_fabricante[i].value;
                }
            }
            for (i = 0; i < f.c_grupo.length; i++) {
                if (f.c_grupo[i].selected == true) {
                    if (grupo != "") grupo += "_";
                    grupo += f.c_grupo[i].value;
                }
            }
            b = false;
            for (i = 0; i < f.rb_detalhe.length; i++) {
                if (f.rb_detalhe[i].checked) {
                    b = true;
                    break;
                }
            }
            if (!b) {
                alert("Selecione o tipo de detalhamento da consulta!!");
                return;
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

            strUrl = 'http://<%=Request.ServerVariables("SERVER_NAME")%>:<%=Request.ServerVariables("SERVER_PORT")%>' + serverVariableUrl + 'WebAPI/api/Relatorios/GetCompras2CSV/';
            strUrl = strUrl + '?usuario=<%=usuario%>';
            strUrl = strUrl + '&dt_inicio=' + dt_inicio;
            strUrl = strUrl + '&dt_termino=' + dt_termino;
            strUrl = strUrl + '&fabricante=' + fabricante;
            strUrl = strUrl + '&produto=' + f.c_produto.value;
            strUrl = strUrl + '&grupo=' + grupo;
            strUrl = strUrl + '&btu=' + f.c_potencia_BTU.value;
            strUrl = strUrl + '&ciclo=' + f.c_ciclo.value;
            strUrl = strUrl + '&pos_mercado=' + f.c_posicao_mercado.value;
            strUrl = strUrl + '&nf=' + f.c_nf.value;
            strUrl = strUrl + '&dt_nf_inicio=' + dt_nf_inicio;
            strUrl = strUrl + '&dt_nf_termino=' + dt_nf_termino;
            strUrl = strUrl + '&visao=' + "ANALITICA";
            strUrl = strUrl + '&detalhamento=' + detalhamentoValor;



            xmlhttp.onreadystatechange = function () {
                var xmlResp;

                if (xmlhttp.readyState == AJAX_REQUEST_IS_COMPLETE) {
                    xmlResp = JSON.parse(xmlhttp.responseText);

                    if (xmlResp.Status == "OK") {

                    	fFILTRO.action = 'http://<%=Request.ServerVariables("SERVER_NAME")%>:<%=Request.ServerVariables("SERVER_PORT")%>' + serverVariableUrl + 'WebAPI/api/Relatorios/downloadCompras2CSV/?fileName=' + xmlResp.fileName;
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

<link href="<%=URL_FILE__E_CSS%>" Rel="stylesheet" Type="text/css">
<link href="<%=URL_FILE__JQUERY_UI_CSS%>" rel="stylesheet" type="text/css">

<style type="text/css">
#rb_detalhe {
	margin: 0pt 2pt 0pt 15pt;
	vertical-align: top;
	}
.rbOpt
{
	vertical-align:bottom;
}
.lblOpt
{
	vertical-align:bottom;
}
</style>


<body onload="if (trim(fFILTRO.c_dt_inicio.value)=='') fFILTRO.c_dt_inicio.focus();">
<center>


    <div id="divMsgAguardeObtendoDados" name="divMsgAguardeObtendoDados" style="background-image: url('../Imagem/ajax_loader_gray_256.gif');background-repeat:no-repeat;background-position: center center;position:absolute;bottom:0px;left:0px;width:100%;height:100%;z-index:9;border: 1pt solid #C0C0C0;background-color: black;opacity:0.6;visibility:hidden;vertical-align: middle">

	</div>
<form id="fFILTRO" name="fFILTRO" method="post" >
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="c_MinDtInicialFiltroPeriodoYYYYMMDD" id="c_MinDtInicialFiltroPeriodoYYYYMMDD" value='<%=strMinDtInicialFiltroPeriodoYYYYMMDD%>'>
<input type="hidden" name="c_MinDtInicialFiltroPeriodoDDMMYYYY" id="c_MinDtInicialFiltroPeriodoDDMMYYYY" value='<%=strMinDtInicialFiltroPeriodoDDMMYYYY%>'>
<input type="hidden" name="c_DtHojeYYYYMMDD" id="c_DtHojeYYYYMMDD" value='<%=strDtHojeYYYYMMDD%>'>
<input type="hidden" name="c_DtHojeDDMMYYYY" id="c_DtHojeDDMMYYYY" value='<%=strDtHojeDDMMYYYY%>'>

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Relatório de Compras II</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>
<br>

<!--  PARÂMETROS DA CONSULTA  -->
<table class="Qx" cellSpacing="0">
<!--  PERÍODO  -->
	<tr bgColor="#FFFFFF">
	<td class="MT" colspan="2" NOWRAP>
		<table cellSpacing="2" cellPadding="0"><tr bgColor="#FFFFFF"><td>
		<span class="PLTe" style="cursor:default">PERÍODO</span>
		<br>
            <input class="PLLc" maxlength="10" style="width:70px;" name="c_dt_inicio" id="c_dt_inicio" onfocus="this.select();" onblur="if (!isDate(this)) {alert('Data de início inválida!'); this.focus();}" onkeypress="if (digitou_enter(true)) fFILTRO.c_dt_termino.focus(); filtra_data();"
					value='<%=get_default_valor_texto_bd(usuario, "RelCompras2Filtro|c_dt_inicio")%>'
					/>&nbsp;<span class="PLLc" style="color:#808080;">&nbsp;&nbsp;&nbsp;até&nbsp;</span>&nbsp;<input class="PLLc" maxlength="10" style="width:70px; " name="c_dt_termino" id="c_dt_termino" onfocus="this.select();" onblur="if (!isDate(this)) {alert('Data de término inválida!'); this.focus();}" onkeypress="if (digitou_enter(true)) fFILTRO.c_fabricante.focus(); filtra_data();"
					value='<%=get_default_valor_texto_bd(usuario, "RelCompras2Filtro|c_dt_termino")%>'
					/>
			</td></tr>
		</table>
		</td></tr>
<!--  FABRICANTE  -->
	<tr bgcolor="#FFFFFF">
	<td class="ME MD MB" align="left" nowrap>
		<span class="PLTe">FABRICANTE</span>
		<br>
		<table cellpadding="0" cellspacing="0">
		<tr>
		<td>
			<select id="c_fabricante" name="c_fabricante" class="LST" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;" size="10"style="width:100px;margin:4px 4px 4px 4px" multiple>
			<% =fabricante_monta_itens_select(get_default_valor_texto_bd(usuario, "RelCompras2Filtro|c_fabricante")) %>
			</select>
		</td>
		<td style="width:1px;"></td>
		<td align="left" valign="top">
			<a name="bLimparFabricante" id="bLimparFabricante" href="javascript:limpaCampoSelectFabricante()" title="limpa o filtro 'Fabricante'">
						<img src="../botao/botao_x_red.gif" style="vertical-align:bottom;margin-bottom:1px;" width="20" height="20" border="0"></a>
		</td>
		</tr>
		</table>
	</td>
	</tr>
    <!--  PRODUTO  -->
	<tr bgColor="#FFFFFF">
	<td class="ME MD MB" align="left"><span class="PLTe">Produto</span>
		<br><input name="c_produto" id="c_produto" class="PLLe" maxlength="13" style="margin-left:2pt;width:100px;" onkeypress="if (digitou_enter(true)) c_grupo.focus(); filtra_produto();" onblur="this.value=normaliza_codigo(this.value,TAM_MIN_PRODUTO); this.value=ucase(trim(this.value));"></td>
	</tr>
	<!-- GRUPO DE PRODUTOS -->
	<tr bgcolor="#FFFFFF">
	<td class="ME MD MB" align="left" nowrap>
		<span class="PLTe">GRUPO DE PRODUTOS</span>
		<br>
		<table cellpadding="0" cellspacing="0">
		<tr>
		<td>
			<select id="c_grupo" name="c_grupo" class="LST" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;" size="10"style="width:100px;margin:4px 4px 4px 4px" multiple>
			<% =grupo_monta_itens_select(get_default_valor_texto_bd(usuario, "RelCompras2Filtro|c_grupo")) %>
			</select>
		</td>
		<td style="width:1px;"></td>
		<td align="left" valign="top">
			<a name="bLimparGrupo" id="bLimparGrupo" href="javascript:limpaCampoSelectProduto()" title="limpa o filtro 'Grupo de Produtos'">
						<img src="../botao/botao_x_red.gif" style="vertical-align:bottom;margin-bottom:1px;margin-top:2px" width="20" height="20" border="0"></a>
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
			<select id="c_potencia_BTU" name="c_potencia_BTU" class="LST" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;" style="margin:4px 4px 4px 4px">
			<% =potencia_BTU_monta_itens_select(get_default_valor_texto_bd(usuario, "RelCompras2Filtro|c_potencia_BTU")) %>
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
			<select id="c_ciclo" name="c_ciclo" class="LST" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;" style="margin:4px 4px 4px 4px">
			<% =ciclo_monta_itens_select(get_default_valor_texto_bd(usuario, "RelCompras2Filtro|c_ciclo")) %>
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
			<select id="c_posicao_mercado" name="c_posicao_mercado" class="LST" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;" style="margin:4px 4px 4px 4px">
			<% =posicao_mercado_monta_itens_select(get_default_valor_texto_bd(usuario, "RelCompras2Filtro|c_posicao_mercado")) %>
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
<!--  NF  -->
	<tr bgColor="#FFFFFF">
	<td class="ME MD MB" align="left"><span class="PLTe">Nº Nota Fiscal</span>
		<br><input name="c_nf" id="c_nf" class="PLLe" maxlength="30" style="margin-left:2pt;width:150px;" onblur="this.value=ucase(trim(this.value));"></td>
	</tr>
<!--  DATA DE EMISSÃO DA NOTA FISCAL DE ENTRADA  -->
	<tr bgColor="#FFFFFF">
	<td class="ME MD MB" colspan="2" NOWRAP>
		<table cellSpacing="2" cellPadding="0"><tr bgColor="#FFFFFF"><td>
		<span class="PLTe" style="cursor:default">DATA NF ENTRADA</span>
		<br>
            <input class="PLLc" maxlength="10" style="width:70px;" name="c_dt_nf_inicio" id="c_dt_nf_inicio" onfocus="this.select();" onblur="if (!isDate(this)) {alert('Data de início inválida!'); this.focus();}" onkeypress="if (digitou_enter(true)) fFILTRO.c_dt_nf_termino.focus(); filtra_data();"
					value='<%=get_default_valor_texto_bd(usuario, "RelCompras2Filtro|c_dt_nf_inicio")%>'
					/>&nbsp;<span class="PLLc" style="color:#808080;">&nbsp;&nbsp;&nbsp;até&nbsp;</span>&nbsp;<input class="PLLc" maxlength="10" style="width:70px; " name="c_dt_nf_termino" id="c_dt_nf_termino" onfocus="this.select();" onblur="if (!isDate(this)) {alert('Data de término inválida!'); this.focus();}" onkeypress="if (digitou_enter(true)) fFILTRO.rb_detalhe.focus(); filtra_data();"
					value='<%=get_default_valor_texto_bd(usuario, "RelCompras2Filtro|c_dt_nf_termino")%>'
					/>
			</td></tr>
		</table>
		</td></tr>
<!--  TIPO DE DETALHAMENTO  -->
	<tr bgColor="#FFFFFF">
	<td colspan="2" class="MDBE" NOWRAP><span class="PLTe">TIPO DE DETALHAMENTO</span>
		<% intIdx = -1 %>

        <% intIdx = intIdx + 1 %>
		<br><input type="radio" class="rbOpt" tabindex="-1" id="rb_detalhe" name="rb_detalhe" value="SINTETICO_NF" <% if get_default_valor_texto_bd(usuario, "RelCompras2Filtro|rb_detalhe") = "SINTETICO_NF" then Response.Write " checked"%>><span class="C lblOpt" style="cursor:default" onclick="fFILTRO.rb_detalhe[<%=Cstr(intIdx)%>].click();"
			>Sintético por Nota Fiscal</span>

		<% intIdx = intIdx + 1 %>
		<br><input type="radio" class="rbOpt" tabindex="-1" id="rb_detalhe" name="rb_detalhe" value="SINTETICO_FABR" <% if get_default_valor_texto_bd(usuario, "RelCompras2Filtro|rb_detalhe") = "SINTETICO_FABR" then Response.Write " checked"%>><span class="C lblOpt" style="cursor:default" onclick="fFILTRO.rb_detalhe[<%=Cstr(intIdx)%>].click();"
			>Sintético por Fabricante</span>

		<% intIdx = intIdx + 1 %>
		<br><input type="radio" class="rbOpt" tabindex="-1" id="rb_detalhe" name="rb_detalhe" value="SINTETICO_PROD" <% if get_default_valor_texto_bd(usuario, "RelCompras2Filtro|rb_detalhe") = "SINTETICO_PROD" then Response.Write " checked"%>><span class="C lblOpt" style="cursor:default" onclick="fFILTRO.rb_detalhe[<%=Cstr(intIdx)%>].click();"
			>Sintético por Produto</span>
		
		<% intIdx = intIdx + 1 %>
		<br><input type="radio" class="rbOpt" tabindex="-1" id="rb_detalhe" name="rb_detalhe" value="CUSTO_MEDIO" <% if get_default_valor_texto_bd(usuario, "RelCompras2Filtro|rb_detalhe") = "CUSTO_MEDIO" then Response.Write " checked"%>><span class="C lblOpt" style="cursor:default" onclick="fFILTRO.rb_detalhe[<%=Cstr(intIdx)%>].click();"
			>Valor Referência Médio</span>

		<% intIdx = intIdx + 1 %>
		<br><input type="radio" class="rbOpt" tabindex="-1" id="rb_detalhe" name="rb_detalhe" value="CUSTO_INDIVIDUAL" <% if get_default_valor_texto_bd(usuario, "RelCompras2Filtro|rb_detalhe") = "CUSTO_INDIVIDUAL" then Response.Write " checked"%>><span class="C lblOpt" style="cursor:default" onclick="fFILTRO.rb_detalhe[<%=Cstr(intIdx)%>].click();"
			>Valor Referência Individual</span>
	</td>
	</tr>


    <!--  RELATÓRIO VIA WEBAPI  -->
	<tr bgcolor="#FFFFFF">
	<td class="ME MD MB" style="padding-bottom:8px;" align="left" nowrap>
		<span class="PLTe">RELATÓRIO COMPRAS 2 (Versão antiga)</span>
		<br />
		<a href="javascript:fFILTROConsulta(fFILTRO)" style="color:#000;font-weight:bold;"><div class="Button" style="width:150px;margin-left:50px;margin-top:4px;padding:3px;color:black;text-align:center;">Gerar relatório</div></a>
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
	<td><a name="bCANCELA" id="bCANCELA" href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="cancela a operação">
		<img src="../botao/cancelar.gif" width="176" height="55" border="0"></a></td>
	<td align="right"><div name="dCONFIRMA" id="dCONFIRMA"><a name="bCONFIRMA" id="bCONFIRMA" href="javascript:geraArquivoXLS(fFILTRO)" title="executa a consulta">
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
