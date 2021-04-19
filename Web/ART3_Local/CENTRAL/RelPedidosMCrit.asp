<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/Global.asp"    -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ========================================================
'	  R E L P E D I D O S M C R I T . A S P
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

	dim s_opcao_filtro

'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

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

	dim url_origem
	url_origem = Trim(Request("url_origem"))

	dim s_memoria

	' PREENCHIMENTO DA LISTA DE INDICADORES: GRAVA ÚLTIMA OPÇÃO DE CONSULTA NO BD
	dim lst_indicadores_carrega
	lst_indicadores_carrega = Request.Form("ckb_carrega_indicadores")
	if url_origem = "" then
		' GRAVA PARÂMETRO APENAS SE O ACIONAMENTO FOI REALIZADO A PARTIR DA PÁGINA INICIAL
		call set_default_valor_texto_bd(usuario, "RelPedidosMCrit|c_carrega_indicadores_estatico", lst_indicadores_carrega)
		end if

	' SE ESTA PÁGINA FOI ACIONADA COMO RETORNO DE OUTRA PÁGINA DECORRENTE DA CONSULTA DE UM PEDIDO DA LISTA DE RESULTADOS, RESTAURA OS FILTROS
	dim strJS, c_FormFieldValues
	strJS = ""
	c_FormFieldValues = ""
	if url_origem <> "" then
		c_FormFieldValues = get_default_valor_texto_bd(usuario, "CENTRAL/RelPedidosMCrit|FormFields")
		if c_FormFieldValues <> "" then
			strJS = "	var formString = '" & c_FormFieldValues & "';" & chr(13) & _
					"	stringToForm(formString, $('#fFILTRO'));" & chr(13)
			end if
		end if





' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________


' ____________________________________________________________________________
' VENDEDORES MONTA ITENS SELECT
'
function vendedores_monta_itens_select(byval id_default)
dim x, r, strResp, ha_default, strSql
	id_default = Trim("" & id_default)
	ha_default=False
	strSql = "SELECT DISTINCT usuario, nome_iniciais_em_maiusculas FROM" & _
			 " (" & _
			 "SELECT usuario, nome_iniciais_em_maiusculas FROM t_USUARIO" & _
				" WHERE (vendedor_loja <> 0)" & _
			 " UNION" & _
			 " SELECT t_USUARIO.usuario AS usuario, t_USUARIO.nome_iniciais_em_maiusculas FROM t_USUARIO" & _
				" INNER JOIN t_PERFIL_X_USUARIO ON (t_USUARIO.usuario=t_PERFIL_X_USUARIO.usuario)" & _
				" INNER JOIN t_PERFIL ON (t_PERFIL_X_USUARIO.id_perfil=t_PERFIL.id)" & _
				" INNER JOIN t_PERFIL_ITEM ON (t_PERFIL.id=t_PERFIL_ITEM.id_perfil)" & _
				" WHERE (t_PERFIL_ITEM.id_operacao=" & OP_CEN_ACESSO_TODAS_LOJAS & ")" & _
			 ") AS t" & _
			 " ORDER BY usuario"
	set r = cn.Execute(strSql)
	strResp = ""
	do while Not r.eof 
		x = UCase(Trim("" & r("usuario")))
		if (id_default<>"") And (id_default=x) then
			strResp = strResp & "<OPTION SELECTED"
			ha_default=True
		else
			strResp = strResp & "<OPTION"
			end if
		strResp = strResp & " VALUE='" & x & "'>"
		strResp = strResp & x & " - " & Trim("" & r("nome_iniciais_em_maiusculas"))
		strResp = strResp & "</OPTION>" & chr(13)
		r.MoveNext
		loop

	if Not ha_default then
		strResp = "<OPTION SELECTED VALUE=''>&nbsp;</OPTION>" & chr(13) & strResp
		end if
		
	vendedores_monta_itens_select = strResp
	r.close
	set r=nothing
end function

' ____________________________________________________________________________
' INDICADORES MONTA ITENS SELECT

function indicadores_monta_itens_select(byval id_default, byval incluirItemBrancoSeNaoHouverDefault)
    dim x, r, strResp, ha_default
	    id_default = Trim("" & id_default)
	    ha_default=False
	    set r = cn.Execute("SELECT apelido, razao_social_nome_iniciais_em_maiusculas FROM t_ORCAMENTISTA_E_INDICADOR WHERE (apelido NOT IN ('" & ID_ORCAMENTISTA_E_INDICADOR_RESTRICAO_FP_TODOS & "')) ORDER BY apelido")
	    strResp = ""
	    do while Not r.eof 
		    x = UCase(Trim("" & r("apelido")))
		    if (id_default<>"") And (id_default=x) then
			    strResp = strResp & "<OPTION SELECTED"
			    ha_default=True
		    else
			    strResp = strResp & "<OPTION"
			    end if
		    strResp = strResp & " VALUE='" & x & "'>"
		    strResp = strResp & x & " - " & Trim("" & r("razao_social_nome_iniciais_em_maiusculas"))
		    strResp = strResp & "</OPTION>" & chr(13)
		    r.MoveNext
		    loop

	    if (Not ha_default) And incluirItemBrancoSeNaoHouverDefault then
		    strResp = "<OPTION SELECTED VALUE=''>&nbsp;</OPTION>" & chr(13) & strResp
		    end if
    		
	    indicadores_monta_itens_select = strResp
	    r.close
	    set r=nothing
end function

' ____________________________________________________________________________
' GRUPO PRODUTO MONTA ITENS SELECT
'
function t_produto_grupo_monta_itens_select(byval id_default)
dim x, r, strResp, ha_default, strSql, v, i
	id_default = Trim("" & id_default)
	v = split(id_default, ", ")
	ha_default=False
	strSql = "select distinct Coalesce(grupo, '') as codigo, tPG.descricao from t_produto tP" & _
                   " left join t_produto_grupo tPG on (tP.grupo=tPG.codigo)" & _
                   " where (Coalesce(grupo,'')<>'')" & _
                   " order by Coalesce(grupo,'')"

	set r = cn.Execute(strSql)
	strResp = ""
	do while Not r.eof 
	    
		x = Trim("" & r("codigo"))
		strResp = strResp & "<option "
            for i=LBound(v) to UBound(v) 
		        if (id_default<>"") And (v(i)=x) then
		            strResp = strResp & "selected"
		         end if
		   	 next
		strResp = strResp & " value='" & x & "'>"
		strResp = strResp & Trim("" & r("codigo"))        
        if Trim("" & r("descricao")) <> "" then strResp = strResp & "&nbsp;-&nbsp;" & Trim("" & r("descricao"))
		strResp = strResp & "</option>" & chr(13)
		r.MoveNext	
 	loop
		
	t_produto_grupo_monta_itens_select = strResp
	r.close
	set r=nothing
end function

'----------------------------------------------------------------------------------------------
' grupo_origem_pedido_monta_itens_select
function grupo_origem_pedido_monta_itens_select(byval id_default)
dim x, r, strResp
	id_default = Trim("" & id_default)

	set r = cn.Execute("SELECT * FROM t_CODIGO_DESCRICAO WHERE (grupo='PedidoECommerce_Origem_Grupo') AND (st_inativo=0) ORDER BY ordenacao")
	strResp = ""
	do while Not r.eof 
		x = Trim("" & r("codigo"))
		if (id_default=x) then
			strResp = strResp & "<option selected"
		else
			strResp = strResp & "<option"
			end if
		strResp = strResp & " value='" & x & "'>"
		strResp = strResp & Trim("" & r("descricao"))
		strResp = strResp & "</option>" & chr(13)
		r.MoveNext
		loop
	
    strResp = "<option value=''>&nbsp;</option>" & strResp

	grupo_origem_pedido_monta_itens_select = strResp
	r.close
	set r=nothing
end function

' __________________________________________________
' origem_pedido_monta_itens_select
'
function origem_pedido_monta_itens_select(byval id_default)
dim x, r, strResp
	id_default = Trim("" & id_default)

	set r = cn.Execute("SELECT * FROM t_CODIGO_DESCRICAO WHERE (grupo='PedidoECommerce_Origem') AND (st_inativo=0) ORDER BY ordenacao")
	strResp = ""
	do while Not r.eof 
		x = Trim("" & r("codigo"))
		if (id_default=x) then
			strResp = strResp & "<option selected"
		else
			strResp = strResp & "<option"
			end if
		strResp = strResp & " value='" & x & "'>"
		strResp = strResp & Trim("" & r("descricao"))
		strResp = strResp & "</option>" & chr(13)
		r.MoveNext
		loop
	
    strResp = "<option value=''>&nbsp;</option>" & strResp

	origem_pedido_monta_itens_select = strResp
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
	</head>



<% if False then 'APENAS P/ HABILITAR O INTELLISENSE DURANTE O DESENVOLVIMENTO!! %>
<script src="../Global/jquery.js" language="JavaScript" type="text/javascript"></script>
<% end if %>

<script src="<%=URL_FILE__JQUERY%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_I18N%>" Language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>
<script src="<%=URL_FILE__AJAX_JS%>" language="JavaScript" type="text/javascript"></script>

<script type="text/javascript">
	$(function() {
	<% if strJS <> "" then Response.Write strJS %>

	<% if lst_indicadores_carrega = "" then %>
	    
	    $("#divMsgAguardeObtendoDados").css('filter', 'alpha(opacity=50)');
	    	    	
	if (fFILTRO.c_hidden_reload.value == 1) {
            if (('localStorage' in window) && window['localStorage'] !== null) {
                if ('c_indicador' in localStorage) {
                    $("#c_indicador").html(localStorage.getItem('c_indicador'));
                    $("#c_indicador").prop('selectedIndex', fFILTRO.c_hidden_indice_indicador.value);
                }
            }
        }
        
     <% end if %>
        
		$("#c_dt_cancelado_inicio").hUtilUI('datepicker_filtro_inicial');
		$("#c_dt_cancelado_termino").hUtilUI('datepicker_filtro_final');

	<% if operacao_permitida(OP_CEN_FILTRO_MCRIT_PEDIDOS_ENTREGUES_ENTRE, s_lista_operacoes_permitidas) then %>
		$("#c_dt_entregue_inicio").hUtilUI('datepicker_filtro_inicial');
		$("#c_dt_entregue_termino").hUtilUI('datepicker_filtro_final');
	<% end if %>

	<% if operacao_permitida(OP_CEN_FILTRO_MCRIT_PEDIDOS_COLOCADOS_ENTRE, s_lista_operacoes_permitidas) then %>
		$("#c_dt_cadastro_inicio").hUtilUI('datepicker_filtro_inicial');
		$("#c_dt_cadastro_termino").hUtilUI('datepicker_filtro_final');
	<% end if %>

	<% if operacao_permitida(OP_CEN_FILTRO_MCRIT_PEDIDOS_ENTREGA_MARC_ENTRE, s_lista_operacoes_permitidas) then %>
		$("#c_dt_entrega_inicio").hUtilUI('datepicker_filtro_inicial');
		$("#c_dt_entrega_termino").hUtilUI('datepicker_filtro_final');
	<% end if%>

        $("#c_dt_previsao_entrega_inicio").hUtilUI('datepicker_filtro_inicial');
        $("#c_dt_previsao_entrega_termino").hUtilUI('datepicker_filtro_final');
	
		//Every resize of window
	    $(window).resize(function() {
		    sizeDivAjaxRunning();
	    });

	        //Every scroll of window
	        $(window).scroll(function() {
		        sizeDivAjaxRunning();
	        });
        	
	        //Dynamically assign height
	        function sizeDivAjaxRunning() {
		        var newTop = $(window).scrollTop() + "px";
		        $("#divMsgAguardeObtendoDados").css("top", newTop);
	        }
    $(document).tooltip();
	});

	function limpaCampoSelectProduto() {
	    $("#c_grupo").children().prop("selected", false);
	}
</script>

<script language="JavaScript" type="text/javascript">
function fFILTROConfirma( f ) {
var i, b;
var strDtRefYYYYMMDD, strDtRefDDMMYYYY;

	if (f.c_opcao_filtro_pedidos_entregues_entre.value=="S") {
		if (f.ckb_st_entrega_entregue.checked) {
			if (!consiste_periodo(f.c_dt_entregue_inicio, f.c_dt_entregue_termino)) return;
			}
		}

	if (f.ckb_st_entrega_cancelado.checked) {
		if (!consiste_periodo(f.c_dt_cancelado_inicio, f.c_dt_cancelado_termino)) return;
		}
		
	if (f.c_opcao_filtro_pedidos_colocados_entre.value=="S") {
		if (f.ckb_periodo_cadastro.checked) {
			if (trim(f.c_dt_cadastro_inicio.value)=="") {
				alert("Preencha a data!!");
				f.c_dt_cadastro_inicio.focus();
				return;
				}
			if (trim(f.c_dt_cadastro_termino.value)=="") {
				alert("Preencha a data!!");
				f.c_dt_cadastro_termino.focus();
				return;
				}
			if (!consiste_periodo(f.c_dt_cadastro_inicio, f.c_dt_cadastro_termino)) return;
			}
		}

	if (f.c_opcao_filtro_pedidos_entrega_marc_entre.value=="S") {
		if (f.ckb_entrega_marcada_para.checked) {
			if (trim(f.c_dt_entrega_inicio.value)=="") {
				alert("Preencha a data!!");
				f.c_dt_entrega_inicio.focus();
				return;
				}
			if (trim(f.c_dt_entrega_termino.value)=="") {
				alert("Preencha a data!!");
				f.c_dt_entrega_termino.focus();
				return;
				}
			if (!consiste_periodo(f.c_dt_entrega_inicio, f.c_dt_entrega_termino)) return;
			}
		}

	if (f.ckb_entrega_imediata_nao.checked) {
        if (!consiste_periodo(f.c_dt_previsao_entrega_inicio, f.c_dt_previsao_entrega_termino)) return;
    }

	if (f.ckb_produto.checked) {
		if (trim(f.c_produto.value)!="") {
			if (!isEAN(f.c_produto.value)) {
				if (trim(f.c_fabricante.value)=="") {
					alert("Preencha o código do fabricante!!");
					f.c_fabricante.focus();
					return;
					}
				}
			}
		if ((trim(f.c_produto.value)=="")&&(trim(f.c_fabricante.value)=="")) {
			alert("Preencha o código do produto!!");
			f.c_produto.focus();
			return;
			}
		}
		
	if (f.rb_loja[1].checked) {
		if (converte_numero(f.c_loja.value)==0) {
			alert("Especifique o número da loja!!");
			f.c_loja.focus();
			return;
			}
		}

	if (f.rb_loja[2].checked) {
		if (trim(f.c_loja_de.value)!="") {
			if (converte_numero(f.c_loja_de.value)==0) {
				alert("Número de loja inválido!!");
				f.c_loja_de.focus();
				return;
				}
			}
		if (trim(f.c_loja_ate.value)!="") {
			if (converte_numero(f.c_loja_ate.value)==0) {
				alert("Número de loja inválido!!");
				f.c_loja_ate.focus();
				return;
				}
			}
		if ((trim(f.c_loja_de.value)=="")&&(trim(f.c_loja_ate.value)=="")) {
			alert("Preencha pelo menos um dos campos!!");
			f.c_loja_de.focus();
			return;
			}
		if ((trim(f.c_loja_de.value)!="")&&(trim(f.c_loja_ate.value)!="")) {
			if (converte_numero(f.c_loja_ate.value)<converte_numero(f.c_loja_de.value)) {
				alert("Faixa de lojas inválida!!");
				f.c_loja_ate.focus();
				return;
				}
			}
		}

//  Período de consulta está restrito por perfil de acesso?
	if (trim(f.c_MinDtInicialFiltroPeriodoYYYYMMDD.value)!="") {
	//  ENTREGUE ENTRE
		strDtRefDDMMYYYY = trim(f.c_dt_entregue_inicio.value);
		if (trim(strDtRefDDMMYYYY)!="") {
			strDtRefYYYYMMDD = retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(strDtRefDDMMYYYY));
			if (strDtRefYYYYMMDD < f.c_MinDtInicialFiltroPeriodoYYYYMMDD.value) {
				alert("Data inválida para consulta: " + strDtRefDDMMYYYY + "\nO período de consulta não pode compreender datas anteriores a " + f.c_MinDtInicialFiltroPeriodoDDMMYYYY.value + "!!");
				return;
				}
			}
		strDtRefDDMMYYYY = trim(f.c_dt_entregue_termino.value);
		if (trim(strDtRefDDMMYYYY)!="") {
			strDtRefYYYYMMDD = retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(strDtRefDDMMYYYY));
			if (strDtRefYYYYMMDD < f.c_MinDtInicialFiltroPeriodoYYYYMMDD.value) {
				alert("Data inválida para consulta: " + strDtRefDDMMYYYY + "\nO período de consulta não pode compreender datas anteriores a " + f.c_MinDtInicialFiltroPeriodoDDMMYYYY.value + "!!");
				return;
				}
			}
	//  CANCELADO ENTRE
		strDtRefDDMMYYYY = trim(f.c_dt_cancelado_inicio.value);
		if (trim(strDtRefDDMMYYYY)!="") {
			strDtRefYYYYMMDD = retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(strDtRefDDMMYYYY));
			if (strDtRefYYYYMMDD < f.c_MinDtInicialFiltroPeriodoYYYYMMDD.value) {
				alert("Data inválida para consulta: " + strDtRefDDMMYYYY + "\nO período de consulta não pode compreender datas anteriores a " + f.c_MinDtInicialFiltroPeriodoDDMMYYYY.value + "!!");
				return;
				}
			}
		strDtRefDDMMYYYY = trim(f.c_dt_cancelado_termino.value);
		if (trim(strDtRefDDMMYYYY)!="") {
			strDtRefYYYYMMDD = retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(strDtRefDDMMYYYY));
			if (strDtRefYYYYMMDD < f.c_MinDtInicialFiltroPeriodoYYYYMMDD.value) {
				alert("Data inválida para consulta: " + strDtRefDDMMYYYY + "\nO período de consulta não pode compreender datas anteriores a " + f.c_MinDtInicialFiltroPeriodoDDMMYYYY.value + "!!");
				return;
				}
			}
	//  COLOCADOS ENTRE
		strDtRefDDMMYYYY = trim(f.c_dt_cadastro_inicio.value);
		if (trim(strDtRefDDMMYYYY)!="") {
			strDtRefYYYYMMDD = retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(strDtRefDDMMYYYY));
			if (strDtRefYYYYMMDD < f.c_MinDtInicialFiltroPeriodoYYYYMMDD.value) {
				alert("Data inválida para consulta: " + strDtRefDDMMYYYY + "\nO período de consulta não pode compreender datas anteriores a " + f.c_MinDtInicialFiltroPeriodoDDMMYYYY.value + "!!");
				return;
				}
			}
		strDtRefDDMMYYYY = trim(f.c_dt_cadastro_termino.value);
		if (trim(strDtRefDDMMYYYY)!="") {
			strDtRefYYYYMMDD = retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(strDtRefDDMMYYYY));
			if (strDtRefYYYYMMDD < f.c_MinDtInicialFiltroPeriodoYYYYMMDD.value) {
				alert("Data inválida para consulta: " + strDtRefDDMMYYYY + "\nO período de consulta não pode compreender datas anteriores a " + f.c_MinDtInicialFiltroPeriodoDDMMYYYY.value + "!!");
				return;
				}
			}
	//  DATA DE COLETA (RÓTULO ANTIGO: ENTREGA MARCADA ENTRE)
		strDtRefDDMMYYYY = trim(f.c_dt_entrega_inicio.value);
		if (trim(strDtRefDDMMYYYY)!="") {
			strDtRefYYYYMMDD = retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(strDtRefDDMMYYYY));
			if (strDtRefYYYYMMDD < f.c_MinDtInicialFiltroPeriodoYYYYMMDD.value) {
				alert("Data inválida para consulta: " + strDtRefDDMMYYYY + "\nO período de consulta não pode compreender datas anteriores a " + f.c_MinDtInicialFiltroPeriodoDDMMYYYY.value + "!!");
				return;
				}
			}
		strDtRefDDMMYYYY = trim(f.c_dt_entrega_termino.value);
		if (trim(strDtRefDDMMYYYY)!="") {
			strDtRefYYYYMMDD = retorna_so_digitos(formata_ddmmyyyy_yyyymmdd(strDtRefDDMMYYYY));
			if (strDtRefYYYYMMDD < f.c_MinDtInicialFiltroPeriodoYYYYMMDD.value) {
				alert("Data inválida para consulta: " + strDtRefDDMMYYYY + "\nO período de consulta não pode compreender datas anteriores a " + f.c_MinDtInicialFiltroPeriodoDDMMYYYY.value + "!!");
				return;
				}
			}
		}

	b = false;
	for (i = 0; i < f.rb_saida.length; i++) {
		if (f.rb_saida[i].checked) {
			b = true;
			break;
		}
	}
	if (!b) {
		alert("Selecione o tipo de saída do relatório!!");
		return;
	}
	
	dCONFIRMA.style.visibility="hidden";
	window.status = "Aguarde ...";

	if (f.rb_saida[1].checked) setTimeout('exibe_botao_confirmar()', 10000);
	
    <% if lst_indicadores_carrega = "" then %>
	    if (('localStorage' in window) && window['localStorage'] !== null) {
	        var d = $("#c_indicador").html();
	        localStorage.setItem('c_indicador', d);
	    }
	<% end if %>

	fFILTRO.c_hidden_reload.value = 1;
	fFILTRO.c_hidden_indice_indicador.value = $("#c_indicador option:selected").index();
	fFILTRO.ultimoVendedor.value = fFILTRO.c_vendedor.value;
    
	f.c_FormFieldValues.value=formToString($("#fFILTRO"));

	f.submit();
}

function exibe_botao_confirmar() {
	dCONFIRMA.style.visibility = "";
	window.status = "";
}
</script>

<script type="text/javascript">

    function LimpaListaIndicadores() {
        var f, oOption;
        f = fFILTRO;
        $("#c_indicador").empty();
        $(".aviso").css('display', 'none');

        //  Cria um item vazio
        oOption = document.createElement("OPTION");
        f.c_indicador.add(oOption);
        oOption.innerText = "                                                                                 ";
        oOption.value = "";
        oOption.selected = true;
    }

    function TrataRespostaAjaxListaIndicadores() {
        var f, i, strApelido, strNome, strResp, xmlDoc, oOption, oNodes;
        f = fFILTRO;
        if (objAjaxListaIndicadores.readyState == AJAX_REQUEST_IS_COMPLETE) {
            strResp = objAjaxListaIndicadores.responseText;
            if (strResp == "") {
                window.status = "Concluído";
                divMsgAguardeObtendoDados.style.visibility = "hidden";
                $(".aviso").css('display', 'inline');
                return;
            }

            if (strResp != "") {
                $(".aviso").css('display', 'none');
                try {
                    xmlDoc = objAjaxListaIndicadores.responseXML.documentElement;
                    for (i = 0; i < xmlDoc.getElementsByTagName("registro").length; i++) {
                        oOption = document.createElement("OPTION");
                        f.c_indicador.options.add(oOption);

                        oNodes = xmlDoc.getElementsByTagName("apelido")[i];
                        if (oNodes.childNodes.length > 0) strApelido = oNodes.childNodes[0].nodeValue; else strApelido = "";
                        if (strApelido == null) strApelido = "";
                        oOption.value = strApelido;

                        oNodes = xmlDoc.getElementsByTagName("razao_social_nome")[i];
                        if (oNodes.childNodes.length > 0) strNome = oNodes.childNodes[0].nodeValue; else strNome = "";
                        if (strNome == null) strNome = "";

                        oOption.value = strApelido;
                        oOption.innerText = strApelido + " - " + strNome;
                    }
                }
                catch (e) {
                    alert("Falha na consulta de indicadores!!" + "\n" + e.description);
                }
            }
            window.status = "Concluído";
            divMsgAguardeObtendoDados.style.visibility = "hidden";


        }
    }

    function CarregaListaIndicadores() {
        var f, strUrl;
        f = fFILTRO;
        if (fFILTRO.ultimoVendedor.value == trim(fFILTRO.c_vendedor.value)) {
            return;
        }
        objAjaxListaIndicadores = GetXmlHttpObject();
        if (objAjaxListaIndicadores == null) {
            alert("O browser NÃO possui suporte ao AJAX!!");
            return;
        }

        //  Limpa lista de Indicadores
        LimpaListaIndicadores();
        divMsgAguardeObtendoDados.style.visibility = "";

        strUrl = "../Global/AjaxListaIndicadoresLojaPesqBD.asp?";
        //  Prevents server from using a cached file
        strUrl = strUrl + "sid=" + Math.random() + Math.random();
        if (trim(fFILTRO.c_vendedor.value) != "") {
            strUrl = strUrl + "&vendedor=" + fFILTRO.c_vendedor.value;
        }
        fFILTRO.ultimoVendedor.value = fFILTRO.c_vendedor.value;
        objAjaxListaIndicadores.onreadystatechange = TrataRespostaAjaxListaIndicadores;
        objAjaxListaIndicadores.open("GET", strUrl, true);
        objAjaxListaIndicadores.send(null);
        
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
 
 .aviso {
    font-family: Arial, Helvetica, sans-serif;
	font-size: 8pt;
	font-weight: bold;
	font-style: normal;
	margin: 0pt 0pt 0pt 0pt;
	color: #f00;
    display: none;
 }

</style>

<body>
<center>

<!-- MENSAGEM: "Aguarde, obtendo dados" -->

	<div id="divMsgAguardeObtendoDados" name="divMsgAguardeObtendoDados" style="background-image: url('../Imagem/ajax_loader_gray_256.gif');background-repeat:no-repeat;background-position: center center;position:absolute;bottom:0px;left:0px;width:100%;height:100%;z-index:9;border: 1pt solid #C0C0C0;background-color: black;opacity: 0.6;visibility:hidden;vertical-align: middle">

	</div>

<form id="fFILTRO" name="fFILTRO" method="post" action="RelPedidosMCritExec.asp">
<input type="hidden" id="ultimoVendedor" name="ultimoVendedor" value="x-x-x-x-x-x" />
<input type="hidden" name="c_hidden_reload" id="c_hidden_reload" value="0" />
<input type="hidden" name="c_hidden_indice_indicador" id="c_hidden_indice_indicador" value="" />
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>

<% if operacao_permitida(OP_CEN_FILTRO_MCRIT_PEDIDOS_ENTREGUES_ENTRE, s_lista_operacoes_permitidas) then s_opcao_filtro="S" else s_opcao_filtro="N" %>
<input type="hidden" name="c_opcao_filtro_pedidos_entregues_entre" id="c_opcao_filtro_pedidos_entregues_entre" value="<%=s_opcao_filtro%>">
<% if Not operacao_permitida(OP_CEN_FILTRO_MCRIT_PEDIDOS_ENTREGUES_ENTRE, s_lista_operacoes_permitidas) then %>
<input type="hidden" name="c_dt_entregue_inicio" id="c_dt_entregue_inicio" value="">
<input type="hidden" name="c_dt_entregue_termino" id="c_dt_entregue_termino" value="">
<% end if %>

<% if operacao_permitida(OP_CEN_FILTRO_MCRIT_PEDIDOS_COLOCADOS_ENTRE, s_lista_operacoes_permitidas) then s_opcao_filtro="S" else s_opcao_filtro="N" %>
<input type="hidden" name="c_opcao_filtro_pedidos_colocados_entre" id="c_opcao_filtro_pedidos_colocados_entre" value="<%=s_opcao_filtro%>">
<% if Not operacao_permitida(OP_CEN_FILTRO_MCRIT_PEDIDOS_COLOCADOS_ENTRE, s_lista_operacoes_permitidas) then %>
<input type="hidden" name="c_dt_cadastro_inicio" id="c_dt_cadastro_inicio" value="">
<input type="hidden" name="c_dt_cadastro_termino" id="c_dt_cadastro_termino" value="">
<% end if %>

<% if operacao_permitida(OP_CEN_FILTRO_MCRIT_PEDIDOS_ENTREGA_MARC_ENTRE, s_lista_operacoes_permitidas) then s_opcao_filtro="S" else s_opcao_filtro="N" %>
<input type="hidden" name="c_opcao_filtro_pedidos_entrega_marc_entre" id="c_opcao_filtro_pedidos_entrega_marc_entre" value="<%=s_opcao_filtro%>">
<% if Not operacao_permitida(OP_CEN_FILTRO_MCRIT_PEDIDOS_ENTREGA_MARC_ENTRE, s_lista_operacoes_permitidas) then %>
<input type="hidden" name="c_dt_entrega_inicio" id="c_dt_entrega_inicio" value="">
<input type="hidden" name="c_dt_entrega_termino" id="c_dt_entrega_termino" value="">
<% end if %>

<input type="hidden" name="c_MinDtInicialFiltroPeriodoYYYYMMDD" id="c_MinDtInicialFiltroPeriodoYYYYMMDD" value='<%=strMinDtInicialFiltroPeriodoYYYYMMDD%>'>
<input type="hidden" name="c_MinDtInicialFiltroPeriodoDDMMYYYY" id="c_MinDtInicialFiltroPeriodoDDMMYYYY" value='<%=strMinDtInicialFiltroPeriodoDDMMYYYY%>'>

<input type="hidden" name="c_FormFieldValues" id="c_FormFieldValues" value="" />



<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black;">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Relatório Multicritério de Pedidos</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>
<br>

<!--  PARÂMETROS  -->
<table class="Qx" cellspacing="0" width="650">
<!--  STATUS DE ENTREGA  -->
<tr bgcolor="#FFFFFF">
<td class="MT" align="left" nowrap><span class="PLTe">STATUS DE ENTREGA</span>
	<br>
	<table cellspacing="0" cellpadding="0" style="margin-bottom:10px;">
	<tr bgcolor="#FFFFFF"><td align="left">
		<input type="checkbox" tabindex="-1" id="ckb_st_entrega_esperar" name="ckb_st_entrega_esperar"
			value="<%=ST_ENTREGA_ESPERAR%>"><span class="C" style="cursor:default" 
			onclick="fFILTRO.ckb_st_entrega_esperar.click();">Esperar</span>
		</td></tr>
	<tr bgcolor="#FFFFFF"><td align="left">
		<input type="checkbox" tabindex="-1" id="ckb_st_entrega_split" name="ckb_st_entrega_split"
			value="<%=ST_ENTREGA_SPLIT_POSSIVEL%>"><span class="C" style="cursor:default" 
			onclick="fFILTRO.ckb_st_entrega_split.click();">Split Possível</span>
		</td></tr>
	<tr bgcolor="#FFFFFF"><td align="left">
		<input type="checkbox" tabindex="-1" id="ckb_st_entrega_separar_sem_marc" name="ckb_st_entrega_separar_sem_marc"
			value="<%=ST_ENTREGA_SEPARAR%>"><span class="C" style="cursor:default" 
			onclick="fFILTRO.ckb_st_entrega_separar_sem_marc.click();">A Separar (sem data de coleta)</span>
		</td></tr>
	<tr bgcolor="#FFFFFF"><td align="left">
		<input type="checkbox" tabindex="-1" id="ckb_st_entrega_separar_com_marc" name="ckb_st_entrega_separar_com_marc"
			value="<%=ST_ENTREGA_SEPARAR%>"><span class="C" style="cursor:default" 
			onclick="fFILTRO.ckb_st_entrega_separar_com_marc.click();">A Separar (com data de coleta)</span>
		</td></tr>
	<tr bgcolor="#FFFFFF"><td align="left">
		<input type="checkbox" tabindex="-1" id="ckb_st_entrega_a_entregar_sem_marc" name="ckb_st_entrega_a_entregar_sem_marc"
			value="<%=ST_ENTREGA_A_ENTREGAR%>"><span class="C" style="cursor:default" 
			onclick="fFILTRO.ckb_st_entrega_a_entregar_sem_marc.click();">A Entregar (sem data de coleta)</span>
		</td></tr>
	<tr bgcolor="#FFFFFF"><td align="left">
		<input type="checkbox" tabindex="-1" id="ckb_st_entrega_a_entregar_com_marc" name="ckb_st_entrega_a_entregar_com_marc"
			value="<%=ST_ENTREGA_A_ENTREGAR%>"><span class="C" style="cursor:default" 
			onclick="fFILTRO.ckb_st_entrega_a_entregar_com_marc.click();">A Entregar (com data de coleta)</span>
		</td></tr>
	<% if operacao_permitida(OP_CEN_FILTRO_MCRIT_PEDIDOS_ENTREGUES_ENTRE, s_lista_operacoes_permitidas) then %>
	<tr bgcolor="#FFFFFF"><td align="left">
		<input type="checkbox" tabindex="-1" id="ckb_st_entrega_entregue" name="ckb_st_entrega_entregue" onclick="if (fFILTRO.ckb_st_entrega_entregue.checked) fFILTRO.c_dt_entregue_inicio.focus();"
			value="<%=ST_ENTREGA_ENTREGUE%>"><span class="C" style="cursor:default" 
			onclick="fFILTRO.ckb_st_entrega_entregue.click();">Entregue entre</span
			><input class="Cc" maxlength="10" style="width:70px;" name="c_dt_entregue_inicio" id="c_dt_entregue_inicio" onblur="if (!isDate(this)) {alert('Data inválida!'); this.focus();}" onkeypress="if (digitou_enter(true)) fFILTRO.c_dt_entregue_termino.focus(); else fFILTRO.ckb_st_entrega_entregue.checked=true; filtra_data();" onclick="fFILTRO.ckb_st_entrega_entregue.checked=true;" onchange="fFILTRO.ckb_st_entrega_entregue.checked=true;"
			>&nbsp;<span class="C">e</span>&nbsp;<input class="Cc" maxlength="10" style="width:70px;" name="c_dt_entregue_termino" id="c_dt_entregue_termino" onblur="if (!isDate(this)) {alert('Data inválida!'); this.focus();}" onkeypress="if (digitou_enter(true)) bCONFIRMA.focus(); else fFILTRO.ckb_st_entrega_entregue.checked=true; filtra_data();" onclick="fFILTRO.ckb_st_entrega_entregue.checked=true;" onchange="fFILTRO.ckb_st_entrega_entregue.checked=true;">
		</td></tr>
	<% end if %>
	<tr bgcolor="#FFFFFF"><td align="left">
		<input type="checkbox" tabindex="-1" id="ckb_st_entrega_cancelado" name="ckb_st_entrega_cancelado" onclick="if (fFILTRO.ckb_st_entrega_cancelado.checked) fFILTRO.c_dt_cancelado_inicio.focus();"
			value="<%=ST_ENTREGA_CANCELADO%>"><span class="C" style="cursor:default" 
			onclick="fFILTRO.ckb_st_entrega_cancelado.click();">Cancelado entre</span
			><input class="Cc" maxlength="10" style="width:70px;" name="c_dt_cancelado_inicio" id="c_dt_cancelado_inicio" onblur="if (!isDate(this)) {alert('Data inválida!'); this.focus();}" onkeypress="if (digitou_enter(true)) fFILTRO.c_dt_cancelado_termino.focus(); else fFILTRO.ckb_st_entrega_cancelado.checked=true; filtra_data();" onclick="fFILTRO.ckb_st_entrega_cancelado.checked=true;" onchange="fFILTRO.ckb_st_entrega_cancelado.checked=true;"
			>&nbsp;<span class="C">e</span>&nbsp;<input class="Cc" maxlength="10" style="width:70px;" name="c_dt_cancelado_termino" id="c_dt_cancelado_termino" onblur="if (!isDate(this)) {alert('Data inválida!'); this.focus();}" onkeypress="if (digitou_enter(true)) bCONFIRMA.focus(); else fFILTRO.ckb_st_entrega_cancelado.checked=true; filtra_data();" onclick="fFILTRO.ckb_st_entrega_cancelado.checked=true;" onchange="fFILTRO.ckb_st_entrega_cancelado.checked=true;">
        <span class="C" style="cursor:default">ordenado por</span>
        <select name="c_cancelados_ordena">
            <option value="VENDEDOR" selected>Vendedor</option>
            <option value="PEDIDO">Pedido</option>
        </select>
		</td></tr>
	<tr bgcolor="#FFFFFF"><td align="left">
		<input type="checkbox" tabindex="-1" id="ckb_st_entrega_exceto_cancelados" name="ckb_st_entrega_exceto_cancelados"
			value="<%=ST_ENTREGA_CANCELADO%>"><span class="C" style="cursor:default" 
			onclick="fFILTRO.ckb_st_entrega_exceto_cancelados.click();">Exceto Cancelados</span>
		</td></tr>
	<tr bgcolor="#FFFFFF"><td align="left">
		<input type="checkbox" tabindex="-1" id="ckb_st_entrega_exceto_entregues" name="ckb_st_entrega_exceto_entregues"
			value="<%=ST_ENTREGA_ENTREGUE%>"><span class="C" style="cursor:default" 
			onclick="fFILTRO.ckb_st_entrega_exceto_entregues.click();">Exceto Entregues</span>
		</td></tr>
	</table>
</td></tr>

<!--  PEDIDOS RECEBIDOS PELO CLIENTE  -->
<tr bgcolor="#FFFFFF">
<td class="MDBE" align="left" nowrap><span class="PLTe">PEDIDOS RECEBIDOS PELO CLIENTE</span>
	<br>
	<table cellspacing="0" cellpadding="0" style="margin-bottom:10px;">
    <tr bgcolor="#FFFFFF"><td align="left">
		<input type="checkbox" tabindex="-1" id="ckb_pedido_nao_recebido_pelo_cliente" name="ckb_pedido_nao_recebido_pelo_cliente"
			value="ON"><span class="C" style="cursor:default" 
			onclick="fFILTRO.ckb_pedido_nao_recebido_pelo_cliente.click();">Não recebido pelo cliente</span>
		</td></tr>
    <tr bgcolor="#FFFFFF"><td align="left">
		<input type="checkbox" tabindex="-1" id="ckb_pedido_recebido_pelo_cliente" name="ckb_pedido_recebido_pelo_cliente"
			value="ON"><span class="C" style="cursor:default" 
			onclick="fFILTRO.ckb_pedido_recebido_pelo_cliente.click();">Recebido pelo cliente</span>
		</td></tr>
	</table>
</td></tr>

<!--  STATUS DE PAGAMENTO  -->
<tr bgcolor="#FFFFFF">
<td class="MDBE" align="left" nowrap><span class="PLTe">STATUS DE PAGAMENTO</span>
	<br>
	<table cellspacing="0" cellpadding="0" style="margin-bottom:10px;">
	<tr bgcolor="#FFFFFF"><td align="left">
		<input type="checkbox" tabindex="-1" id="ckb_st_pagto_pago" name="ckb_st_pagto_pago"
			value="<%=ST_PAGTO_PAGO%>"><span class="C" style="cursor:default" 
			onclick="fFILTRO.ckb_st_pagto_pago.click();">Pago</span>
		</td></tr>
	<tr bgcolor="#FFFFFF"><td align="left">
		<input type="checkbox" tabindex="-1" id="ckb_st_pagto_nao_pago" name="ckb_st_pagto_nao_pago"
			value="<%=ST_PAGTO_NAO_PAGO%>"><span class="C" style="cursor:default" 
			onclick="fFILTRO.ckb_st_pagto_nao_pago.click();">Não-Pago</span>
		</td></tr>
	<tr bgcolor="#FFFFFF"><td align="left">
		<input type="checkbox" tabindex="-1" id="ckb_st_pagto_pago_parcial" name="ckb_st_pagto_pago_parcial"
			value="<%=ST_PAGTO_PARCIAL%>"><span class="C" style="cursor:default" 
			onclick="fFILTRO.ckb_st_pagto_pago_parcial.click();">Pago Parcial</span>
		</td></tr>
	</table>
</td></tr>

<!--  ANÁLISE DE CRÉDITO  -->
<tr bgcolor="#FFFFFF">
<td class="MDBE" nowrap align="left"><span class="PLTe">ANÁLISE DE CRÉDITO</span>
	<br>
	<table cellspacing="0" cellpadding="0" style="margin-bottom:10px;">
	<tr>
		<td width="50%" valign="top">
			<table cellspacing="0" cellpadding="0">
				<tr bgcolor="#FFFFFF"><td align="left">
					<input type="checkbox" tabindex="-1" id="ckb_analise_credito_st_inicial" name="ckb_analise_credito_st_inicial"
						value="<%=COD_AN_CREDITO_ST_INICIAL%>"><span class="C" style="cursor:default;color:<%=x_analise_credito_cor(COD_AN_CREDITO_ST_INICIAL)%>;" 
						onclick="fFILTRO.ckb_analise_credito_st_inicial.click();">Status Inicial</span>
					</td></tr>
				<tr bgcolor="#FFFFFF"><td align="left">
					<input type="checkbox" tabindex="-1" id="ckb_analise_credito_pendente_vendas" name="ckb_analise_credito_pendente_vendas"
						value="<%=COD_AN_CREDITO_PENDENTE_VENDAS%>"><span class="C" style="cursor:default;color:<%=x_analise_credito_cor(COD_AN_CREDITO_PENDENTE_VENDAS)%>;" 
						onclick="fFILTRO.ckb_analise_credito_pendente_vendas.click();">Pendente Vendas</span>
					</td></tr>
				<tr bgcolor="#FFFFFF"><td align="left">
					<input type="checkbox" tabindex="-1" id="ckb_analise_credito_pendente_endereco" name="ckb_analise_credito_pendente_endereco"
						value="<%=COD_AN_CREDITO_PENDENTE_ENDERECO%>"><span class="C" style="cursor:default;color:<%=x_analise_credito_cor(COD_AN_CREDITO_PENDENTE_ENDERECO)%>;" 
						onclick="fFILTRO.ckb_analise_credito_pendente_endereco.click();">Pendente Endereço</span>
					</td></tr>
				<tr bgcolor="#FFFFFF"><td align="left">
					<input type="checkbox" tabindex="-1" id="ckb_analise_credito_pendente" name="ckb_analise_credito_pendente"
						value="<%=COD_AN_CREDITO_PENDENTE%>"><span class="C" style="cursor:default;color:<%=x_analise_credito_cor(COD_AN_CREDITO_PENDENTE)%>;" 
						onclick="fFILTRO.ckb_analise_credito_pendente.click();">Pendente</span>
					</td></tr>
				<tr bgcolor="#FFFFFF"><td align="left">
					<input type="checkbox" tabindex="-1" id="ckb_analise_credito_pendente_cartao" name="ckb_analise_credito_pendente_cartao"
						value="<%=COD_AN_CREDITO_PENDENTE_CARTAO%>"><span class="C" style="cursor:default;color:<%=x_analise_credito_cor(COD_AN_CREDITO_PENDENTE_CARTAO)%>;" 
						onclick="fFILTRO.ckb_analise_credito_pendente_cartao.click();">Pendente Cartão de Crédito</span>
					</td></tr>
			</table>
		</td>
		<td width="50%" valign="top">
			<table cellspacing="0" cellpadding="0">
				<tr bgcolor="#FFFFFF"><td align="left">
					<input type="checkbox" tabindex="-1" id="ckb_analise_credito_pendente_pagto_antecipado_boleto" name="ckb_analise_credito_pendente_pagto_antecipado_boleto"
						value="<%=COD_AN_CREDITO_PENDENTE_PAGTO_ANTECIPADO_BOLETO%>" /><span class="C" style="cursor:default;color:<%=x_analise_credito_cor(COD_AN_CREDITO_PENDENTE_PAGTO_ANTECIPADO_BOLETO)%>;"
						onclick="fFILTRO.ckb_analise_credito_pendente_pagto_antecipado_boleto.click();"><%=x_analise_credito(COD_AN_CREDITO_PENDENTE_PAGTO_ANTECIPADO_BOLETO)%></span>
					</td></tr>
				<tr bgcolor="#FFFFFF"><td align="left">
					<input type="checkbox" tabindex="-1" id="ckb_analise_credito_ok_aguardando_deposito" name="ckb_analise_credito_ok_aguardando_deposito"
						value="<%=COD_AN_CREDITO_OK_AGUARDANDO_DEPOSITO%>"><span class="C" style="cursor:default;color:<%=x_analise_credito_cor(COD_AN_CREDITO_OK_AGUARDANDO_DEPOSITO)%>;" 
						onclick="fFILTRO.ckb_analise_credito_ok_aguardando_deposito.click();">Crédito OK (aguardando depósito)</span>
					</td></tr>
				<tr bgcolor="#FFFFFF"><td align="left">
					<input type="checkbox" tabindex="-1" id="ckb_analise_credito_ok_deposito_aguardando_desbloqueio" name="ckb_analise_credito_ok_deposito_aguardando_desbloqueio"
						value="<%=COD_AN_CREDITO_OK_DEPOSITO_AGUARDANDO_DESBLOQUEIO%>"><span class="C" style="cursor:default;color:<%=x_analise_credito_cor(COD_AN_CREDITO_OK_DEPOSITO_AGUARDANDO_DESBLOQUEIO)%>;" 
						onclick="fFILTRO.ckb_analise_credito_ok_deposito_aguardando_desbloqueio.click();">Crédito OK (depósito aguardando desbloqueio)</span>
					</td></tr>
				<tr bgcolor="#FFFFFF"><td align="left">
					<input type="checkbox" tabindex="-1" id="ckb_analise_credito_ok_aguardando_pagto_boleto_av" name="ckb_analise_credito_ok_aguardando_pagto_boleto_av"
						value="<%=COD_AN_CREDITO_OK_AGUARDANDO_PAGTO_BOLETO_AV%>" /><span class="C" style="cursor:default;color:<%=x_analise_credito_cor(COD_AN_CREDITO_OK_AGUARDANDO_PAGTO_BOLETO_AV)%>;"
						onclick="fFILTRO.ckb_analise_credito_ok_aguardando_pagto_boleto_av.click();"><%=x_analise_credito(COD_AN_CREDITO_OK_AGUARDANDO_PAGTO_BOLETO_AV)%></span>
					</td></tr>
				<tr bgcolor="#FFFFFF"><td align="left">
					<input type="checkbox" tabindex="-1" id="ckb_analise_credito_ok" name="ckb_analise_credito_ok"
						value="<%=COD_AN_CREDITO_OK%>"><span class="C" style="cursor:default;color:<%=x_analise_credito_cor(COD_AN_CREDITO_OK)%>;" 
						onclick="fFILTRO.ckb_analise_credito_ok.click();">Crédito OK</span>
					</td></tr>
			</table>
		</td>
	</tr>
	</table>
</td></tr>

<!--  ENTREGA IMEDIATA  -->
<tr bgcolor="#FFFFFF">
<td class="MDBE" align="left" nowrap><span class="PLTe">ENTREGA IMEDIATA</span>
	<br>
	<table cellspacing="0" cellpadding="0" width="100%" style="margin-bottom:10px;">
	<tr bgcolor="#FFFFFF"><td align="left">
		<input type="checkbox" tabindex="-1" id="ckb_entrega_imediata_sim" name="ckb_entrega_imediata_sim"
			value="<%=COD_ETG_IMEDIATA_SIM%>"><span class="C" style="cursor:default" 
			onclick="fFILTRO.ckb_entrega_imediata_sim.click();">Sim</span>
		</td></tr>
	<tr bgcolor="#FFFFFF"><td align="left">
		<input type="checkbox" tabindex="-1" id="ckb_entrega_imediata_nao" name="ckb_entrega_imediata_nao"
			value="<%=COD_ETG_IMEDIATA_NAO%>"><span class="C" style="cursor:default" 
			onclick="fFILTRO.ckb_entrega_imediata_nao.click();">Não</span>
			<span style="width:50px;">&nbsp;</span>
			<span class="C" style="cursor:default;" onclick="fFILTRO.ckb_entrega_imediata_nao.click();">Previsão de Entrega entre</span>
			<input class="Cc" maxlength="10" style="width:70px;" name="c_dt_previsao_entrega_inicio" id="c_dt_previsao_entrega_inicio" onblur="if (!isDate(this)) {alert('Data inválida!'); this.focus();}" onkeypress="if (digitou_enter(true)) fFILTRO.c_dt_previsao_entrega_termino.focus(); else fFILTRO.ckb_entrega_imediata_nao.checked=true; filtra_data();" onclick="fFILTRO.ckb_entrega_imediata_nao.checked = true;" onchange="fFILTRO.ckb_entrega_imediata_nao.checked=true;"
			/>&nbsp;<span class="C">e</span>&nbsp;<input class="Cc" maxlength="10" style="width:70px;" name="c_dt_previsao_entrega_termino" id="c_dt_previsao_entrega_termino" onblur="if (!isDate(this)) {alert('Data inválida!'); this.focus();}" onkeypress="if (digitou_enter(true)) bCONFIRMA.focus(); else fFILTRO.ckb_entrega_imediata_nao.checked=true; filtra_data();" onclick="fFILTRO.ckb_entrega_imediata_nao.checked = true;" onchange="fFILTRO.ckb_entrega_imediata_nao.checked=true;" />
		</td></tr>
	</table>
</td></tr>

<!--  GERAL  -->
<tr bgcolor="#FFFFFF">
<td class="MDBE" align="left" nowrap><span class="PLTe">GERAL</span>
	<br>
	<table cellspacing="0" cellpadding="0" style="margin-bottom:10px;">
	<tr bgcolor="#FFFFFF"><td align="left">
		<input type="checkbox" tabindex="-1" id="ckb_obs2_preenchido" name="ckb_obs2_preenchido"
			value="ON"><span class="C" style="cursor:default" 
			onclick="fFILTRO.ckb_obs2_preenchido.click();">OBS II preenchido</span>
		</td></tr>
	<tr bgcolor="#FFFFFF"><td align="left">
		<input type="checkbox" tabindex="-1" id="ckb_obs2_nao_preenchido" name="ckb_obs2_nao_preenchido"
			value="ON"><span class="C" style="cursor:default" 
			onclick="fFILTRO.ckb_obs2_nao_preenchido.click();">OBS II não preenchido</span>
		</td></tr>
	<tr bgcolor="#FFFFFF"><td align="left">
		<%	s_memoria = get_default_valor_texto_bd(usuario, "CENTRAL/RelPedidosMCrit|ckb_nao_exibir_rastreio") %>
		<input type="checkbox" tabindex="-1" id="ckb_nao_exibir_rastreio" name="ckb_nao_exibir_rastreio"
			value="ON" <%if s_memoria <> "" then Response.Write " checked"%> /><span class="C" style="cursor:default" 
			onclick="fFILTRO.ckb_nao_exibir_rastreio.click();">Não exibir link de rastreamento</span>
		</td></tr>
	</table>
</td></tr>

<!--  PEDIDOS COM OU SEM INDICADOR  -->
<tr bgcolor="#FFFFFF">
<td class="MDBE" align="left" nowrap><span class="PLTe">INDICADOR</span>
	<br>
	<table cellspacing="0" cellpadding="0" style="margin-bottom:10px;">
	<tr bgcolor="#FFFFFF"><td align="left">
		<input type="checkbox" tabindex="-1" id="ckb_indicador_preenchido" name="ckb_indicador_preenchido"
			value="ON"><span class="C" style="cursor:default" 
			onclick="fFILTRO.ckb_indicador_preenchido.click();">Indicador preenchido</span>
		</td></tr>
	<tr bgcolor="#FFFFFF"><td align="left">
		<input type="checkbox" tabindex="-1" id="ckb_indicador_nao_preenchido" name="ckb_indicador_nao_preenchido"
			value="ON"><span class="C" style="cursor:default" 
			onclick="fFILTRO.ckb_indicador_nao_preenchido.click();">Indicador não preenchido</span>
		</td></tr>
	</table>
</td></tr>

<% if operacao_permitida(OP_CEN_FILTRO_MCRIT_PEDIDOS_COLOCADOS_ENTRE, s_lista_operacoes_permitidas) Or _
	  operacao_permitida(OP_CEN_FILTRO_MCRIT_PEDIDOS_ENTREGA_MARC_ENTRE, s_lista_operacoes_permitidas) then %>
<!--  PERÍODO  -->
<tr bgcolor="#FFFFFF">
<td class="MDBE" align="left" nowrap><span class="PLTe">PERÍODO</span>
	<br>
	<table cellspacing="0" cellpadding="0" style="margin-bottom:10px;">
	<% if operacao_permitida(OP_CEN_FILTRO_MCRIT_PEDIDOS_COLOCADOS_ENTRE, s_lista_operacoes_permitidas) then %>
	<tr bgcolor="#FFFFFF"><td align="left">
		<input type="checkbox" tabindex="-1" id="ckb_periodo_cadastro" name="ckb_periodo_cadastro" onclick="if (fFILTRO.ckb_periodo_cadastro.checked) fFILTRO.c_dt_cadastro_inicio.focus();"
			value="PERIODO_ON"><span class="C" style="cursor:default" 
			onclick="fFILTRO.ckb_periodo_cadastro.click();">Somente pedidos colocados entre</span
			><input class="Cc" maxlength="10" style="width:70px;" name="c_dt_cadastro_inicio" id="c_dt_cadastro_inicio" onblur="if (!isDate(this)) {alert('Data inválida!'); this.focus();}" onkeypress="if (digitou_enter(true)) fFILTRO.c_dt_cadastro_termino.focus(); else fFILTRO.ckb_periodo_cadastro.checked=true; filtra_data();" onclick="fFILTRO.ckb_periodo_cadastro.checked=true;" onchange="fFILTRO.ckb_periodo_cadastro.checked=true;"
			>&nbsp;<span class="C">e</span>&nbsp;<input class="Cc" maxlength="10" style="width:70px;" name="c_dt_cadastro_termino" id="c_dt_cadastro_termino" onblur="if (!isDate(this)) {alert('Data inválida!'); this.focus();}" onkeypress="if (digitou_enter(true)) bCONFIRMA.focus(); else fFILTRO.ckb_periodo_cadastro.checked=true; filtra_data();" onclick="fFILTRO.ckb_periodo_cadastro.checked=true;" onchange="fFILTRO.ckb_periodo_cadastro.checked=true;">
		</td></tr>
	<% end if %>
	<% if operacao_permitida(OP_CEN_FILTRO_MCRIT_PEDIDOS_ENTREGA_MARC_ENTRE, s_lista_operacoes_permitidas) then %>
	<tr bgcolor="#FFFFFF"><td align="left">
		<input type="checkbox" tabindex="-1" id="ckb_entrega_marcada_para" name="ckb_entrega_marcada_para" onclick="if (fFILTRO.ckb_entrega_marcada_para.checked) fFILTRO.c_dt_entrega_inicio.focus();"
			value="ENTREGA_MARCADA_PARA_ON"><span class="C" style="cursor:default" 
			onclick="fFILTRO.ckb_entrega_marcada_para.click();">Data de coleta no período entre</span
			><input class="Cc" maxlength="10" style="width:70px;" name="c_dt_entrega_inicio" id="c_dt_entrega_inicio" onblur="if (!isDate(this)) {alert('Data inválida!'); this.focus();}" onkeypress="if (digitou_enter(true)) fFILTRO.c_dt_entrega_termino.focus(); else fFILTRO.ckb_entrega_marcada_para.checked=true; filtra_data();" onclick="fFILTRO.ckb_entrega_marcada_para.checked=true;" onchange="fFILTRO.ckb_entrega_marcada_para.checked=true;"
			>&nbsp;<span class="C">e</span>&nbsp;<input class="Cc" maxlength="10" style="width:70px;" name="c_dt_entrega_termino" id="c_dt_entrega_termino" onblur="if (!isDate(this)) {alert('Data inválida!'); this.focus();}" onkeypress="if (digitou_enter(true)) bCONFIRMA.focus(); else fFILTRO.ckb_entrega_marcada_para.checked=true; filtra_data();" onclick="fFILTRO.ckb_entrega_marcada_para.checked=true;" onchange="fFILTRO.ckb_entrega_marcada_para.checked=true;">
		</td></tr>
	<% end if%>
	</table>
</td></tr>
<% end if %>

<!--  PRODUTO  -->
<tr bgcolor="#FFFFFF">
<td class="MDBE" align="left" nowrap><span class="PLTe">PRODUTO</span>
	<br>
	<table cellspacing="0" cellpadding="0" style="margin-bottom:10px;">
	<tr bgcolor="#FFFFFF"><td align="left">
		<input type="checkbox" tabindex="-1" id="ckb_produto" name="ckb_produto" onclick="if (fFILTRO.ckb_produto.checked) fFILTRO.c_fabricante.focus();"
			value="PRODUTO_ON"><span class="C" style="cursor:default" 
			onclick="fFILTRO.ckb_produto.click();">Somente pedidos que incluam:</span
			><br><span class="C" style="margin-left:30px;">Fabricante</span><input maxlength="4" class="Cc" style="width:50px;" name="c_fabricante" id="c_fabricante" onblur="this.value=normaliza_codigo(this.value,TAM_MIN_FABRICANTE);" onkeypress="if (digitou_enter(true)) fFILTRO.c_produto.focus(); else fFILTRO.ckb_produto.checked=true; filtra_fabricante();" onclick="fFILTRO.ckb_produto.checked=true;">
			<span class="C">&nbsp;&nbsp;&nbsp;Produto</span><input maxlength="13" class="Cc" style="width:100px;" name="c_produto" id="c_produto" onblur="this.value=ucase(normaliza_codigo(this.value,TAM_MIN_PRODUTO));" onkeypress="if (digitou_enter(true)) bCONFIRMA.focus(); else fFILTRO.ckb_produto.checked=true; filtra_produto();" onclick="fFILTRO.ckb_produto.checked=true;">
		</td></tr>
	</table>
</td></tr>

<!-- GRUPO DE PRODUTOS -->
<tr bgcolor="#FFFFFF">
	<td class="ME MD MB" colspan="2" align="left" nowrap>
		<span class="PLTe">GRUPO DE PRODUTOS</span>
		<br>
		<table cellpadding="0" cellspacing="0" style="margin:1px 3px 6px 10px;">
		<tr>
		<td>
			<select id="c_grupo" name="c_grupo" class="LST" size="5" style="width:200px" multiple>
			<% =t_produto_grupo_monta_itens_select(Null) %>
			</select>
		</td>
		<td style="width:1px;"></td>
		<td align="left" valign="top">
			<a name="bLimparGrupo" id="bLimparGrupo" href="javascript:limpaCampoSelectProduto()" title="limpa o filtro 'Grupo de Produtos'">
						<img src="../botao/botao_x_red.gif" style="vertical-align:bottom;margin-bottom:1px;" width="20" height="20" border="0"></a>
		</td>
		</tr>
		</table>
	</td>
	</tr>


<!-- ORIGEM DO PEDIDO (GRUPO) -->
    <tr bgcolor="#FFFFFF">
		<td class="MDBE" align="left" nowrap><span class="PLTe">ORIGEM DO PEDIDO (GRUPO)</span>
		<br>           
			<select id="c_grupo_pedido_origem" name="c_grupo_pedido_origem" style="margin:1px 3px 6px 10px;width: 200px" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
			<% =grupo_origem_pedido_monta_itens_select(Null) %>
			</select>
			
        </td></tr>

<!-- ORIGEM DO PEDIDO -->
    <tr bgcolor="#FFFFFF">
		<td class="MDBE" align="left" nowrap><span class="PLTe">ORIGEM DO PEDIDO</span>
		<br>
			<select id="c_pedido_origem" name="c_pedido_origem" style="margin:1px 3px 6px 10px;width: 200px" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
			<% =origem_pedido_monta_itens_select(Null) %>
			</select>
			
        </td></tr>

<!-- EMPRESA -->
    <tr bgcolor="#FFFFFF">
		<td class="MDBE" align="left" nowrap><span class="PLTe">EMPRESA</span>
		<br>
			<select id="c_empresa" name="c_empresa" style="margin:1px 3px 6px 10px;min-width:100px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;">
			<% =apelido_empresa_nfe_emitente_monta_itens_select(Null) %>
			</select>
			
        </td>
    </tr>
    
<!--  LOJAS  -->
<tr bgcolor="#FFFFFF">
<td class="MDBE" align="left" nowrap><span class="PLTe">LOJAS</span>
	<br>
	<table cellspacing="0" cellpadding="0" style="margin-bottom:10px;">
	<tr bgcolor="#FFFFFF"><td align="left">
		<input type="radio" tabindex="-1" id="rb_loja_todas" name="rb_loja"
			value="TODAS" checked><span class="C" style="cursor:default" 
			onclick="fFILTRO.rb_loja[0].click();">Todas as lojas</span>
		</td></tr>
	<tr bgcolor="#FFFFFF"><td align="left">
		<input type="radio" tabindex="-1" id="rb_loja_uma" name="rb_loja" onclick="if (fFILTRO.rb_loja[1].checked) fFILTRO.c_loja.focus();"
			value="UMA"><span class="C" style="cursor:default" 
			onclick="fFILTRO.rb_loja[1].click();">Loja(s)</span>
			<input class="C" maxlength="200" style="width:300px;" name="c_loja" id="c_loja" onblur="this.value=normaliza_codigo(this.value, TAM_MIN_LOJA);" onkeypress="if (digitou_enter(true) && tem_info(this.value)) bCONFIRMA.focus(); else this.click();" onclick="fFILTRO.rb_loja[1].checked=true;" title="Nº da loja (separar múltiplas lojas com espaço em branco ou vírgula)" />
		</td></tr>
	<tr bgcolor="#FFFFFF"><td align="left">
		<input type="radio" tabindex="-1" id="rb_loja_faixa" name="rb_loja" onclick="if (fFILTRO.rb_loja[2].checked) fFILTRO.c_loja_de.focus();"
			value="FAIXA"><span class="C" style="cursor:default"
			onclick="fFILTRO.rb_loja[2].click();">Lojas</span>
			<input class="Cc" maxlength="3" style="width:40px;" name="c_loja_de" id="c_loja_de" onblur="this.value=normaliza_codigo(this.value, TAM_MIN_LOJA);" onkeypress="if (digitou_enter(true) && tem_info(this.value)) fFILTRO.c_loja_ate.focus(); else this.click(); filtra_numerico();" onclick="fFILTRO.rb_loja[2].checked=true;">
			<span class="C">a</span>
			<input class="Cc" maxlength="3" style="width:40px;" name="c_loja_ate" id="c_loja_ate" onblur="this.value=normaliza_codigo(this.value, TAM_MIN_LOJA);" onkeypress="this.click(); if (digitou_enter(true) && tem_info(this.value)) bCONFIRMA.focus(); filtra_numerico();" onclick="fFILTRO.rb_loja[2].checked=true;">
		</td></tr>
	</table>
</td></tr>

<!--  NOVA VERSÃO DA FORMA DE PAGAMENTO  -->
<tr bgcolor="#FFFFFF">
<td class="MDBE" align="left" nowrap><span class="PLTe">FORMA DE PAGAMENTO</span>
	<br>
	<table cellspacing="0" cellpadding="0" style="margin-bottom:10px;">
	<tr bgcolor="#FFFFFF"><td align="left">
		<span class="C" style="margin-left:30px;">Forma de Pagamento</span>
			<select id="op_forma_pagto" name="op_forma_pagto">
			<% =forma_pagto_monta_itens_select(Null) %>
			</select>
		</td></tr>
	<tr bgcolor="#FFFFFF"><td align="left">
		<span class="C" style="margin-left:30px;">Nº Parcelas</span>
			<input class="Cc" maxlength="2" style="width:40px;" name="c_forma_pagto_qtde_parc" id="c_forma_pagto_qtde_parc" onkeypress="if (digitou_enter(true) && tem_info(this.value)) bCONFIRMA.focus(); filtra_numerico();">
		</td></tr>
	</table>
</td></tr>

<!--  CLIENTE  -->
<tr bgcolor="#FFFFFF">
<td class="MDBE" align="left" nowrap><span class="PLTe">CLIENTE</span>
	<br>
	<table cellspacing="0" cellpadding="0" style="margin-bottom:10px;">
	<tr bgcolor="#FFFFFF"><td align="left">
		<span class="C" style="margin-left:30px;">CNPJ/CPF</span>
			<input class="C" maxlength="18" style="width:140px;" name="c_cliente_cnpj_cpf" id="c_cliente_cnpj_cpf" onblur="if (!cnpj_cpf_ok(this.value)) {alert('CNPJ/CPF inválido!!');this.focus();} else this.value=cnpj_cpf_formata(this.value);" onkeypress="if (digitou_enter(true)&&((!tem_info(this.value))||(tem_info(this.value)&&cnpj_cpf_ok(this.value)))) {this.value=cnpj_cpf_formata(this.value); bCONFIRMA.focus();} filtra_cnpj_cpf();">
		</td></tr>
    <tr bgcolor="#FFFFFF"><td align="left">
		<span class="C" style="margin-left:30px;">UF</span>
			<select name="c_cliente_uf" id="c_cliente_uf">
                <%=UF_monta_itens_select(Null) %>
			</select>
		</td></tr>
	</table>
</td></tr>

<!--  CARTÃO DE CRÉDITO  -->
<tr bgcolor="#FFFFFF">
<td class="MDBE" align="left" nowrap><span class="PLTe">CARTÃO DE CRÉDITO</span>
	<br>
	<table cellspacing="0" cellpadding="0" style="margin-bottom:10px;">
	<tr bgcolor="#FFFFFF"><td align="left">
		<input type="checkbox" tabindex="-1" id="ckb_visanet" name="ckb_visanet"
			value="VISANET_ON"><span class="C" style="cursor:default" 
			onclick="fFILTRO.ckb_visanet.click();">Somente pedidos pagos usando cartão de crédito</span>
		</td></tr>
	</table>
</td></tr>

<!--  TRANSPORTADORA  -->
<tr bgcolor="#FFFFFF">
<td class="MDBE" align="left" nowrap><span class="PLTe">TRANSPORTADORA</span>
	<br>
	<table cellspacing="0" cellpadding="0" style="margin-bottom:10px;">
	<tr bgcolor="#FFFFFF"><td align="left">
		<span class="C" style="margin-left:30px;">Identificação</span>
			<input class="C" maxlength="10" style="width:110px;" name="c_transportadora" id="c_transportadora" onblur="this.value=trim(this.value);" onkeypress="if (digitou_enter(true)) bCONFIRMA.focus(); filtra_nome_identificador();">
		</td></tr>
	</table>
</td></tr>

<!--  CADASTRAMENTO  -->
<tr bgcolor="#FFFFFF">
<td class="MDBE" align="left" nowrap><span class="PLTe">CADASTRAMENTO</span>
    	<br>
	<table cellspacing="6" cellpadding="0" style="margin-bottom:0px;">
	<tr bgcolor="#FFFFFF">
		<td style="width:70px; text-align:right"><span class="C" style="text-align: right; margin-left: 2px">Vendedor</span></td>
		<td align="left">
			<select id="c_vendedor" name="c_vendedor" style="margin-right:10px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;" <% if lst_indicadores_carrega = "" then %>onchange="LimpaListaIndicadores()" <% end if %>>
			<% =vendedores_monta_itens_select(Null) %>
			</select>
		</td>
	</tr>
	<tr bgcolor="#FFFFFF">
		<td align="right" valign="top" style="width: 70px"><span class="C" style="margin-left:2px;"><% if lst_indicadores_carrega = "" then %><img id="exclamacao" src="../IMAGEM/exclamacao_14x14.png" title="Reduza o tempo de carregamento da lista de indicadores, filtrando por vendedor." style="cursor:pointer;" />&nbsp;<% end if %>Indicador</span></td>
		<td align="left">
			<select id="c_indicador" name="c_indicador" style="margin-right:10px;" onkeyup="if (window.event.keyCode==KEYCODE_DELETE) this.options[0].selected=true;" <% if lst_indicadores_carrega = "" then %> onfocus="CarregaListaIndicadores();" <% end if %>>
			    <option selected value=''>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</option>
			<% if lst_indicadores_carrega <> "" then
			    Response.Write indicadores_monta_itens_select(Null, False)
			   end if
			 %>
			</select><br />
			<span class="aviso">Vendedor selecionado não possui indicadores.</span>&nbsp;
		</td>
	</tr>
	</table>
</td></tr>

<!--  COLUNAS OPCIONAIS  -->
<tr bgcolor="#FFFFFF">
<td class="MDBE" align="left" nowrap><span class="PLTe">CAMPOS OPCIONAIS</span>
	<br>
	<table cellspacing="0" cellpadding="0" style="margin-bottom:10px;">
	<tr bgcolor="#FFFFFF"><td align="left">
		<%	s_memoria = get_default_valor_texto_bd(usuario, "CENTRAL/RelPedidosMCrit|ckb_exibir_vendedor") %>
		<input type="checkbox" tabindex="-1" id="ckb_exibir_vendedor" name="ckb_exibir_vendedor"
			value="ON" <%if s_memoria <> "" then Response.Write " checked"%> /><span class="C" style="cursor:default" 
			onclick="fFILTRO.ckb_exibir_vendedor.click();">Vendedor</span>
		</td></tr>
	<tr bgcolor="#FFFFFF"><td align="left">
		<%	s_memoria = get_default_valor_texto_bd(usuario, "CENTRAL/RelPedidosMCrit|ckb_exibir_parceiro") %>
		<input type="checkbox" tabindex="-1" id="ckb_exibir_parceiro" name="ckb_exibir_parceiro"
			value="ON" <%if s_memoria <> "" then Response.Write " checked"%> /><span class="C" style="cursor:default" 
			onclick="fFILTRO.ckb_exibir_parceiro.click();">Parceiro</span>
		</td></tr>
	<tr bgcolor="#FFFFFF"><td align="left">
		<%	s_memoria = get_default_valor_texto_bd(usuario, "CENTRAL/RelPedidosMCrit|ckb_exibir_uf") %>
		<input type="checkbox" tabindex="-1" id="ckb_exibir_uf" name="ckb_exibir_uf"
			value="ON" <%if s_memoria <> "" then Response.Write " checked"%> /><span class="C" style="cursor:default" 
			onclick="fFILTRO.ckb_exibir_uf.click();">UF</span>
		</td></tr>
	<tr bgcolor="#FFFFFF"><td align="left">
		<%	s_memoria = get_default_valor_texto_bd(usuario, "CENTRAL/RelPedidosMCrit|ckb_exibir_data_previsao_entrega") %>
		<input type="checkbox" tabindex="-1" id="ckb_exibir_data_previsao_entrega" name="ckb_exibir_data_previsao_entrega"
			value="ON" <%if s_memoria <> "" then Response.Write " checked"%> /><span class="C" style="cursor:default" 
			onclick="fFILTRO.ckb_exibir_data_previsao_entrega.click();">Previsão de Entrega</span>
		</td></tr>
	</table>
</td></tr>

<!--  SAÍDA DO RELATÓRIO  -->
<tr bgcolor="#FFFFFF">
<td class="MDBE" align="left" nowrap><span class="PLTe">Saída do Relatório</span>
	<br><input type="radio" tabindex="-1" id="rb_saida_html" name="rb_saida" value="Html" onclick="dCONFIRMA.style.visibility='';" checked><span class="C" style="cursor:default" onclick="fFILTRO.rb_saida[0].click(); dCONFIRMA.style.visibility='';"
		>Html</span>

	<br><input type="radio" tabindex="-1" id="rb_saida_xls" name="rb_saida" value="XLS" onclick="dCONFIRMA.style.visibility='';"><span class="C" style="cursor:default" onclick="fFILTRO.rb_saida[1].click(); dCONFIRMA.style.visibility='';"
		>Excel</span>
</td></tr>
</table>

<!-- ************   SEPARADOR   ************ -->
<table width="649" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc" align="left">&nbsp;</td></tr>
</table>
<br>


<table width="649" cellspacing="0">
<tr>
	<td align="left"><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()">
		<img src="../botao/voltar.gif" width="176" height="55" border="0" title="volta para a página anterior"></a></td>
	<td align="right"><div name="dCONFIRMA" id="dCONFIRMA"><a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fFILTROConfirma(fFILTRO)">
		<img src="../botao/confirmar.gif" width="176" height="55" border="0" title="executa a consulta"></a></div>
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