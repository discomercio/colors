<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ===================================================================
'	  R E L C O M I S S A O I N D I C A D O R E S F I N A L I Z A . A S P
'     ===================================================================
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
	
	Server.ScriptTimeout = MAX_SERVER_SCRIPT_TIMEOUT_EM_SEG
	
	dim usuario
	usuario = Trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	CONECTA COM O BANCO DE DADOS
	dim cn, cn2, rs, msg_erro

	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)            
    if Not bdd_conecta_RPIFC(cn2) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)    

	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)
	
	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	if Not operacao_permitida(OP_CEN_REL_PEDIDOS_INDICADORES_PAGAMENTO, s_lista_operacoes_permitidas) then 
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	dim alerta
	dim s, s_aux, s_filtro, s_id,c_dt_entregue_mes, c_dt_entregue_ano,vendedor, rb_visao
    dim aviso
    
	alerta = ""

    rb_visao = Request("rb_visao")
    s_id = Request("id")


' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________

    function x_meio_pagto(x)
        dim s
        select case x
           case "DEP" : s = "Pagamento em Depósito (Bradesco)"
           case "DEP1" : s = "Pagamento em Depósito (Banco Inter)"
           case "CHQ" : s = "Pagamento com Cheque"
           case "DIN" : s = "Pagamento em Dinheiro"
           case else : s=""
        end select
    x_meio_pagto = s
    end function
 
function empresa_monta_itens_select()
dim x, r, strResp

    set r = cn2.Execute("SELECT * FROM t_FIN_PLANO_CONTAS_EMPRESA")

	strResp = ""
	do while Not r.eof 
		x = Trim("" & r("id"))

		strResp = strResp & "<option"
        if r("descricao") = "CENTRAL" then strResp = strResp & " selected"

		strResp = strResp & " value='" & x & "'>"
		strResp = strResp & Trim("" & r("id")) & " - " & Trim("" & r("descricao"))
		strResp = strResp & "</option>" & chr(13)
		r.MoveNext
		loop
		
	empresa_monta_itens_select = strResp
	r.close

end function

function conta_corrente_monta_itens_select()
dim x, r, strResp

    set r = cn2.Execute("SELECT * FROM t_FIN_CONTA_CORRENTE")

	strResp = ""
	do while Not r.eof 
		x = Trim("" & r("id"))

		strResp = strResp & "<option"
        if x = "1" then strResp = strResp & " selected"

		strResp = strResp & " value='" & x & "'>"
		strResp = strResp & Trim("" & r("banco")) & " - " & Trim("" & r("agencia")) & " " & Trim("" & r("conta")) & " " & Trim("" & r("descricao"))
		strResp = strResp & "</option>" & chr(13)
		r.MoveNext
		loop

	conta_corrente_monta_itens_select = strResp
	r.close

end function

' _____________________________________
' CONSULTA EXECUTA
'
sub consulta_executa
const VENDA_NORMAL = "VENDA_NORMAL"
const DEVOLUCAO = "DEVOLUCAO"
const PERDA = "PERDA"
dim r
dim s, s_aux, s_sql, x, cab_table, meio_pagto_a, n_reg, n_reg_total, qtde_indicadores, indicador_a, vendedor_a
dim idx_bloco, inc, comissao, num_pag,indice,qtde_vendedor
dim v_Banco,strAuxbanco,blnAchou,vl_aux,n_reg_BD,intIdxBanco,strAuxBancoAnterior,intIdxVetor,strCampoOrdenacao,v_OutrosBancos, num_linhas
dim regex

	'Devido ao código "DEP1" usado para "Pagamento em Depósito (Banco Inter)", é usada a regex para retirar todos os dígitos 1 que estejam no final do código
	set regex = New RegExp
	regex.Pattern = "1+$"

	s_sql = "SELECT " & _
				"*" & _
			" FROM t_COMISSAO_INDICADOR_N1" & _
				" INNER JOIN t_COMISSAO_INDICADOR_N2 ON (t_COMISSAO_INDICADOR_N1.id = t_COMISSAO_INDICADOR_N2.id_comissao_indicador_n1)" & _
				" INNER JOIN t_COMISSAO_INDICADOR_N3 ON (t_COMISSAO_INDICADOR_N2.id = t_COMISSAO_INDICADOR_N3.id_comissao_indicador_n2)" & _
			" WHERE" & _
				" (t_COMISSAO_INDICADOR_N1.id = " & s_id & ")" & _
				" AND (t_COMISSAO_INDICADOR_N3.st_tratamento_manual=0)" & _
			" ORDER BY" & _
				" t_COMISSAO_INDICADOR_N2.vendedor,"& _
				" CASE t_COMISSAO_INDICADOR_N3.meio_pagto WHEN 'DEP' THEN -3 END DESC ," & _
				" CASE t_COMISSAO_INDICADOR_N3.meio_pagto WHEN 'DEP1' THEN -2 END DESC ," & _
				" CASE t_COMISSAO_INDICADOR_N3.meio_pagto WHEN 'CHQ' THEN -1 END DESC ," & _
				" t_COMISSAO_INDICADOR_N3.indicador," & _
				" t_COMISSAO_INDICADOR_N3.numero_banco," & _
				" t_COMISSAO_INDICADOR_N3.vl_total_comissao_liquido_arredondado_NFS"

  ' CABEÇALHO
	cab_table = "<table cellspacing='0' id='tableDados'  style='border:1.2px solid black;'>" & chr(13)
	
	x = ""
	n_reg = 0
	n_reg_total = 0
	idx_bloco = 0
	num_pag=1
	num_linhas = 0
	qtde_vendedor = 1

	meio_pagto_a = "XXXXXXXXXXXX"
	indicador_a = "XXXXXXXXXX"
	vendedor_a = "XXXXXXXX"
	set r = cn.execute(s_sql)

	aviso=""

	if r.Eof then
		aviso="Não há valores a serem pagos."
	else
		c_dt_entregue_mes = r("competencia_mes")
		c_dt_entregue_ano = r("competencia_ano")
		if Trim("" & r("proc_automatico_status"))=1 then
			aviso= "O Relatório já foi processado por " & r("proc_automatico_usuario") & " em " & r("proc_automatico_data_hora") & "."
			end if
		end if

	if aviso <> "" then
		x = "<div class='MtAlerta notPrint' style='width:649px;font-weight:bold;' align='center'><p style='margin:5px 2px 5px 2px;'>" & aviso & "</p></div><br />"
		end if

	do while Not r.Eof
		if (r("vl_total_comissao_liquido_arredondado_NFS") + r("vl_total_RA_liquido_arredondado_NFS")) = 0 then
			if Trim("" & r("vendedor")) <> vendedor_a then
				if vendedor = "" then
					vendedor = vendedor & r("vendedor")
				else
					vendedor = vendedor & ","
					vendedor = vendedor & r("vendedor")
					end if
				end if 'if Trim("" & r("vendedor")) <> vendedor_a

			indicador_a = Trim("" & r("indicador"))
			vendedor_a = Trim("" & r("vendedor"))
			end if 'if (r("vl_total_comissao_liquido_arredondado_NFS") + r("vl_total_RA_liquido_arredondado_NFS")) = 0


		if (r("vl_total_comissao_liquido_arredondado_NFS") + r("vl_total_RA_liquido_arredondado_NFS")) > 0 then
		'	MUDOU DE INDICADOR OU VENDEDOR?
		'	LEMBRANDO QUE O MESMO INDICADOR PODE TER PEDIDOS COM VENDEDORES DIFERENTES
			if (Trim("" & r("indicador")) <> indicador_a) Or (Trim("" & r("vendedor")) <> vendedor_a) then
				idx_bloco = idx_bloco + 1
				if num_linhas > 10 then num_linhas = 9

			  ' FECHA TABELA DO INDICADOR ANTERIOR
				if n_reg_total > 0 then
					inc = inc + 1
					if (inc mod 2) = 0  then
						num_linhas = num_linhas + 1
						x = x & _
							"		</td>" & chr(13) & _
							"	</tr>" & chr(13) & _
							"	<tr>" & chr(13) & _
							"		<td width='50%'>" & chr(13)
					else
						x = x & _
							"		</td>" & chr(13) & _
							"		<td width='50%'>" & chr(13)
						end if
					
					if (Trim("" & r("meio_pagto")) <> meio_pagto_a) Or (Trim("" & r("vendedor")) <> vendedor_a) then
						if (inc mod 2) = 1 then num_linhas = num_linhas + 1
						if n_reg_total > 0 then
							x = x & _
								"		</td>" & chr(13) & _
								"	</tr>" & chr(13) & _
								"</table>" & chr(13) & _
								"<br />" & chr(13)
							end if
						end if

					if ((inc mod 2) = 0) Or (Trim("" & r("meio_pagto")) <> meio_pagto_a) Or (Trim("" & r("vendedor")) <> vendedor_a) then
						if num_pag = 1 then
							if (num_linhas >= 8) Or ((qtde_vendedor >= 4) And (num_linhas >= 6)) then
								if (Trim("" & r("meio_pagto")) <> meio_pagto_a) Or (Trim("" & r("vendedor")) <> vendedor_a) then
									x = x & _
										"<div style='break-after:always'></div>" & chr(13)
								else
									x = x & _
										"		</td>" & chr(13) & _
										"	</tr>" & chr(13) & _
										"</table>" & chr(13) & _
										"<div style='break-after:always'></div>" & chr(13) & _
										"<table width='649' style='border-left:1.2px solid #000;border-right:1.2px solid #000;border-bottom:1.2px solid #000;' cellpadding='0' cellspacing='0'>" & chr(13) & _
										"	<tr>" & chr(13) & _
										"		<td width='50%'>" & chr(13)
									end if

								num_linhas = 1
								qtde_vendedor = 1
								num_pag = 0
								end if
						else
							if (num_linhas >= 9) Or ((qtde_vendedor >= 4) And (num_linhas >= 6)) then
								if (Trim("" & r("meio_pagto")) <> meio_pagto_a) Or (Trim("" & r("vendedor")) <> vendedor_a) then
									x = x & _
									"<div style='break-after:always'></div>" & chr(13)
								else
									x = x & _
										"		</td>" & chr(13) & _
										"	</tr>" & chr(13) & _
										"</table>" & chr(13) & _
										"<div style='break-after:always'></div>" & chr(13) & _
										"<table width='649' style='border-left:1.2px solid #000;border-right:1.2px solid #000;border-bottom:1.2px solid #000;' cellpadding='0' cellspacing='0'>" & chr(13) & _
										"	<tr>" & chr(13) & _
										"		<td width='50%'>" & chr(13)
									end if

								num_linhas = 1
								qtde_vendedor = 1
								end if
							end if
						end if 'if ((inc mod 2) = 0) Or (Trim("" & r("meio_pagto")) <> meio_pagto_a) Or (Trim("" & r("vendedor")) <> vendedor_a)

					Response.Write x

					x = ""
					end if 'if n_reg_total > 0

			'   MUDOU DE VENDEDOR?
				if Trim("" & r("vendedor")) <> vendedor_a then
					qtde_vendedor = qtde_vendedor + 1
					if vendedor <> "" then vendedor = vendedor & ", "
					vendedor = vendedor & r("vendedor")

'					if n_reg_total > 0 then x = x & "</td></tr></table><br />" & chr(13)

					x = x & chr(13) & _
						"<table width='649' style='border: 0' cellpadding='0' cellspacing='0'>" & chr(13) & _
						"	<tr>" & chr(13) & _
						"		<td align='left' valign='bottom' class='MB' style='background:#fff;border-bottom:1.2px solid #000'><span class='N'>" & r("vendedor") & " - " & x_usuario(r("vendedor")) & "</span>" & "</td>" & chr(13) & _
						"	</tr>" & chr(13) & _
						"</table>" & chr(13) & _
						"<br />" & chr(13)
					end if 'if Trim("" & r("vendedor")) <> vendedor_a

			'   MUDOU DE MEIO PAGTO?
				if (Trim("" & r("meio_pagto")) <> meio_pagto_a) Or (Trim("" & r("vendedor")) <> vendedor_a) then
					x = x & "<table width='649' style='border:1.2px solid black;' cellpadding='0' cellspacing='0'>" & chr(13) & _
							"	<tr>" & chr(13) & _
							"		<td colspan='2' align='left' valign='bottom' style='background:white;border-bottom:1.2px solid black;'><span class='N'>&nbsp;" & regex.Replace(r("meio_pagto"), "") & " - " & x_meio_pagto(r("meio_pagto")) & "</span>" & "</td>" & chr(13) & _
							"	</tr>" & chr(13) & _
							"	<tr>" & chr(13) & _
							"		<td width='50%'>" & chr(13)

					inc = 0
					meio_pagto_a = r("meio_pagto")
					end if 'if (Trim("" & r("meio_pagto")) <> meio_pagto_a) Or (Trim("" & r("vendedor")) <> vendedor_a)

				indice = indice + 1
				n_reg = 0

				x = x & _
					"			" & Replace(cab_table, "tableDados", "tableDados_" & idx_bloco)

				x = x & _
					"				<tr>" & chr(13) & _
					"					<td colspan='10' width='50%' align='left' valign='bottom' style='background:white;'>" & chr(13) & _
					"						<table width='100%' cellspacing='0' cellpadding='0'>" & chr(13) & _
					"							<tr class='notPrint'>" & chr(13) & _
					"								<td colspan='3' align='left' valign='bottom' style='height:15px;vertical-align:middle;border-bottom:1px solid #c0c0c0'><span class='Cn'>Indicador: " & r("indicador") & "</span></td>" & chr(13) & _
					"							</tr>" & chr(13) & _
					"							<tr>" & chr(13) & _
					"								<td colspan='3' align='left' valign='bottom' style='vertical-align:middle'><div valign='bottom' style='height:14px;max-height:14px;overflow:hidden;vertical-align:middle'><span class='Cn'>Banco: " & r("banco") & " - " & x_banco(r("banco")) &  "</span></div></td>" & chr(13) & _
					"							</tr>" & chr(13) & _
					"							<tr>" & chr(13) & _
					"								<td class='MTD' align='left' valign='bottom' style='height:15px;vertical-align:middle'><span class='Cn'>Agência: " & r("agencia")

				if Trim("" & r("agencia_dv")) <> "" then
					x = x & "-" & r("agencia_dv")
					end if

				x = x & "</span></td>" & chr(13) & _
					"								<td class='MC' align='left' valign='bottom' style='height:15px;vertical-align:middle'><span class='Cn'>"

				if Trim("" & r("tipo_conta")) <> "" then
					if r("tipo_conta") = "P" then
						x = x & "C/P: "
					elseif r("tipo_conta") = "C" then
						x = x & "C/C: "
						end if
				else
					x = x & "Conta: "
					end if

				if Trim("" & r("conta_operacao")) <> "" then
					x = x & r("conta_operacao") & "-"
					end if

				x = x & r("conta")
	
				if Trim("" & r("conta_dv")) <> "" then
					x = x & "-" & r("conta_dv")
					end if

				x = x & "</span></td>" & chr(13) & _
					"							</tr>" & chr(13) & _
					"							<tr>" & chr(13) & _
					"								<td class='MC' width='60%' colspan='2' align='left' valign='bottom' style='height:15px;vertical-align:middle'><span class='Cn'>Favorecido: " & r("favorecido") & "</span></td>" & chr(13) & _
					"							</tr>" & chr(13)

				if Len(retorna_so_digitos(Trim("" & r("favorecido_cnpj_cpf")))) = 11 then
					s_aux = "CPF"
				else
					s_aux = "CNPJ"
					end if

				x = x & _
					"							<tr>" & chr(13) & _
					"								<td class='MC' width='60%' colspan='2' align='left' valign='bottom' style='height:15px;vertical-align:middle'><span class='Cn'>" & s_aux & ": " & cnpj_cpf_formata(Trim("" & r("favorecido_cnpj_cpf"))) & "</span></td>" & chr(13) & _
					"							</tr>" & chr(13)

				x = x & _
					"						</table>" & chr(13) & _
					"					</td>" & chr(13) & _
					"				</tr>" & chr(13)
				
				comissao = r("vl_total_pagto_liquido_NFS")
				
				x = x & _
					"				<tr nowrap style='background: #FFF'>" & chr(13) & _
					"					<td class='MC' align='right' width='80%' style='height:15px;border-left:0;' nowrap><span class='Cd'>" & SIMBOLO_MONETARIO & " :</span></td>" & chr(13) & _
					"					<td class='MC' align='right' style='height:15px'><span class='Cd'>" & formata_moeda(comissao) & "</span></td>" & chr(13) & _
					"				</tr>" & chr(13) & _ 
					"			</table>" & chr(13)

				end if 'if (Trim("" & r("indicador")) <> indicador_a) Or (Trim("" & r("vendedor")) <> vendedor_a)

			indicador_a = Trim("" & r("indicador"))
			vendedor_a = Trim("" & r("vendedor"))

		  ' CONTAGEM
			n_reg = n_reg + 1
			n_reg_total = n_reg_total + 1

			if (n_reg_total mod 50) = 0 then
				Response.Write x
				x = ""
				end if
			
			end if 'if (r("vl_total_comissao_liquido_arredondado_NFS") + r("vl_total_RA_liquido_arredondado_NFS")) > 0

		r.MoveNext
		loop


  ' MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	if n_reg_total = 0 then
		x = cab_table
		x = x & _
			"	<tr nowrap>" & chr(13) & _
			"		<td class='MT ALERTA' width='649'  colspan='12' align='center'><span class='ALERTA'>&nbsp;NENHUM REGISTRO ENCONTRADO&nbsp;</span></td>" & chr(13) & _
			"	</tr>" & chr(13) & _
			"</table>" & chr(13)
	else
	  ' FECHA TABELA DO ÚLTIMO BANCO
		if (inc mod 2) = 0  then
			x = x & _
				"		</td>" & chr(13) & _
				"		<td width='50%'>" & chr(13) & _
				"		</td>" & chr(13) & _
				"	</tr>" & chr(13) & _
				"</table>" & chr(13)
		else
			x = x & _
				"		</td>" & chr(13) & _
				"	</tr>" & chr(13) & _
				"</table>" & chr(13)
			end if
		end if

	x = x & _
		chr(13) & chr(13) & _
		"<input type='hidden' name='ConsultaVendedor' id='ConsultaVendedor' value='" & vendedor & "' />" & chr(13) & _
		"<input type='hidden' name='filtroMes' id='filtroMes' value='" & c_dt_entregue_mes & "' />" & chr(13) & _
		"<input type='hidden' name='filtroAno' id='filtroAno' value='" & c_dt_entregue_ano & "' />" & chr(13)

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
<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_I18N%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">

$(function () {
    $("#FormFluxoCaixaBase").css('filter', 'alpha(opacity=30)');
    $("#dt_competencia_fluxo_caixa").hUtilUI('datepicker_filtro_inicial');

    var MostraVendedor, mes, ano;
    MostraVendedor = ""; mes = ""; ano = "";
    MostraVendedor = $("#ConsultaVendedor").val();
    mes = $("#filtroMes").val() + "/" + $("#filtroAno").val();
    $("#MostraVendedor").text(MostraVendedor);
    $("#Competencia").text(mes);

    //Every resize of window
    $(window).resize(function () {
        sizeFormFluxoCaixa();
    });

    //Every scroll of window
    $(window).scroll(function () {
        sizeFormFluxoCaixa();
    });

    //Dynamically assign height
    function sizeFormFluxoCaixa() {
        var newTop = $(window).scrollTop() + "px";
        $("#FormFluxoCaixaBase").css("top", newTop);
        $("#FormFluxoCaixa").css("top", newTop);
    }

    if ($(document).height() > $(window).height()) {
        $("#FormFluxoCaixa").css('margin-top', '12%');
    }
    else {
        $("#FormFluxoCaixa").css('margin-top', '-175px');
    }
    
});

function abreFormularioFluxoCaixa() {
    $("#FormFluxoCaixaBase").css('display', 'block');
    $("#FormFluxoCaixa").css('display', 'block');
}

function fechaFormularioFluxoCaixa() {
    $("#FormFluxoCaixaBase").css('display', 'none');
    $("#FormFluxoCaixa").css('display', 'none');
}

function fRELGravaDados(f) {

    if (f.dt_competencia_fluxo_caixa.value == "") {
        alert("Preencha a data de lançamento!");
        f.dt_competencia_fluxo_caixa.focus();
        return;
    }
    if (f.conta_comissao_fluxo_caixa.value == "") {
        alert("Preencha a conta comissão!");
        f.conta_comissao_fluxo_caixa.focus();
        return;
    }
    else {
        if (f.conta_comissao_fluxo_caixa.value != "1400") {
            var r = confirm("Você alterou o plano de contas (comissão)!");
            if (r == false) {
                f.conta_comissao_fluxo_caixa.focus();
                return;
            }
        }
    }

    if (f.conta_RA_fluxo_caixa.value == "") {
        alert("Preencha a conta RA!");
        f.conta_RA_fluxo_caixa.focus();
        return;
    }
    else {
        if (f.conta_RA_fluxo_caixa.value != "1405") {
            var r = confirm("Você alterou o plano de contas (RA)!");
            if (r == false) {
                f.conta_RA_fluxo_caixa.focus();
                return;
            }
        }
    }

	window.status = "Aguarde ...";
	bCONFIRMA.style.visibility = "hidden";
	f.action = "RelComissaoIndicadoresPagDescGravaDados.asp";
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
<link href="<%=URL_FILE__EPRINTER_CSS%>" rel="stylesheet" type="text/css" media="print">
<link href="<%=URL_FILE__JQUERY_UI_CSS%>" rel="stylesheet" type="text/css">

<style type="text/css">
    #FormFluxoCaixaBase {
        position: absolute;
        width: 100%;
        height: 100%;
        top: 0;
        left: 0;
        background-color: #000;
        opacity:0.3;
        display: none;
    }

    #FormFluxoCaixa {
        position: absolute;
        left: 50%;
        top: 50%;
        width: 600px;
        height:350px;
        margin-left: -300px;
        background-color: #fff;
        z-index: 100;
        border:1px solid #000;
        display: none;
        padding: 20px;
    }
    .break{
        break-after:always;
    }
</style>

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
<table cellspacing="0">
<tr>
	<td align="center"><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back()"><img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
</tr>
</table>
</center>
</body>



<% else
     %>
<!-- ***************************************************** -->
<!-- **********  PÁGINA PARA EXIBIR RESULTADO   ********** -->
<!-- ***************************************************** -->
<body onload="window.status='Concluído';">

<center>

<form id="fREL" name="fREL" method="post">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>

<input type="hidden" name="pedido_selecionado" id="pedido_selecionado" value="">
<input type="hidden" name="orcamento_selecionado" id="orcamento_selecionado" value="">
<input type="hidden" name="c_dt_entregue_mes" id="c_dt_entregue_mes" value="<%=c_dt_entregue_mes%>">
<input type="hidden" name="c_dt_entregue_ano" id="c_dt_entregue_ano" value="<%=c_dt_entregue_ano%>">
<input type="hidden" name="rb_visao" id="rb_visao" value="<%=rb_visao%>" />

    <div id="caixa-confirmacao" title="Deseja realmente sair?">
  <span id="msgEtq" style="display:none">Você fez alterações nos dados para etiqueta. Tem certeza que deseja sair sem salvá-las?</span>
</div>

<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="709" cellpadding="4" cellspacing="0" class="notPrint" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Relação de Depósitos Com Desconto</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>


<!-- FILTROS -->
<% 
	s_filtro = "<table width='709' class='notPrint' cellpadding='0' cellspacing='0' style='border-bottom:1px solid black' border='0'>" & chr(13)

'	PERÍODO: MÊS DE COMPETÊNCIA
	s = ""
	if (c_dt_entregue_mes <> "") Or (c_dt_entregue_ano <> "") then
	'	DEVIDO AO WORD WRAP: SÓ FAZ WORD WRAP QUANDO ENCONTRA CHR(32), OU SEJA, MANTÉM AGRUPADO TEXTO COM &nbsp;
		if s <> "" then s = s & ",&nbsp; "
		s_aux = c_dt_entregue_mes
		if s_aux = "" then 
            s_aux = "N.I."
        else
            if c_dt_entregue_mes = "" then 
                s_aux = "N.I."
            else
	    	    s_aux = " " & s_aux & "/"
		        s_aux = replace(s_aux, " ", "&nbsp;")
		        s = s & s_aux
		        s_aux = c_dt_entregue_ano
		        s_aux = replace(s_aux, " ", "&nbsp;")
            end if
        end if
		s = s & s_aux  
		end if


		s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='top' nowrap><span class='N'>Mês de competência:&nbsp;</span></td>" & chr(13) & _
					"		<td align='left' valign='top' width='99%'><span id='Competencia' class='N'>" & s & "</span></td>" & chr(13) & _
					"	</tr>" & chr(13)
		

'	VENDEDOR
		s =  vendedor
		's_aux = x_usuario(vendedor)
		s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='top' nowrap><span class='N'>Vendedor(es):&nbsp;</span></td>" & chr(13) & _
					"		<td align='left'  valign='top' width='99%'><span id='MostraVendedor' class='N'>"&s&"</span></td>" & chr(13) & _
					"	</tr>" & chr(13)



	s_filtro = s_filtro & "</table>" & chr(13)
	Response.Write s_filtro
     
%>
<br />
<!--  RELATÓRIO  -->
<br>
<% consulta_executa %>
<input type="hidden" name="c_id" id="c_id" value="<%=s_id%>" />
<!-- ************   SEPARADOR   ************ -->

<table class="notPrint" width="709" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
	<td class="Rc" align="left">&nbsp;</td>
</tr>
</table>

<br />
<table class="notPrint" width="709" cellspacing="0">
<tr>
	<td align="left"><a name="bVOLTA" id="bVOLTA" href="javascript:history.back()" title="volta para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a>
	</td>
	<td align="left">&nbsp;</td>
	<td align="right">
        <%if aviso="" then %>
		<div name="dMAGICO" id="dMAGICO"><a name="bMAGICO" id="bMAGICO" href="javascript:abreFormularioFluxoCaixa()" title="avançar"><img src="../botao/magico.gif" width="176" height="55" border="0"></a></div>
	    <% else %>
		<div name="dIMPRIME" id="dIMPRIME"><a name="bIMPRIME" id="bIMPRIME" href="javascript:window.print();" title="imprimir relatório"><img src="../botao/imprimir.gif" width="176" height="55" border="0"></a></div>
        <%end if %>
    </td>
</tr>
</table>

<!-- div formulario fluxo caixa -->

<div id="FormFluxoCaixaBase"></div>
    <div id="FormFluxoCaixa">
    <h1>Dados para lançamento</h1>
    <table cellspacing="0" style="width:600px;">
        <tr>
            <td align="right" valign="middle" width="180" style="padding-top:20px">
                <span class="C" style="font-size: 11pt">Data de competência:</span>
            </td>
            <td valign="bottom">
                &nbsp;
                <input type="text" id="dt_competencia_fluxo_caixa" name="dt_competencia_fluxo_caixa" style="width:100px" onblur="if (!isDate(this)) {alert('Data inválida!'); this.focus();}" onkeypress="filtra_data();" />
            </td>
        </tr>
        <tr>
            <td align="right" valign="middle" width="180" style="padding-top:20px">
                <span class="C" style="font-size: 11pt">Conta corrente:</span>
            </td>
            <td valign="bottom">
                &nbsp;
                <select id="conta_corrente_fluxo_caixa" name="conta_corrente_fluxo_caixa" style="width:380px">
                    <%=conta_corrente_monta_itens_select()%>
                </select>
            </td>
        </tr>
        <tr>
            <td align="right" valign="middle" width="180" style="padding-top:20px;">
                <span class="C" style="font-size: 11pt">Empresa:</span>
            </td>
            <td valign="bottom">
                &nbsp;
                <select id="empresa_fluxo_caixa" name="empresa_fluxo_caixa" style="width:380px">
                    <%=empresa_monta_itens_select()%>
                </select>
            </td>
        </tr>
        <tr>
            <td align="right"  width="180" style="padding-top:20px;">
                <span class="C" style="font-size: 11pt;">Plano de contas (comissão):</span>
            </td>
            <td valign="bottom">
                &nbsp;
                <input type="text" id="conta_comissao_fluxo_caixa" name="conta_comissao_fluxo_caixa" style="width:100px" value="1400" />
            </td>
        </tr>
        <tr>
            <td align="right" valign="middle" width="180" style="padding-top:20px">
                <span class="C" style="font-size: 11pt">Plano de contas (RA):</span>
            </td>
            <td valign="bottom">
                &nbsp;
                <input type="text" id="conta_RA_fluxo_caixa" name="conta_RA_fluxo_caixa" style="width:100px" value="1405" />
            </td>
        </tr>
       
    </table>
    <table style="position:relative;bottom:-20px;width:100%">
        <tr>
            <td width="50%" align="right" style="padding:5px">
                <a name="bCANCELA" href="javascript:fechaFormularioFluxoCaixa();" title="cancela">
                <img src="../botao/cancelar.gif" width="176" height="55" border="0"></a>
            </td>
            <td width="50%" align="left" style="padding:5px">
                <a name="bCONFIRMA" href="javascript:fRELGravaDados(fREL)" title="grava os dados">
                <img id="bCONFIRMA" src="../botao/processar.gif" width="176" height="55" border="0"></a>
            </td>
        </tr>
    </table>

</div>

</form>

</center>

</body>

<% end if %>

</html>


<%
	if rs.State <> 0 then rs.Close
	set rs = nothing

'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing

    cn2.Close
    set cn2 = nothing

%>
