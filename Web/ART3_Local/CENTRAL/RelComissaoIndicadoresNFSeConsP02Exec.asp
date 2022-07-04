<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/Global.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =======================================================
'	  RelComissaoIndicadoresNFSeConsP02Exec.asp
'     =======================================================
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

'	CONECTA AO BANCO DE DADOS
'	=========================
	dim cn
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

	if Not operacao_permitida(OP_CEN_REL_COMISSAO_INDICADORES, s_lista_operacoes_permitidas) then
		Response.Redirect("aviso.asp?id=" & ERR_NIVEL_ACESSO_INSUFICIENTE)
		end if

	dim s, s_aux, s_filtro

	dim alerta
	alerta = ""

	dim dt_mes_competencia, dt_proc_comissao_inicio, dt_proc_comissao_termino
	dim c_competencia_mes, c_competencia_ano, c_dt_proc_comissao_inicio, c_dt_proc_comissao_termino, c_vendedor, c_cnpj_nfse, c_numero_nfse, rb_proc_fluxo_caixa
	c_competencia_mes = Trim(Request.Form("c_competencia_mes"))
	c_competencia_ano = Trim(Request.Form("c_competencia_ano"))
	c_dt_proc_comissao_inicio = Trim(Request.Form("c_dt_proc_comissao_inicio"))
	c_dt_proc_comissao_termino = Trim(Request.Form("c_dt_proc_comissao_termino"))
	c_vendedor = Trim(Request.Form("c_vendedor"))
	c_cnpj_nfse = retorna_so_digitos(Trim(Request.Form("c_cnpj_nfse")))
	c_numero_nfse = retorna_so_digitos(Trim(Request.Form("c_numero_nfse")))
	rb_proc_fluxo_caixa = Trim(Request.Form("rb_proc_fluxo_caixa"))

	dt_mes_competencia = Null
	if (c_competencia_mes <> "") Or (c_competencia_ano <> "") then
		if c_competencia_mes = "" then
			alerta=texto_add_br(alerta)
			alerta=alerta & "Filtro 'Mês de Competência': o mês não foi informado"
			end if
		if c_competencia_ano = "" then
			alerta=texto_add_br(alerta)
			alerta=alerta & "Filtro 'Mês de Competência': o ano não foi informado"
			end if
		if alerta = "" then
			s = "01/" & normaliza_a_esq(c_competencia_mes, 2) & "/" & c_competencia_ano
			dt_mes_competencia = StrToDate(s)
			end if
		end if

	if (c_dt_proc_comissao_inicio <> "") Or (c_dt_proc_comissao_termino <> "") then
		if c_dt_proc_comissao_inicio = "" then
			alerta=texto_add_br(alerta)
			alerta=alerta & "Filtro 'Data Processamento da Comissão': a data de início do período não foi informada"
			end if
		if c_dt_proc_comissao_termino = "" then
			alerta=texto_add_br(alerta)
			alerta=alerta & "Filtro 'Data Processamento da Comissão': a data de término do período não foi informada"
			end if
		end if

	dt_proc_comissao_inicio = Null
	dt_proc_comissao_termino = Null
	if alerta = "" then
		if c_dt_proc_comissao_inicio <> "" then dt_proc_comissao_inicio = StrToDate(c_dt_proc_comissao_inicio)
		if c_dt_proc_comissao_termino <> "" then dt_proc_comissao_termino = StrToDate(c_dt_proc_comissao_termino)
		if dt_proc_comissao_termino < dt_proc_comissao_inicio then
			alerta=texto_add_br(alerta)
			alerta=alerta & "Filtro 'Data Processamento da Comissão': a data de término do período é anterior à data de início"
			end if
		end if





' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________

' _____________________________________
' CONSULTA EXECUTA
'
sub consulta_executa
const VENDA_NORMAL = "VEN"
const DEVOLUCAO = "DEV"
const PERDA = "PER"
dim x, cab_table, cab, s_cor
dim n_reg
dim sql, sql_base, sql_ped_venda, sql_ped_devolucao, sql_ped_perda, sql_vendedor, sql_indicador, s_where, s_link
dim r

	'CABEÇALHO
	cab_table = "<table cellspacing='0' id='tableDados'>" & chr(13)
	cab = "	<tr style='background:azure' nowrap>" & chr(13) & _
		  "		<td class='MT tdNsu' align='right' style='vertical-align:bottom;'><span class='Rd spnTit'>NSU</span></td>" & chr(13) & _
		  "		<td class='MTBD tdDtProcCom' align='center' style='vertical-align:bottom;'><span class='Rc spnTit'>DT Proc Comissão</span></td>" & chr(13) & _
		  "		<td class='MTBD tdUsuProcCom' align='left' style='vertical-align:bottom;'><span class='R spnTit'>Usuário Proc Comissão</span></td>" & chr(13) & _
		  "		<td class='MTBD tdStProcFC' align='center' style='vertical-align:bottom;'><span class='Rc spnTit'>Status Fluxo Caixa</span></td>" & chr(13) & _
		  "		<td class='MTBD tdVendedor' align='left' style='vertical-align:bottom;'><span class='R spnTit'>Vendedor(es)</span></td>" & chr(13) & _
		  "		<td class='MTBD tdIndicador' align='left' style='vertical-align:bottom;'><span class='R spnTit'>Indicador(es)</span></td>" & chr(13) & _
		  "		<td class='MTBD tdVlRT' align='right' style='vertical-align:bottom;'><span class='Rd spnTit'>Comissão</span></td>" & chr(13) & _
		  "		<td class='MTBD tdVlRALiq' align='right' style='vertical-align:bottom;'><span class='Rd spnTit'>RA Líq</span></td>" & chr(13) & _
		  "		<td class='MTBD tdVlTotalRtRaLiq' align='right' style='vertical-align:bottom;'><span class='Rd spnTit'>VL Total (RT + RA Líq)</span></td>" & chr(13) & _
		  "		<td class='MTBD tdPedVenda' align='left' style='vertical-align:bottom;'><span class='R spnTit'>Venda</span></td>" & chr(13) & _
		  "		<td class='MTBD tdPedDevolucao' align='left' style='vertical-align:bottom;'><span class='R spnTit' >Devolução</span></td>" & chr(13) & _
		  "		<td class='MTBD tdPedPerda' align='left' style='vertical-align:bottom;'><span class='R spnTit'>Perda</span></td>" & chr(13) & _
		  "	</tr>" & chr(13)

	s_where = ""

	'FILTRO: MÊS DE COMPETÊNCIA
	if IsDate(dt_mes_competencia) then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (n1Base.competencia_data = " & bd_formata_data(dt_mes_competencia) & ")"
		end if

	'FILTRO: DATA PROCESSAMENTO DA COMISSÃO
	if IsDate(dt_proc_comissao_inicio) And IsDate(dt_proc_comissao_termino) then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " ((n1Base.proc_comissao_data >= " & bd_formata_data(dt_proc_comissao_inicio) & ") AND (n1Base.proc_comissao_data < " & bd_formata_data(dt_proc_comissao_termino+1) & "))"
		end if

	'FILTRO: VENDEDOR
	if c_vendedor <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (n2Base.vendedor = '" & c_vendedor & "')"
		end if

	'FILTRO: CNPJ NFS-e
	if c_cnpj_nfse <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (n1Base.NFSe_cnpj = '" & c_cnpj_nfse & "')"
		end if

	'FILTRO: Nº NFS-e
	if c_numero_nfse <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (n1Base.NFSe_numero = " & c_numero_nfse & ")"
		end if

	'FILTRO: STATUS PROCESSAMENTO FLUXO CAIXA
	if rb_proc_fluxo_caixa <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (n1Base.proc_fluxo_caixa_status = " & rb_proc_fluxo_caixa & ")"
		end if

	sql_base = "SELECT DISTINCT" & _
			" n1Base.id" & _
		" FROM t_COMISSAO_INDICADOR_NFSe_N1 n1Base" & _
			" INNER JOIN t_COMISSAO_INDICADOR_NFSe_N2 n2Base ON (n1Base.id = n2Base.id_comissao_indicador_nfse_n1)" & _
		" WHERE" & _
			" (n1Base.status <> 0)" & _
			" AND" & _
			s_where

	sql_ped_venda = "STUFF((" & _
					"SELECT" & _
						" ', ' + pedido" & _
					" FROM " & _
						"(" & _
							"SELECT DISTINCT" & _
								" n3Aux.pedido" & _
							" FROM t_COMISSAO_INDICADOR_NFSe_N1 n1Aux" & _
								" INNER JOIN t_COMISSAO_INDICADOR_NFSe_N2 n2Aux ON (n1Aux.id = n2Aux.id_comissao_indicador_nfse_n1)" & _
								" INNER JOIN t_COMISSAO_INDICADOR_NFSe_N3_PEDIDO n3Aux ON (n2Aux.id = n3Aux.id_comissao_indicador_nfse_n2)" & _
							" WHERE" & _
								" (n1Aux.id = n1.id)" & _
								" AND (n3Aux.operacao = '" & VENDA_NORMAL & "')" & _
								" AND (n3Aux.st_selecionado = 1)" & _
						") tN3Aux" & _
					" ORDER BY" & _
						" pedido" & _
					" FOR XML PATH('')" & _
				"), 1, 2, '') AS pedidos_venda"

	sql_ped_devolucao = "STUFF((" & _
					"SELECT" & _
						" ', ' + pedido" & _
					" FROM " & _
						"(" & _
							"SELECT DISTINCT" & _
								" n3Aux.pedido" & _
							" FROM t_COMISSAO_INDICADOR_NFSe_N1 n1Aux" & _
								" INNER JOIN t_COMISSAO_INDICADOR_NFSe_N2 n2Aux ON (n1Aux.id = n2Aux.id_comissao_indicador_nfse_n1)" & _
								" INNER JOIN t_COMISSAO_INDICADOR_NFSe_N3_PEDIDO n3Aux ON (n2Aux.id = n3Aux.id_comissao_indicador_nfse_n2)" & _
							" WHERE" & _
								" (n1Aux.id = n1.id)" & _
								" AND (n3Aux.operacao = '" & DEVOLUCAO & "')" & _
								" AND (n3Aux.st_selecionado = 1)" & _
						") tN3Aux" & _
					" ORDER BY" & _
						" pedido" & _
					" FOR XML PATH('')" & _
				"), 1, 2, '') AS pedidos_devolucao"

	sql_ped_perda = "STUFF((" & _
					"SELECT" & _
						" ', ' + pedido" & _
					" FROM " & _
						"(" & _
							"SELECT DISTINCT" & _
								" n3Aux.pedido" & _
							" FROM t_COMISSAO_INDICADOR_NFSe_N1 n1Aux" & _
								" INNER JOIN t_COMISSAO_INDICADOR_NFSe_N2 n2Aux ON (n1Aux.id = n2Aux.id_comissao_indicador_nfse_n1)" & _
								" INNER JOIN t_COMISSAO_INDICADOR_NFSe_N3_PEDIDO n3Aux ON (n2Aux.id = n3Aux.id_comissao_indicador_nfse_n2)" & _
							" WHERE" & _
								" (n1Aux.id = n1.id)" & _
								" AND (n3Aux.operacao = '" & PERDA & "')" & _
								" AND (n3Aux.st_selecionado = 1)" & _
						") tN3Aux" & _
					" ORDER BY" & _
						" pedido" & _
					" FOR XML PATH('')" & _
				"), 1, 2, '') AS pedidos_perda"

	sql_vendedor = "STUFF((" & _
					"SELECT" & _
						" ', ' + vendedor" & _
					" FROM " & _
						"(" & _
							"SELECT DISTINCT" & _
								" n2Aux.vendedor" & _
							" FROM t_COMISSAO_INDICADOR_NFSe_N1 n1Aux" & _
								" INNER JOIN t_COMISSAO_INDICADOR_NFSe_N2 n2Aux ON (n1Aux.id = n2Aux.id_comissao_indicador_nfse_n1)" & _
								" INNER JOIN t_COMISSAO_INDICADOR_NFSe_N3_PEDIDO n3Aux ON (n2Aux.id = n3Aux.id_comissao_indicador_nfse_n2)" & _
							" WHERE" & _
								" (n1Aux.id = n1.id)" & _
								" AND (n3Aux.st_selecionado = 1)" & _
						") tN2Aux" & _
					" ORDER BY" & _
						" vendedor" & _
					" FOR XML PATH('')" & _
				"), 1, 2, '') AS vendedores"

	sql_indicador = "STUFF((" & _
					"SELECT" & _
						" ', ' + indicador" & _
					" FROM " & _
						"(" & _
							"SELECT DISTINCT" & _
								" n2Aux.indicador" & _
							" FROM t_COMISSAO_INDICADOR_NFSe_N1 n1Aux" & _
								" INNER JOIN t_COMISSAO_INDICADOR_NFSe_N2 n2Aux ON (n1Aux.id = n2Aux.id_comissao_indicador_nfse_n1)" & _
								" INNER JOIN t_COMISSAO_INDICADOR_NFSe_N3_PEDIDO n3Aux ON (n2Aux.id = n3Aux.id_comissao_indicador_nfse_n2)" & _
							" WHERE" & _
								" (n1Aux.id = n1.id)" & _
								" AND (n3Aux.st_selecionado = 1)" & _
						") tN2Aux" & _
					" ORDER BY" & _
						" indicador" & _
					" FOR XML PATH('')" & _
				"), 1, 2, '') AS indicadores"

	sql = "SELECT" & _
			" n1.*" & _
			"," & sql_ped_venda & _
			"," & sql_ped_devolucao & _
			"," & sql_ped_perda & _
			"," & sql_vendedor & _
			"," & sql_indicador & _
		" FROM t_COMISSAO_INDICADOR_NFSe_N1 n1" & _
		" WHERE" & _
			" (n1.id IN (" & sql_base & "))" & _
		" ORDER BY" & _
			" n1.id"

	x = cab_table & _
		cab
	n_reg = 0

	set r = cn.Execute(sql)
	do while Not r.Eof
		n_reg = n_reg + 1
		
		x = x & "	<tr nowrap>" & chr(13)

		s_link = "javascript:fConsultar(" & chr(34) & Trim("" & r("id")) & chr(34) & "," & chr(34) & Trim("" & r("NFSe_cnpj")) & chr(34) & ");"

		'NSU
		x = x & "		<td class='MDBE tdNsu' align='right'><a href='" & s_link & "'><span class='Cnd'>" & Trim("" & r("id")) & "</span></a></td>" & chr(13)

		'DATA PROCESSAMENTO COMISSÃO
		x = x & "		<td class='MDB tdDtProcCom' align='center'><a href='" & s_link & "'><span class='Cnc'>" & formata_data_hora(r("proc_comissao_data_hora")) & "</span></a></td>" & chr(13)

		'USUÁRIO PROCESSAMENTO COMISSÃO
		x = x & "		<td class='MDB tdUsuProcCom' align='left'><span class='Cn'>" & Trim("" & r("proc_comissao_usuario")) & "</span></td>" & chr(13)

		'STATUS PROCESSAMENTO FLUXO CAIXA
		if r("proc_fluxo_caixa_status") = 0 then
			s = "Não"
			s_cor = "red"
		else
			s = "Sim"
			s_cor = "green"
			end if
		x = x & "		<td class='MDB tdStProcFC' align='center'><span class='Cnc' style='color:" & s_cor & ";'>" & s & "</span></td>" & chr(13)

		'VENDEDOR
		x = x & "		<td class='MDB tdVendedor' align='left'><span class='Cn spnDados'>" & Replace(Trim("" & r("vendedores")), ", ", ", <br />") & "</span></td>" & chr(13)

		'INDICADOR
		x = x & "		<td class='MDB tdIndicador' align='left'><span class='Cn spnDados'>" & Replace(Trim("" & r("indicadores")), ", ", ", <br />") & "</span></td>" & chr(13)

		'VL Comissão
		x = x & "		<td class='MDB tdVlRT' align='right'><span class='Cnd'>" & formata_moeda(r("vl_total_geral_selecionado_RT")) & "</span></td>" & chr(13)

		'VL RA Líquido
		x = x & "		<td class='MDB tdVlRALiq' align='right'><span class='Cnd'>" & formata_moeda(r("vl_total_geral_selecionado_RA_liquido")) & "</span></td>" & chr(13)

		'VL Total (RT + RA Líquido)
		x = x & "		<td class='MDB tdVlTotalRtRaLiq' align='right'><span class='Cnd'>" & formata_moeda(r("vl_total_geral_selecionado_RT") + r("vl_total_geral_selecionado_RA_liquido")) & "</span></td>" & chr(13)

		'Relação de pedidos de venda processados
		x = x & "		<td class='MDB tdPedVenda' align='left'><span class='Cn'>" & Trim("" & r("pedidos_venda")) & "</span></td>" & chr(13)

		'Relação de pedidos de devolução processados
		x = x & "		<td class='MDB tdPedDevolucao' align='left'><span class='Cn'>" & Trim("" & r("pedidos_devolucao")) & "</span></td>" & chr(13)

		'Relação de pedidos de perda processados
		x = x & "		<td class='MDB tdPedPerda' align='left'><span class='Cn'>" & Trim("" & r("pedidos_perda")) & "</span></td>" & chr(13)

		x = x & "</tr>" & chr(13)

		r.MoveNext
		loop

	if n_reg = 0 then
		x = x & "	<tr nowrap>" & chr(13) & _
				"<td colspan='12' class='MtAlerta' align='center'><span class='MtAlerta'>NENHUM REGISTRO ENCONTRADO</span></td>" & chr(13) & _
				"	</tr>" & chr(13)
		end if

	x = x & "</table>" & chr(13)

	Response.Write x
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

<script src="<%=URL_FILE__JQUERY%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__AJAX_JS%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_I18N%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>

<script type="text/javascript">
	$(function () {
	});
</script>

<script language="JavaScript" type="text/javascript">
	function fVoltar(f) {
		f.action = "RelComissaoIndicadoresNFSeConsP01Filtro.asp";
		f.submit();
	}

	function fConsultar(nsu_N1, cnpj_nfse) {
		fConsulta.id_nsu_N1.value = nsu_N1;
		fConsulta.c_cnpj_nfse.value = cnpj_nfse;
		fConsulta.submit();
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
.tdNsu{
	width:40px;
	text-align:right;
	vertical-align:middle;
	font-weight:bold;
}
.tdDtProcCom{
	width:60px;
	text-align:center;
	vertical-align:middle;
	font-weight:bold;
	word-break:break-word;
}
.tdUsuProcCom{
	width:60px;
	text-align:center;
	vertical-align:middle;
	font-weight:bold;
	word-break:break-word;
}
.tdVendedor{
	width:80px;
	text-align:left;
	vertical-align:middle;
	font-weight:bold;
	word-break:break-word;
}
.tdIndicador{
	width:100px;
	text-align:left;
	vertical-align:middle;
	font-weight:bold;
	word-break:break-word;
}
.tdStProcFC{
	width:40px;
	text-align:center;
	vertical-align:middle;
	font-weight:bold;
	word-break:break-word;
}
.tdVlRT{
	width:70px;
	text-align:right;
	vertical-align:middle;
	font-weight:bold;
	word-break:break-word;
}
.tdVlRALiq{
	width:70px;
	text-align:right;
	vertical-align:middle;
	font-weight:bold;
	word-break:break-word;
}
.tdVlTotalRtRaLiq{
	width:70px;
	text-align:right;
	vertical-align:middle;
	font-weight:bold;
	word-break:break-word;
}
.tdPedVenda{
	width:300px;
	text-align:left;
	vertical-align:middle;
	font-weight:bold;
	word-break:break-word;
}
.tdPedDevolucao{
	width:70px;
	text-align:left;
	vertical-align:middle;
	font-weight:bold;
	word-break:break-word;
}
.tdPedPerda{
	width:70px;
	text-align:left;
	vertical-align:middle;
	font-weight:bold;
	word-break:break-word;
}
.spnTit{
	display:block;
	vertical-align:bottom;
}
. spnDados{
	display:block;
	vertical-align:middle;
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



<% else %>
<body>
<center>
<form id="fFILTRO" name="fFILTRO" method="post" action="RelComissaoIndicadoresNFSeConsP02Exec.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="c_competencia_mes" id="c_competencia_mes" value="<%=c_competencia_mes%>" />
<input type="hidden" name="c_competencia_ano" id="c_competencia_ano" value="<%=c_competencia_ano%>" />
<input type="hidden" name="c_dt_proc_comissao_inicio" id="c_dt_proc_comissao_inicio" value="<%=c_dt_proc_comissao_inicio%>" />
<input type="hidden" name="c_dt_proc_comissao_termino" id="c_dt_proc_comissao_termino" value="<%=c_dt_proc_comissao_termino%>" />
<input type="hidden" name="c_vendedor" id="c_vendedor" value="<%=c_vendedor%>" />
<input type="hidden" name="c_cnpj_nfse" id="c_cnpj_nfse" value="<%=c_cnpj_nfse%>" />
<input type="hidden" name="c_numero_nfse" id="c_numero_nfse" value="<%=c_numero_nfse%>" />
<input type="hidden" name="rb_proc_fluxo_caixa" id="rb_proc_fluxo_caixa" value="<%=rb_proc_fluxo_caixa%>" />


<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="1060" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Relatório de Pedidos Indicadores (via NFS-e) (Consulta)</span>
	<br /><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>

<!-- FILTROS -->
<%
	s_filtro = "<table width='1060' cellpadding='0' cellspacing='0' style='border-bottom:1px solid black' border='0'>" & chr(13)

'	PERÍODO: MÊS DE COMPETÊNCIA
	s = ""
	if (c_competencia_mes <> "") And (c_competencia_ano <> "") then
		s = normaliza_a_esq(c_competencia_mes, 2) & " / " & c_competencia_ano
		end if
	if s = "" then s = "N.I."
	if s <> "" then
		s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='top' nowrap><span class='N'>Mês de Competência:&nbsp;</span></td>" & chr(13) & _
					"		<td align='left' valign='top' width='99%'><span class='N'>" & s & "</span></td>" & chr(13) & _
					"	</tr>" & chr(13)
		end if

'	DATA PROCESSAMENTO DA COMISSÃO
	s = ""
	if (c_dt_proc_comissao_inicio <> "") Or (c_dt_proc_comissao_termino <> "") then
		s_aux = c_dt_proc_comissao_inicio
		if s_aux = "" then s_aux = "N.I."
		s = s_aux & " até "
		s_aux = c_dt_proc_comissao_termino
		if s_aux = "" then s_aux = "N.I."
		s = s & s_aux
		end if
	if s = "" then s = "N.I."
	if s <> "" then
		s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='top' nowrap><span class='N'>Data Processamento da Comissão:&nbsp;</span></td>" & chr(13) & _
					"		<td align='left' valign='top' width='99%'><span class='N'>" & s & "</span></td>" & chr(13) & _
					"	</tr>" & chr(13)
		end if

'	VENDEDOR
	s = c_vendedor
	if s = "" then s = "N.I."
	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' nowrap><span class='N'>Vendedor:&nbsp;</span></td>" & chr(13) & _
				"		<td align='left' valign='top' width='99%'><span class='N'>" & s & "</span></td>" & chr(13) & _
				"	</tr>" & chr(13)

'	CNPJ EMITENTE NFS-e
	if c_cnpj_nfse = "" then
		s = "N.I."
	else
		s = cnpj_cpf_formata(c_cnpj_nfse)
		end if

	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' nowrap><span class='N'>CNPJ NFS-e:&nbsp;</span></td>" & chr(13) & _
				"		<td align='left' valign='top' width='99%'><span class='N'>" & s & "</span></td>" & chr(13) & _
				"	</tr>" & chr(13)

'	Nº NFS-e
	s = c_numero_nfse
	if s = "" then s = "N.I."
	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' nowrap><span class='N'>Nº NFS-e:&nbsp;</span></td>" & chr(13) & _
				"		<td align='left' valign='top' width='99%'><span class='N'>" & s & "</span></td>" & chr(13) & _
				"	</tr>" & chr(13)

'	STATUS PROCESSAMENTO FLUXO CAIXA
	s = ""
	if rb_proc_fluxo_caixa = "" then
		s = "Todos"
	elseif rb_proc_fluxo_caixa = "0" then
		s = "Somente Não Processados"
	elseif rb_proc_fluxo_caixa = "1" then
		s = "Somente Já Processados"
		end if
	if s = "" then s = "N.I."
	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' nowrap><span class='N'>Status Processamento Fluxo Caixa:&nbsp;</span></td>" & chr(13) & _
				"		<td align='left' valign='top' width='99%'><span class='N'>" & s & "</span></td>" & chr(13) & _
				"	</tr>" & chr(13)

'	EMISSÃO
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>Emissão:&nbsp;</span></td><td align='left' valign='top' width='99%'>" & _
			   "<span class='N'>" & formata_data_hora_sem_seg(Now) & "</span></td></tr>" & chr(13)

	s_filtro = s_filtro & "</table>" & chr(13)
	Response.Write s_filtro
%>

<!--  RELATÓRIO  -->
<br />
<% consulta_executa %>


<!-- ************   SEPARADOR   ************ -->
<table width="1060" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr><td class="Rc">&nbsp;</td></tr>
</table>
<br />


<table width="1060" cellSpacing="0">
<tr>
	<td align="center"><a name="bVOLTAR" id="bVOLTAR" href="javascript:fVoltar(fFILTRO);" title="volta para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
</tr>
</table>
</form>

<form id="fConsulta" name="fConsulta" method="post" action="RelComissaoIndicadoresNFSeP06BotaoMagico.asp">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="c_competencia_mes" id="c_competencia_mes" value="<%=c_competencia_mes%>" />
<input type="hidden" name="c_competencia_ano" id="c_competencia_ano" value="<%=c_competencia_ano%>" />
<input type="hidden" name="c_dt_proc_comissao_inicio" id="c_dt_proc_comissao_inicio" value="<%=c_dt_proc_comissao_inicio%>" />
<input type="hidden" name="c_dt_proc_comissao_termino" id="c_dt_proc_comissao_termino" value="<%=c_dt_proc_comissao_termino%>" />
<input type="hidden" name="c_vendedor" id="c_vendedor" value="<%=c_vendedor%>" />
<input type="hidden" name="c_cnpj_nfse" id="c_cnpj_nfse" value="<%=c_cnpj_nfse%>" />
<input type="hidden" name="rb_proc_fluxo_caixa" id="rb_proc_fluxo_caixa" value="<%=rb_proc_fluxo_caixa%>" />
<input type="hidden" name="origem" id="origem" value="QUERY" />
<input type="hidden" name="id_nsu_N1" id="id_nsu_N1" />
<input type="hidden" name="c_cnpj_nfse" id="c_cnpj_nfse" />
</form>
</center>
</body>

<% end if %>

</html>

<%
'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>
