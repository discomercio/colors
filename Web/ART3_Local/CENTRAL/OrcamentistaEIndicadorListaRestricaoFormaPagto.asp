<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =================================================================================
'	  OrcamentistaEIndicadorListaRestricaoFormaPagto.asp
'     =================================================================================
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


' _____________________________________________________________________________________________
'
'			I N I C I A L I Z A     P Á G I N A     A S P    N O    S E R V I D O R
' _____________________________________________________________________________________________

	On Error GoTo 0
	Err.Clear
	
	Server.ScriptTimeout = MAX_SERVER_SCRIPT_TIMEOUT_EM_SEG
	
	Const COD_ID_FORMA_PAGTO_SEM_RESTRICOES = 999
	
	class cl_RESTRICAO_FORMA_PAGTO
		dim strIdFormaPagto
		dim blnRestricaoAtiva
		end class
	
'	OBTEM USUÁRIO
	dim usuario
	usuario = trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	CONECTA COM O BANCO DE DADOS
	Dim cn
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	dim ordenacao_selecionada
	ordenacao_selecionada=Trim(Request("ord"))

	dim i, s
	dim ckb_somente_ativos, c_lista_id_forma_pagto, v_lista_id_forma_pagto, v_FP_PF, v_FP_PJ, ckb_value
	ckb_somente_ativos = Trim(Request.Form("ckb_somente_ativos"))
	c_lista_id_forma_pagto = Trim(Request.Form("c_lista_id_forma_pagto"))
	v_lista_id_forma_pagto = Split(c_lista_id_forma_pagto, "|")

'	PF
	redim v_FP_PF(0)
	set v_FP_PF(UBound(v_FP_PF)) = New cl_RESTRICAO_FORMA_PAGTO
	v_FP_PF(UBound(v_FP_PF)).strIdFormaPagto = ""
'	PJ
	redim v_FP_PJ(0)
	set v_FP_PJ(UBound(v_FP_PJ)) = New cl_RESTRICAO_FORMA_PAGTO
	v_FP_PJ(UBound(v_FP_PJ)).strIdFormaPagto = ""

'	LAÇO P/ LEITURA DOS CAMPOS
	for i=LBound(v_lista_id_forma_pagto) to UBound(v_lista_id_forma_pagto)
		if Trim(v_lista_id_forma_pagto(i)) <> "" then
		'	PF
			if v_FP_PF(UBound(v_FP_PF)).strIdFormaPagto <> "" then
				redim preserve v_FP_PF(UBound(v_FP_PF)+1)
				set v_FP_PF(UBound(v_FP_PF)) = New cl_RESTRICAO_FORMA_PAGTO
				v_FP_PF(UBound(v_FP_PF)).strIdFormaPagto = ""
				end if
			s = "ckb_" & ID_PF & "_" & Trim(v_lista_id_forma_pagto(i))
			ckb_value = Trim(Request.Form(s))
			v_FP_PF(UBound(v_FP_PF)).strIdFormaPagto = Trim(v_lista_id_forma_pagto(i))
			if ckb_value <> "" then
			'	CHECKBOX MARCADO: RESTRIÇÃO DA FORMA DE PAGTO ESTÁ ATIVA
				v_FP_PF(UBound(v_FP_PF)).blnRestricaoAtiva = True
			else
				v_FP_PF(UBound(v_FP_PF)).blnRestricaoAtiva = False
				end if
		'	PJ
			if v_FP_PJ(UBound(v_FP_PJ)).strIdFormaPagto <> "" then
				redim preserve v_FP_PJ(UBound(v_FP_PJ)+1)
				set v_FP_PJ(UBound(v_FP_PJ)) = New cl_RESTRICAO_FORMA_PAGTO
				v_FP_PJ(UBound(v_FP_PJ)).strIdFormaPagto = ""
				end if
			s = "ckb_" & ID_PJ & "_" & Trim(v_lista_id_forma_pagto(i))
			ckb_value = Trim(Request.Form(s))
			v_FP_PJ(UBound(v_FP_PJ)).strIdFormaPagto = Trim(v_lista_id_forma_pagto(i))
			if ckb_value <> "" then
			'	CHECKBOX MARCADO: RESTRIÇÃO DA FORMA DE PAGTO ESTÁ ATIVA
				v_FP_PJ(UBound(v_FP_PJ)).blnRestricaoAtiva = True
			else
				v_FP_PJ(UBound(v_FP_PJ)).blnRestricaoAtiva = False
				end if
			end if
		next






' ________________________________
' E X E C U T A _ C O N S U L T A
'
Sub executa_consulta
dim consulta, s_where, s_sql_FP, s_where_FP, s, i, x, cab, s_ddd, s_tel, s_telefones
dim r

  ' CABEÇALHO
	cab="<table class='Q' cellspacing=0>" & chr(13)
	cab=cab & "<tr style='background: #FFF0E0'>"
	cab=cab & "<td width='90' class='MD MB' align='left' valign='bottom' nowrap><span class='R' style='cursor: pointer;' title='clique para ordenar a lista por este campo' onclick='fORDConcluir(" & chr(34) & "1" & chr(34) & ");'" & ">Identificação</span></td>"
	cab=cab & "<td width='210' class='MD MB' align='left' valign='bottom'><span class='R' style='cursor: pointer;' title='clique para ordenar a lista por este campo' onclick='fORDConcluir(" & chr(34) & "2" & chr(34) & ");'" & ">Nome</span></td>"
	cab=cab & "<td width='95' class='MD MB' align='left' valign='bottom'><span class='R'>Telefone</span></td>"
	cab=cab & "<td width='35' class='MD MB' align='right' valign='bottom' nowrap><span class='Rd' style='font-weight:bold; cursor: pointer;' title='clique para ordenar a lista por este campo' onclick='fORDConcluir(" & chr(34) & "3" & chr(34) & ");'" & ">Loja</span></td>"
	cab=cab & "<td width='50' class='MD MB' align='left' valign='bottom' nowrap><span class='R' style='cursor: pointer;' title='clique para ordenar a lista por este campo' onclick='fORDConcluir(" & chr(34) & "4" & chr(34) & ");'" & ">Acesso Sistema</span></td>"
	cab=cab & "<td width='45' class='MB' align='left' valign='bottom' nowrap><span class='R' style='cursor: pointer;' title='clique para ordenar a lista por este campo' onclick='fORDConcluir(" & chr(34) & "5" & chr(34) & ");'" & ">Status</span></td>"
	cab=cab & "</tr>" & chr(13)
	
	s_where = ""
	
	if ckb_somente_ativos <> "" then
		if s_where <> "" then s_where = s_where & " AND "
		s_where = s_where & "(tOI.status = 'A')"
		end if
	
	s_sql_FP = "SELECT DISTINCT" & _
					" id_orcamentista_e_indicador" & _
				" FROM t_ORCAMENTISTA_E_INDICADOR_RESTRICAO_FORMA_PAGTO tOIRFP"
	
	s_where_FP = " WHERE" & _
					" (tOIRFP.id_orcamentista_e_indicador = tOI.apelido)" & _
					" AND (st_restricao_ativa <> 0)"
	
	for i=LBound(v_FP_PF) to UBound(v_FP_PF)
		if (Trim(v_FP_PF(i).strIdFormaPagto) <> "") And (Trim(v_FP_PF(i).strIdFormaPagto) <> Cstr(COD_ID_FORMA_PAGTO_SEM_RESTRICOES)) then
			if v_FP_PF(i).blnRestricaoAtiva then
				if s_where <> "" then s_where = s_where & " AND "
				s_where = s_where & "(tOI.apelido IN (" & _
							s_sql_FP & _
							s_where_FP & _
							" AND (tOIRFP.tipo_cliente = '" & ID_PF & "')" & _
							" AND (tOIRFP.id_forma_pagto = " & Trim(v_FP_PF(i).strIdFormaPagto) & ")" & _
							"))"
				end if
		elseif Trim(v_FP_PF(i).strIdFormaPagto) = Cstr(COD_ID_FORMA_PAGTO_SEM_RESTRICOES) then
			if v_FP_PF(i).blnRestricaoAtiva then
				if s_where <> "" then s_where = s_where & " AND "
				s_where = s_where & _
							"(" & _
								"0 = Coalesce(" & _
									"(" & _
									"SELECT Count(*) FROM (" &_ 
									s_sql_FP & _
									s_where_FP & _
									" AND (tOIRFP.tipo_cliente = '" & ID_PF & "')" & _
									") t" & _
								"), -1)" & _
							")"
				end if
			end if
		next
	
	for i=LBound(v_FP_PJ) to UBound(v_FP_PJ)
		if (Trim(v_FP_PJ(i).strIdFormaPagto) <> "") And (Trim(v_FP_PJ(i).strIdFormaPagto) <> Cstr(COD_ID_FORMA_PAGTO_SEM_RESTRICOES)) then
			if v_FP_PJ(i).blnRestricaoAtiva then
				if s_where <> "" then s_where = s_where & " AND "
				s_where = s_where & "(tOI.apelido IN (" & _
							s_sql_FP & _
							s_where_FP & _
							" AND (tOIRFP.tipo_cliente = '" & ID_PJ & "')" & _
							" AND (tOIRFP.id_forma_pagto = " & Trim(v_FP_PJ(i).strIdFormaPagto) & ")" & _
							"))"
				end if
		elseif Trim(v_FP_PJ(i).strIdFormaPagto) = Cstr(COD_ID_FORMA_PAGTO_SEM_RESTRICOES) then
			if v_FP_PJ(i).blnRestricaoAtiva then
				if s_where <> "" then s_where = s_where & " AND "
				s_where = s_where & _
							"(" & _
								"0 = Coalesce(" & _
									"(" & _
									"SELECT Count(*) FROM (" & _
									s_sql_FP & _
									s_where_FP & _
									" AND (tOIRFP.tipo_cliente = '" & ID_PJ & "')" & _
									") t" & _
								"), -1)" & _
							")"
				end if
			end if
		next
	
	consulta = "SELECT " & _
					"*" & _
				" FROM t_ORCAMENTISTA_E_INDICADOR tOI"
	
	if s_where <> "" then s_where = " WHERE " & s_where
	consulta = consulta & s_where
	
	consulta = consulta & " ORDER BY "
	select case ordenacao_selecionada
		case "1": consulta = consulta & "apelido"
		case "2": consulta = consulta & "razao_social_nome, apelido"
		case "3": consulta = consulta & "CONVERT(smallint,loja), apelido"
		case "4": consulta = consulta & "hab_acesso_sistema"
		case "5": consulta = consulta & "status"
		case else: consulta = consulta & "apelido"
		end select

  ' EXECUTA CONSULTA
	x=cab
	i=0
	
	set r = cn.Execute( consulta )

	while not r.eof
	  ' CONTAGEM
		i = i + 1

	  ' ALTERNÂNCIA NAS CORES DAS LINHAS
		if (i AND 1)=0 then
			x=x & "<tr style='background: #FFF0E0'>"
		else
			x=x & "<tr>"
			end if

	 '> APELIDO
		x=x & " <td class='MDB' align='left' valign='top'><span class='C'>"
		x=x & "<a href='javascript:fCADConcluir(" & chr(34) & r("apelido") & chr(34)
		x=x & ")' title='clique para consultar o cadastro'>"
		x=x & r("apelido") & "</a></span></td>"

	 '> NOME
		x=x & " <td class='MDB' style='width:210px;' align='left' valign='top'><span class='C'>" 
		x=x & "<a href='javascript:fCADConcluir(" & chr(34) & r("apelido") & chr(34)
		x=x & ")' title='clique para consultar o cadastro'>"
		x=x & r("razao_social_nome_iniciais_em_maiusculas") & "</a></span></td>"

	 '> TELEFONE
		s_telefones = ""
		s_ddd = Trim("" & r("ddd"))
		if s_ddd <> "" then s_ddd = "(" & s_ddd & ") "
		s_tel = Trim("" & r("telefone"))
		if s_tel <> "" then
			if s_telefones <> "" then s_telefones = s_telefones & "<br>"
			s_tel = telefone_formata(s_tel)
			s_telefones = s_telefones & "<span class='C'>" & s_ddd & s_tel & "</span>"
			end if
	'	FAX
		s_tel = Trim("" & r("fax"))
		if s_tel <> "" then
			if s_telefones <> "" then s_telefones = s_telefones & "<br>"
			s_tel = telefone_formata(s_tel)
			s_telefones = s_telefones & "<span class='C'>" & s_ddd & s_tel & "</span>"
			end if
	'	CELULAR
		s_ddd = Trim("" & r("ddd_cel"))
		if s_ddd <> "" then s_ddd = "(" & s_ddd & ") "
		s_tel = Trim("" & r("tel_cel"))
		if s_tel <> "" then
			if s_telefones <> "" then s_telefones = s_telefones & "<br>"
			s_tel = telefone_formata(s_tel)
			s_telefones = s_telefones & "<span class='C'>" & s_ddd & s_tel & "</span>"
			end if
	'	NEXTEL
		if Trim("" & r("nextel")) <> "" then
			if s_telefones <> "" then s_telefones = s_telefones & "<br>"
			s_telefones = s_telefones & "<span class='C'>" & Trim("" & r("nextel")) & "</span>"
			end if
		
		if s_telefones = "" then s_telefones = "<span class='C'>&nbsp;</span>"
		x=x & " <td class='MDB' style='width:95px;' align='left' valign='top'>" & s_telefones & "</td>"

	 '> LOJA
		s = normaliza_codigo(Trim("" & r("loja")), TAM_MIN_LOJA)
		if s="" then s="&nbsp;"
		x=x & " <td class='MDB' align='right' valign='top' nowrap><span class='Cd'>" & s & "</span></td>"

	 '> ACESSO AO SISTEMA
		if r("hab_acesso_sistema") = 1 then 
 			s="<span style='color:#006600'>Liberado</span>"
 		else 
 			s="<span style='color:#ff0000'>Bloqueado</span>"
			end if
		x=x & " <td class='MDB' align='left' valign='top' nowrap><span class='C'>" & s & "</span></td>"

	 '> STATUS
		if Trim("" & r("status"))="A" then 
			s="<span style='color:#006600'>Ativo</span>"
		else 
			s="<span style='color:#ff0000'>Inativo</span>"
			end if
		x=x & " <td class='MB' align='left' valign='top' nowrap><span class='C'>" & s & "</span></td>"

		x=x & "</tr>"

		if (i mod 100) = 0 then
			Response.Write x
			x = ""
			end if

		r.MoveNext
		wend


  ' MOSTRA TOTAL
	x=x & "<tr style='background: #FFFFDD' nowrap><td colspan='6' align='left' nowrap><span class='Cd'>" & "TOTAL:&nbsp;&nbsp;&nbsp;" & cstr(i) & "</span></td></tr>"

  ' FECHA TABELA
	x=x & "</table>"
	

	Response.write x

	r.close
	set r=nothing

End sub

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
	<title>CENTRAL ADMINISTRATIVA</title>
	</head>

<script language="JavaScript" type="text/javascript">
window.status='Aguarde, executando consulta ...';

function fORDConcluir(s_ord){
	window.status = "Aguarde ...";
	fOP.ord.value=s_ord;
	fOP.action = "OrcamentistaEIndicadorListaRestricaoFormaPagto.asp";
	fOP.submit(); 
}

function fCADConcluir(s_user){
	window.status = "Aguarde ...";
	fOP.id_selecionado.value=s_user;
	fOP.action="OrcamentistaEIndicadorEdita.asp";
	fOP.submit(); 
}

</script>

<link href="<%=URL_FILE__E_CSS%>" rel="stylesheet" type="text/css">
<link href="<%=URL_FILE__EPRINTER_CSS%>" rel="stylesheet" type="text/css" media="print">


<body onload="window.status='Concluído';" link=#000000 alink=#000000 vlink=#000000>

<!--  I D E N T I F I C A Ç Ã O  -->
<table width="100%" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom" nowrap><span class="PEDIDO">Orçamentistas / Indicadores</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>


<!--  RELAÇÃO DE ORÇAMENTISTAS/INDICADORES  -->
<br>
<center>
<form method="post" id="fOP" name="fOP">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="id_selecionado" id="id_selecionado" value=''>
<input type="hidden" name="operacao_selecionada" id="operacao_selecionada" value='<%=OP_CONSULTA%>'>
<input type="hidden" name="ord" id="ord" value=''>
<% executa_consulta %>
</form>

<br>

<p class="TracoBottom"></p>

<table class="notPrint" cellspacing="0">
<tr>
	<td align="center"><a href="javascript:history.back();" title="volta para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a></td>
</tr>
</table>

</center>


</body>
</html>


<%
'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>