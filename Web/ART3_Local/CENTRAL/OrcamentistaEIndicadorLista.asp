<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<%Response.Buffer = False %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =============================================================
'	  O R C A M E N T I S T A E I N D I C A D O R L I S T A . A S P
'     =============================================================
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
	
'	OBTEM USUÁRIO
	dim usuario
	usuario = trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	CONECTA COM O BANCO DE DADOS
	Dim cn
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

	'Parâmetro 'ord' é a ordenação realizada ao clicar no label de uma coluna da tabela de resultado
	'Parâmetro 'ordenacao' é a ordenação selecionada pelo usuário ao selecionar o tipo de consulta na página de menu do cadastro de orçamentistas/indicadores
	dim ordenacao, ordenacao_selecionada
	ordenacao_selecionada=Trim(request("ord"))
	ordenacao = Trim("" & Request.Form("ordenacao"))
	if ordenacao_selecionada = "" then ordenacao_selecionada = ordenacao

	'MEMORIZA OPÇÃO DE ORDENAÇÃO FEITA NO MENU DE ORÇAMENTISTAS/INDICADORES
	if ordenacao <> "" then
		call set_default_valor_texto_bd(usuario, "MenuOrcamentistaEIndicador|ordenacao", ordenacao)
		end if

	dim opcao_consulta
	opcao_consulta=UCase(Trim(request("op")))

	dim filtro_loja
	filtro_loja = Trim(Request("filtro_loja"))



' ________________________________
' E X E C U T A _ C O N S U L T A
'
Sub executa_consulta
dim consulta, s_where, s, i, x, cab, s_op, s_ddd, s_tel, s_telefones, strCidade, strUf
dim iLineNumber
dim r

	s_op = ""
	if opcao_consulta <> "" then s_op = "&op=" & opcao_consulta
	
  ' CABEÇALHO
	cab="<table class='Q' cellspacing=0 style='border-top:0;border-right:0;border-bottom:0'>" & chr(13)
	cab=cab & "<tr style='background: #FFF0E0'>"
	cab=cab & "<td align='left' class='MD MB MC' valign='top' style='min-width:30px;'></td>"
	cab=cab & "<td width='90' class='MD MB MC' align='left' valign='bottom' nowrap><span class='R' style='cursor: pointer;' title='clique para ordenar a lista por este campo' onclick=" & chr(34) & "window.location='OrcamentistaEIndicadorLista.asp?ord=1" & s_op & "&filtro_loja=" & filtro_loja & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")) & "';" & chr(34) & ">Identificação</span></td>"
	cab=cab & "<td width='210' class='MD MB MC' align='left' valign='bottom'><span class='R' style='cursor: pointer;' title='clique para ordenar a lista por este campo' onclick=" & chr(34) & "window.location='OrcamentistaEIndicadorLista.asp?ord=2" & s_op & "&filtro_loja=" & filtro_loja & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")) & "';" & chr(34) & ">Nome</span></td>"
	cab=cab & "<td width='105' class='MD MB MC' align='left' valign='bottom'><span class='R' style='cursor: pointer;' title='clique para ordenar a lista por este campo' onclick=" & chr(34) & "window.location='OrcamentistaEIndicadorLista.asp?ord=8" & s_op & "&filtro_loja=" & filtro_loja & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")) & "';" & chr(34) & ">CPF/CNPJ</span></td>"
	cab=cab & "<td width='95' class='MD MB MC' align='left' valign='bottom'><span class='R'>Telefone</span></td>"
    cab=cab & "<td width='150' class='MD MB MC' valign='bottom'><p class='R'style='cursor: pointer;' title='clique para ordenar a lista por este campo' onclick=" & chr(34) & "window.location='OrcamentistaEIndicadorLista.asp?ord=3" & s_op & "&filtro_loja=" & filtro_loja & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")) & "';" & chr(34) & ");'" & ">Cidade</p></td>"
	cab=cab & "<td width='35' class='MD MB MC' align='right' valign='bottom' nowrap><span class='Rd' style='font-weight:bold; cursor: pointer;' title='clique para ordenar a lista por este campo' onclick=" & chr(34) & "window.location='OrcamentistaEIndicadorLista.asp?ord=4" & s_op & "&filtro_loja=" & filtro_loja & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")) & "';" & chr(34) & ">Loja</span></td>"
	cab=cab & "<td width='90' class='MD MB MC' align='left' valign='bottom' nowrap><span class='R' style='font-weight:bold; cursor: pointer;' title='clique para ordenar a lista por este campo' onclick=" & chr(34) & "window.location='OrcamentistaEIndicadorLista.asp?ord=5" & s_op & "&filtro_loja=" & filtro_loja & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")) & "';" & chr(34) & ">Vendedor</span></td>"
	cab=cab & "<td width='50' class='MD MB MC' align='left' valign='bottom' nowrap><span class='R' style='cursor: pointer;' title='clique para ordenar a lista por este campo' onclick=" & chr(34) & "window.location='OrcamentistaEIndicadorLista.asp?ord=6" & s_op & "&filtro_loja=" & filtro_loja & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")) & "';" & chr(34) & ">Acesso Sistema</span></td>"
	cab=cab & "<td width='45' class='MB MC MD' align='left' valign='bottom' nowrap><span class='R' style='cursor: pointer;' title='clique para ordenar a lista por este campo' onclick=" & chr(34) & "window.location='OrcamentistaEIndicadorLista.asp?ord=7" & s_op & "&filtro_loja=" & filtro_loja & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")) & "';" & chr(34) & ">Status</span></td>"
	cab=cab & "</tr>" & chr(13)

	consulta = "SELECT * FROM t_ORCAMENTISTA_E_INDICADOR"

	s_where = ""
	if opcao_consulta = "A" then
		s_where = "(status = 'A')"
	elseif opcao_consulta = "I" then
		s_where = "(status = 'I')"
		end if
	
	if filtro_loja <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (loja = '" & filtro_loja & "')"
		end if

	if s_where <> "" then s_where = " WHERE " & s_where
	consulta = consulta & s_where
	
	consulta = consulta & " ORDER BY "
	select case ordenacao_selecionada
		case "1": consulta = consulta & "apelido"
		case "2": consulta = consulta & "razao_social_nome, apelido"
		case "3": consulta = consulta & "cidade"
		case "4": consulta = consulta & "CONVERT(smallint,loja), apelido"
		case "5": consulta = consulta & "vendedor, apelido"
		case "6": consulta = consulta & "hab_acesso_sistema"
		case "7": consulta = consulta & "status"
		case "8": consulta = consulta & "LEN(cnpj_cpf), cnpj_cpf"
		case "ID": consulta = consulta & "apelido"
		case "UF": consulta = consulta & "uf, cidade" & SQL_COLLATE_CASE_ACCENT & ", apelido"
		case else: consulta = consulta & "apelido"
		end select

  ' EXECUTA CONSULTA
	x=cab
	i=0
	iLineNumber = 0
	
	set r = cn.Execute( consulta )

	while not r.eof 
	  ' CONTAGEM
		i = i + 1
		iLineNumber = iLineNumber + 1

	  ' ALTERNÂNCIA NAS CORES DAS LINHAS
		if (i AND 1)=0 then
			x=x & "<tr style='background: #FFF0E0'>"
		else
			x=x & "<tr>"
			end if

	 '> Nº LINHA
		x=x & " <td class='MDB' align='right' valign='top'><span class='Rd' style='margin-right:2px;'>" & CStr(iLineNumber) & ".</span></td>"

	 '> APELIDO
		x=x & " <td class='MDB' align='left' valign='top'><span class='C'>"
		x=x & "<a href='javascript:fOPConsultar(" & chr(34) & r("apelido") & chr(34)
		x=x & ")' title='clique para consultar o cadastro'>"
		x=x & r("apelido") & "</a></span></td>"

	 '> NOME
		x=x & " <td class='MDB' style='width:210px;' align='left' valign='top'><span class='C'>" 
		x=x & "<a href='javascript:fOPConsultar(" & chr(34) & r("apelido") & chr(34)
		x=x & ")' title='clique para consultar o cadastro'>"
		x=x & r("razao_social_nome_iniciais_em_maiusculas") & "</a></span></td>"

	 '> CPF/CNPJ
		x=x & " <td class='MDB' style='width:105px;' align='left' valign='top'><span class='C'>" 
		x=x & "<a href='javascript:fOPConsultar(" & chr(34) & r("apelido") & chr(34)
		x=x & ")' title='clique para consultar o cadastro'>"
		x=x & cnpj_cpf_formata(Trim("" & r("cnpj_cpf"))) & "</a></span></td>"

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

    '>  CIDADE
		strCidade = iniciais_em_maiusculas(Trim("" & r("cidade")))
		strUF = Trim("" & r("uf"))
		if (strCidade <> "") And (strUF <> "") then 
			strCidade = strCidade & " / " & strUF 
		else 
			strCidade = strCidade & strUF
			end if
		if strCidade = "" then strCidade = "&nbsp;"
		x = x & "		<td class='MD MB' valign='top' style='width:150px'>" & _
							"<p class='C'>" & strCidade & "</p>" & _
						"</td>" & chr(13)

	 '> LOJA
		s = normaliza_codigo(Trim("" & r("loja")), TAM_MIN_LOJA)
		if s="" then s="&nbsp;"
		x=x & " <td class='MDB' align='right' valign='top' nowrap><span class='Cd'>" & s & "</span></td>"

	 '> VENDEDOR
		s=Trim("" & r("vendedor"))
		if s="" then s="&nbsp;"
		x=x & " <td class='MDB' align='left' valign='top' nowrap><span class='C'>" & s & "</span></td>"

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
		x=x & " <td class='MB MD' align='left' valign='top' nowrap><span class='C'>" & s & "</span></td>"

	 '> CONSULTA / EDITA CADASTRO

		x=x & " <TD valign='middle' NOWRAP style='background-color:#fff;'><a href='javascript:fOPConsultar(""" & r("apelido") & """)'><img src='../imagem/lupa_20x20.png' style='border:0;width:18px;height:18px' title='Consultar cadastro'></a>"
		x=x & " <a href='javascript:fOPEditar(""" & r("apelido") & """)'><img src='../imagem/edita_20x20.gif' style='border:0;width:20px;height:20px' title='Editar cadastro'></a></TD>"
		
        x=x & "</TR>" & chr(13)

		if (i mod 100) = 0 then
			Response.Write x
			x = ""
			end if

		r.MoveNext
		wend


  ' MOSTRA TOTAL
	x=x & "<tr style='background: #FFFFDD' nowrap><td colspan='10' align='right' nowrap class='MB MD'><span class='Cd'>" & "TOTAL:&nbsp;&nbsp;&nbsp;" & cstr(i) & "</span></td></tr>"

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

function fOPConsultar(s_user) {
    window.status = "Aguarde ...";
    fOP.id_selecionado.value = s_user;
    fOP.action = "OrcamentistaEIndicadorConsulta.asp";
    fOP.submit();
}
function fOPEditar(s_user) {
    window.status = "Aguarde ...";
    fOP.id_selecionado.value = s_user;
    fOP.action = "OrcamentistaEIndicadorEdita.asp";
    fOP.submit();
}

</script>

<link href="<%=URL_FILE__E_CSS%>" rel="stylesheet" type="text/css">
<link href="<%=URL_FILE__EPRINTER_CSS%>" rel="stylesheet" type="text/css" media="print">


<body onload="window.status='Concluído';" link=#000000 alink=#000000 vlink=#000000>

<!--  I D E N T I F I C A Ç Ã O  -->  
<table width="100%" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom" nowrap><span class="PEDIDO">Relação de Orçamentistas / Indicadores Cadastrados</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>


<!--  RELAÇÃO DE USUÁRIOS  -->
<br>
<center>
<form method="post" action="OrcamentistaEIndicadorEdita.asp" id="fOP" name="fOP">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name='id_selecionado' id="id_selecionado" value=''>
<input type="hidden" name='operacao_selecionada' id="operacao_selecionada" value='<%=OP_CONSULTA%>'>
<input type="hidden" name="url_origem" id="url_origem" value="MenuOrcamentistaEIndicador.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" />

<% executa_consulta %>
</form>

<br>

<p class="TracoBottom"></p>

<table class="notPrint" cellspacing="0">
<tr>
	<td align="center"><a href="MenuOrcamentistaEIndicador.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="volta para a página anterior">
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