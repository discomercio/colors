<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     =================================================================================
'	  O R C A M E N T I S T A E I N D I C A D O R A S S O C A O V E N D E D O R . A S P
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
	
'	OBTEM USUÁRIO
	dim usuario
	usuario = trim(Session("usuario_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	CONECTA COM O BANCO DE DADOS
	Dim cn, rs, msg_erro
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
    If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

	'Parâmetro 'ord' é a ordenação realizada ao clicar no label de uma coluna da tabela de resultado
	'Parâmetro 'ordenacao' é a ordenação selecionada pelo usuário ao selecionar o tipo de consulta na página de menu do cadastro de orçamentistas/indicadores
	dim ordenacao, ordenacao_selecionada
	ordenacao_selecionada=Trim(Request("ord"))
	ordenacao = Trim("" & Request.Form("ordenacao"))
	if ordenacao_selecionada = "" then ordenacao_selecionada = ordenacao

	'MEMORIZA OPÇÃO DE ORDENAÇÃO FEITA NO MENU DE ORÇAMENTISTAS/INDICADORES
	if ordenacao <> "" then
		call set_default_valor_texto_bd(usuario, "MenuOrcamentistaEIndicador|ordenacao", ordenacao)
		end if

	dim vendedor_selecionado, url_origem
	vendedor_selecionado=Trim(Request("vendedor"))

    url_origem = Request("url_origem")

' ________________________________
' E X E C U T A _ C O N S U L T A
'
Sub executa_consulta
dim consulta, s_where, s, i, x, cab, s_ddd, s_tel, s_telefones, strCidade, strUF
dim iLineNumber
dim r
dim uf_anterior, uf_qtde_total, s_colspan, xRow

  ' CABEÇALHO
	cab="<TABLE class='Q' cellSpacing=0 style='border-top:0;border-right:0;border-bottom:0'>" & chr(13)
	cab=cab & "<TR style='background: #FFF0E0'>"
	cab=cab & "<td align='left' class='MD MB MC' valign='top' style='min-width:30px;'></td>"
	cab=cab & "<TD width='90' nowrap class='MD MB MC' valign='bottom'><P class='R' style='cursor: pointer;' title='clique para ordenar a lista por este campo' onclick='fORDConcluir(" & chr(34) & "1" & chr(34) & ");'" & ">Identificação</P></TD>"
	cab=cab & "<TD width='210' class='MD MB MC' valign='bottom'><P class='R' style='cursor: pointer;' title='clique para ordenar a lista por este campo' onclick='fORDConcluir(" & chr(34) & "2" & chr(34) & ");'" & ">Nome</P></TD>"
	cab=cab & "<TD width='105' class='MD MB MC' valign='bottom'><P class='R' style='cursor: pointer;' title='clique para ordenar a lista por este campo' onclick='fORDConcluir(" & chr(34) & "7" & chr(34) & ");'" & ">CPF/CNPJ</P></TD>"
	cab=cab & "<TD width='95' class='MD MB MC' valign='bottom'><P class='R'>Telefone</P></TD>"
	cab=cab & "<TD width='150' class='MD MB MC' valign='bottom'><P class='R'style='cursor: pointer;' title='clique para ordenar a lista por este campo' onclick='fORDConcluir(" & chr(34) & "4" & chr(34) & ");'" & ">Cidade</P></TD>"
    if ordenacao_selecionada = "UF" then cab=cab & "<td align='right' class='MD MB MC' valign='top' style='min-width:30px;'><P class='R'>Qtde<br />UF</P></td>"
    cab=cab & "<TD width='50' nowrap class='MD MB MC' valign='bottom'><P class='R' style='cursor: pointer;' title='clique para ordenar a lista por este campo' onclick='fORDConcluir(" & chr(34) & "5" & chr(34) & ");'" & ">Acesso Sistema</P></TD>"
	cab=cab & "<TD width='45' nowrap class='MB MD MC' valign='bottom'><P class='R' style='cursor: pointer;' title='clique para ordenar a lista por este campo' onclick='fORDConcluir(" & chr(34) & "6" & chr(34) & ");'" & ">Status</P></TD>"
	cab=cab & "</TR>" & chr(13)

	consulta = "SELECT * FROM t_ORCAMENTISTA_E_INDICADOR"

	s_where = "vendedor = '" & vendedor_selecionado & "'"
		
	if s_where <> "" then s_where = " WHERE " & s_where
	consulta = consulta & s_where
	
	consulta = consulta & " ORDER BY "
	select case ordenacao_selecionada
		case "1": consulta = consulta & "apelido"
		case "2": consulta = consulta & "razao_social_nome, apelido"
		case "3": consulta = consulta & "CONVERT(smallint,loja), apelido"
		case "4": consulta = consulta & "cidade"
		case "5": consulta = consulta & "hab_acesso_sistema"
		case "6": consulta = consulta & "status"
		case "7": consulta = consulta & "LEN(cnpj_cpf), cnpj_cpf"
		case "ID": consulta = consulta & "apelido"
		case "UF": consulta = consulta & "uf, cidade" & SQL_COLLATE_CASE_ACCENT & ", apelido"
		case else: consulta = consulta & "apelido"
		end select

  ' EXECUTA CONSULTA
    x=""
	i=0
	iLineNumber = 0
    uf_qtde_total = 0
    s_colspan = 8
    uf_anterior = "XX"
	
	set r = cn.Execute( consulta )

	while not r.eof 
	  ' CONTAGEM
		i = i + 1
		iLineNumber = iLineNumber + 1
        uf_qtde_total = uf_qtde_total + 1
        xRow = ""

        if Trim("" & r("uf")) <> uf_anterior then
            uf_qtde_total = 1
            end if

	  ' ALTERNÂNCIA NAS CORES DAS LINHAS
		if (i AND 1)=0 then
			xRow=xRow & "<TR style='background: #FFF0E0'>"
		else
			xRow=xRow & "<TR>"
			end if

	 '> Nº LINHA
		xRow=xRow & " <td class='MDB' align='right' valign='top'><span class='Rd' style='margin-right:2px;'>" & CStr(iLineNumber) & ".</span></td>"

	 '> APELIDO
		xRow=xRow & " <TD class='MDB' valign='top'><P class='C'>"
		xRow=xRow & "<a href='javascript:fOPConsultar(" & chr(34) & r("apelido") & chr(34)
		xRow=xRow & ")' title='clique para consultar o cadastro'>"
		xRow=xRow & r("apelido") & "</a></P></TD>"

	 '> NOME & VENDEDORES
		xRow=xRow & " <TD class='MDB' style='width:210px;' valign='top'><P class='C'>" 
		xRow=xRow & "<a href='javascript:fOPConsultar(" & chr(34) & r("apelido") & chr(34)
		xRow=xRow & ")' title='clique para consultar o cadastro'>"
		xRow=xRow & r("razao_social_nome_iniciais_em_maiusculas")
        if Trim("" & r("responsavel_principal"))<>"" then
            xRow=xRow & "<br />Princ: " & Trim("" & r("responsavel_principal"))
            end if
        xRow=xRow & "</a></P>"

        s = "SELECT * FROM t_ORCAMENTISTA_E_INDICADOR_CONTATOS WHERE (indicador = '" & r("apelido") & "') ORDER BY dt_cadastro DESC"
        if rs.State <> 0 then rs.Close
        rs.open s, cn
        if Not rs.Eof then
            xRow=xRow & "<table style='padding: 0px;border:0px;width:100%;'>"
            do while Not rs.Eof
                xRow=xRow & "<tr>"
                xRow=xRow & "<td style='width:60%'><span class='C'>" & Trim("" & rs("nome")) & "</span></td>"
                xRow=xRow & "<td style='width:40%'><span class='C'>" & formata_data(Trim("" & rs("dt_cadastro"))) & "</span></td>"
                xRow=xRow & "</tr>"
                rs.MoveNext
                loop
            xRow=xRow & "</table>"
            end if

        xRow=xRow & "</TD>"

	 '> CPF/CNPJ
		xRow=xRow & " <TD class='MDB' style='width:105px;' valign='top'><P class='C'>" 
		xRow=xRow & "<a href='javascript:fOPConsultar(" & chr(34) & r("apelido") & chr(34)
		xRow=xRow & ")' title='clique para consultar o cadastro'>"
		xRow=xRow & cnpj_cpf_formata(Trim("" & r("cnpj_cpf")))
		xRow=xRow & "</a></P>"
        xRow=xRow & "</TD>"

	 '> TELEFONE
		s_telefones = ""
		s_ddd = Trim("" & r("ddd"))
		if s_ddd <> "" then s_ddd = "(" & s_ddd & ") "
		s_tel = Trim("" & r("telefone"))
		if s_tel <> "" then
			if s_telefones <> "" then s_telefones = s_telefones & "<br>"
			s_tel = telefone_formata(s_tel)
			s_telefones = s_telefones & s_ddd & s_tel
			end if
	'	FAX
		s_tel = Trim("" & r("fax"))
		if s_tel <> "" then
			if s_telefones <> "" then s_telefones = s_telefones & "<br>"
			s_tel = telefone_formata(s_tel)
			s_telefones = s_telefones & s_ddd & s_tel
			end if
	'	CELULAR
		s_ddd = Trim("" & r("ddd_cel"))
		if s_ddd <> "" then s_ddd = "(" & s_ddd & ") "
		s_tel = Trim("" & r("tel_cel"))
		if s_tel <> "" then
			if s_telefones <> "" then s_telefones = s_telefones & "<br>"
			s_tel = telefone_formata(s_tel)
			s_telefones = s_telefones & s_ddd & s_tel
			end if
	'	NEXTEL
		if Trim("" & r("nextel")) <> "" then
			if s_telefones <> "" then s_telefones = s_telefones & "<br>"
			s_telefones = s_telefones & Trim("" & r("nextel"))
			end if
		
		if s_telefones = "" then s_telefones = "&nbsp;"
		xRow=xRow & " <TD class='MDB' style='width:95px;' valign='top'><P class='C'>" & s_telefones & "</P></TD>"

    '>  CIDADE
		strCidade = iniciais_em_maiusculas(Trim("" & r("cidade")))
		strUF = Trim("" & r("uf"))
		if (strCidade <> "") And (strUF <> "") then 
			strCidade = strCidade & " / " & strUF 
		else 
			strCidade = strCidade & strUF
			end if
		if strCidade = "" then strCidade = "&nbsp;"
		xRow = xRow & "		<TD class='MD MB' valign='top' style='width:150px'>" & _
							"<P class='C'>" & strCidade & "</P>" & _
						"</TD>" & chr(13)
     '> QTDE UF
        if ordenacao_selecionada = "UF" then
            xRow=xRow & " <TD class='MDB' style='width:30px;' valign='top' align='right'><P class='Rd contUF'>" & CStr(uf_qtde_total) & "</P></TD>"
            end if

	 '> ACESSO AO SISTEMA
		if r("hab_acesso_sistema") = 1 then 
			s="<span style='color:#006600'>Liberado</span>"
		else 
			s="<span style='color:#ff0000'>Bloqueado</span>"
			end if
		xRow=xRow & " <TD class='MDB' valign='top' NOWRAP><P class='C'>" & s & "</P></TD>"

	 '> STATUS
		if Trim("" & r("status"))="A" then 
			s="<span style='color:#006600'>Ativo</span>"
		else 
			s="<span style='color:#ff0000'>Inativo</span>"
			end if
		xRow=xRow & " <TD class='MB MD' valign='top' NOWRAP><P class='C'>" & s & "</P></TD>"

    '> CONSULTA / EDITA CADASTRO

		xRow=xRow & " <TD valign='middle' NOWRAP style='background-color:#fff;'><a href='javascript:fOPConsultar(""" & r("apelido") & """)'><img src='../imagem/lupa_20x20.png' style='border:0;width:18px;height:18px' class='notPrint' title='Consultar cadastro'></a>"
		xRow=xRow & " <a href='javascript:fOPEditar(""" & r("apelido") & """)'><img src='../imagem/edita_20x20.gif' style='border:0;width:20px;height:20px' class='notPrint' title='Editar cadastro'></a></TD>"
		
        xRow=xRow & "</TR>" & chr(13)
        
        uf_anterior = Trim("" & r("uf"))

		r.MoveNext
        if Not r.Eof then
            if Trim("" & r("uf")) <> uf_anterior then
                xRow = replace(xRow, "contUF", "contUFBold")
                end if
        else
            xRow = replace(xRow, "contUF", "contUFBold")
            end if
            
        x = x & xRow
		wend


  ' MOSTRA TOTAL
    if ordenacao_selecionada = "UF" then s_colspan = s_colspan + 1
	x=x & "<TR NOWRAP style='background: #FFFFDD'><TD COLSPAN='" & CStr(s_colspan) & "' NOWRAP class='MB MD'><P class='Cd'>" & "TOTAL:&nbsp;&nbsp;&nbsp;" & cstr(i) & "</P></TD></TR>"

  ' FECHA TABELA
	x=x & "</TABLE>"
	
    x = cab & x
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

<html>


<head>
	<title>CENTRAL ADMINISTRATIVA</title>
	</head>

<script language="JavaScript" type="text/javascript">
window.status='Aguarde, executando consulta ...';

function fORDConcluir(s_ord){
	window.status = "Aguarde ...";
	fOP.ord.value=s_ord;
	fOP.action="OrcamentistaEIndicadorAssocAoVendedor.asp";
	fOP.submit(); 
}

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

<style type="text/css">
.contUFBold {
	color: #000;
    font-weight: bold;
}
</style>

<body onload="window.status='Concluído';" link=#000000 alink=#000000 vlink=#000000>

<!--  I D E N T I F I C A Ç Ã O  -->
<table width="100%" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom" nowrap><span class="PEDIDO">Orçamentistas / Indicadores Associados ao Vendedor</span>
	<br><span class="C"><%=UCase(vendedor_selecionado)%> - <%=x_usuario(vendedor_selecionado)%></span>
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
<input type="hidden" name="vendedor" id="vendedor" value='<%=vendedor_selecionado%>'>
<input type="hidden" name="ord" id="ord" value=''>
<input type="hidden" name="url_origem" id="url_origem" value='<%=url_origem%>' />
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