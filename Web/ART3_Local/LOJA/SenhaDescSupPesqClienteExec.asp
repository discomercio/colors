<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ======================================================================
'	  S E N H A D E S C S U P P E S Q C L I E N T E E X E C . A S P
'     ======================================================================
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
	dim usuario, loja
	usuario = trim(Session("usuario_atual"))
	loja = Trim(Session("loja_atual"))
	If (usuario = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 
	If (loja = "") then Response.Redirect("aviso.asp?id=" & ERR_SESSAO) 

'	CONECTA COM O BANCO DE DADOS
	Dim cn
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)

'	PARÂMETROS DE PESQUISA
	dim rb_op, c_cnpj_cpf, c_nome_completo, c_nome_parcial
	rb_op = Trim(Request.Form("rb_op"))
	c_cnpj_cpf = retorna_so_digitos(Trim(Request.Form("c_cnpj_cpf")))
	c_nome_completo = Trim(Request.Form("c_nome_completo"))
	c_nome_parcial = Trim(Request.Form("c_nome_parcial"))
	
	dim ordenacao_selecionada
	ordenacao_selecionada=Trim(request("ordenacao_selecionada"))

	dim s, cliente_unico_encontrado





' ________________________________
' E X E C U T A _ C O N S U L T A
'
sub executa_consulta(byref id_registro_unico_encontrado)
dim consulta, s, i, x, cab, s_aux, s_ult_id, qtde_registros
dim r

	id_registro_unico_encontrado = ""
	qtde_registros = 0
	
  ' CABEÇALHO
	cab="<TABLE class='Q' cellSpacing=0 width='640'>" & chr(13)
	cab=cab & "<TR style='background: #FFF0E0'>"
	cab=cab & "<TD width='35' class='MD MB'><P class='R' style='cursor:pointer; margin-right:2pt;' title='clique para ordenar a lista por este campo' onclick=" & chr(34) & "fORDConcluir('1');" & chr(34) & ">CNPJ/CPF</P></TD>"
	cab=cab & "<TD width='200' class='MD MB'><P class='R' style='cursor:pointer; margin-left:2pt' title='clique para ordenar a lista por este campo' onclick=" & chr(34) & "fORDConcluir('2');" & chr(34) & ">NOME</P></TD>"
	cab=cab & "<TD width='250' class='MB'><P class='R' style='cursor:pointer; margin-left:2pt' title='clique para ordenar a lista por este campo' onclick=" & chr(34) & "fORDConcluir('3');" & chr(34) & ">ENDEREÇO</P></TD>"
	cab=cab & "</TR>" & chr(13)

	select case rb_op
		case "1"
			if c_cnpj_cpf <> "" then s = " (cnpj_cpf='" & retorna_so_digitos(c_cnpj_cpf) & "')"
		case "2"
			if c_nome_completo <> "" then s = " (Upper(nome)='" & ucase(c_nome_completo) & "'" & SQL_COLLATE_CASE_ACCENT & ")"
		case "3"
			if c_nome_parcial <> "" then s = " (Upper(nome) LIKE '" & BD_CURINGA_TODOS & ucase(c_nome_parcial) & BD_CURINGA_TODOS & "'" & SQL_COLLATE_CASE_ACCENT & ")"
		case else
			s = ""
		end select
	
	if s <> "" then s = " WHERE " & s
	
	consulta= "SELECT * FROM t_CLIENTE" & s & " ORDER BY "
	select case ordenacao_selecionada
		case "1": consulta = consulta & "cnpj_cpf"
		case "2": consulta = consulta & "nome, cnpj_cpf"
		case "3": consulta = consulta & "endereco, cnpj_cpf"
		case else: consulta = consulta & "cnpj_cpf"
		end select

  ' EXECUTA CONSULTA
	x=cab
	i=0
	
	set r = cn.Execute( consulta )

	while not r.eof 
	  ' CONTAGEM
		i = i + 1
		qtde_registros = qtde_registros + 1

	  ' ALTERNÂNCIA NAS CORES DAS LINHAS
		x=x & "<TR NOWRAP >"

	 '> CNPJ/CPF
		x=x & " <TD class='MDB' NOWRAP valign='top'><P class='C' style='margin-left:2pt'>"
		x=x & "<a href='javascript:fOPConcluir(" & chr(34) & r("id") & chr(34)
		x=x & ")' title='clique para consultar o cadastro deste cliente'>"
		x=x & cnpj_cpf_formata(r("cnpj_cpf")) & "</a></P></TD>"

	 '> NOME
		x=x & " <TD class='MDB' valign='top'><P class='Cn' style='margin-left:2pt'>" 
		x=x & "<a href='javascript:fOPConcluir(" & chr(34) & r("id") & chr(34)
		x=x & ")' title='clique para consultar o cadastro deste cliente'>"
		x=x & Trim("" & r("nome")) & "</a></P></TD>"

	 '> ENDEREÇO
		s=Trim("" & r("endereco"))
		if s<>"" then
			s_aux=Trim("" & r("endereco_numero"))
			if s_aux<>"" then s=s & ", " & s_aux
			s_aux=Trim("" & r("endereco_complemento"))
			if s_aux<>"" then s=s & " " & s_aux
			s_aux=Trim("" & r("bairro"))
			if s_aux<>"" then s=s & " - " & s_aux
			s_aux=Trim("" & r("cidade"))
			if s_aux<>"" then s=s & " - " & s_aux
			s_aux=Trim("" & r("uf"))
			if s_aux<>"" then s=s & " - " & s_aux
			end if

		if s="" then s="&nbsp;"
		x=x & " <TD class='MB' valign='top'><P class='Cn' style='margin-left:2pt'>"
		x=x & "<a href='javascript:fOPConcluir(" & chr(34) & r("id") & chr(34)
		x=x & ")' title='clique para consultar o cadastro deste cliente'>"
		x=x & s & "</a></P></TD>"

		x=x & "</TR>" & chr(13)

		s_ult_id = Trim("" & r("id"))
		
		if (qtde_registros mod 100) = 0 then
			Response.Write x
			x = ""
			end if
		
		r.MoveNext
		wend


  ' SE FOI ENCONTRADO APENAS UM ÚNICO REGISTRO, RETORNA SEU ID
	if i = 1 then id_registro_unico_encontrado = s_ult_id

  ' MOSTRA TOTAL DE CLIENTES
	x=x & "<TR NOWRAP style='background: #FFFFDD'><TD COLSPAN='3' NOWRAP><P class='Cd'>" & "TOTAL:&nbsp;&nbsp;&nbsp;" & cstr(i) & "&nbsp;&nbsp;clientes" & "</P></TD></TR>"

  ' FECHA TABELA
	x=x & "</TABLE>"
	

	Response.Write x

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
	<title>LOJA</title>
	</head>

<script src="<%=URL_FILE__GLOBAL_JS%>" Language="JavaScript" Type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
window.status='Aguarde, executando consulta ...';

function fORDConcluir(ordenacao){
	window.status = "Aguarde ...";
	fORD.ordenacao_selecionada.value=ordenacao;
	fORD.submit(); 
}

function fOPConcluir(s_id){
	window.status = "Aguarde ...";
	fOP.cliente_selecionado.value=s_id;
	fOP.submit(); 
}

</script>

<link href="<%=URL_FILE__E_CSS%>" Rel="stylesheet" Type="text/css">
<link href="<%=URL_FILE__EPRINTER_CSS%>" Rel="stylesheet" Type="text/css" media="print">



<body onload="window.status='Concluído';bVOLTAR.focus();" link=#000000 alink=#000000 vlink=#000000>

<!--  I D E N T I F I C A Ç Ã O  -->
<table width="100%" cellPadding="4" CellSpacing="0" style="border-bottom:1px solid black">
<tr>
	<td align="right" valign="bottom" nowrap><span class="PEDIDO">Resultado da Pesquisa de Clientes</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>


<!--  RELAÇÃO DE CLIENTES  -->
<br>
<center>
<form method="post" action="SenhaDescSupPesqClienteExec.ASP" id="fORD" name="fORD">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="c_nome_parcial" id="c_nome_parcial" value='<%=c_nome_parcial%>'>
<input type="hidden" name="ordenacao_selecionada" id="ordenacao_selecionada" value=''>
</form>

<form method="post" action="SenhaDescSup.asp" id="fOP" name="fOP">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<input type="hidden" name="cliente_selecionado" id="cliente_selecionado" value='<%=cliente_unico_encontrado%>'>
<% executa_consulta cliente_unico_encontrado %>
</form>

<br>

<p class="TracoBottom"></p>

<table class="notPrint" cellSpacing="0">
<tr>
	<td align="CENTER"><a name="bVOLTAR" id="bVOLTAR" href="javascript:history.back();" title="volta para a página anterior">
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

	if cliente_unico_encontrado<>"" then
		Response.Redirect("SenhaDescSup.asp?cliente_selecionado=" & cliente_unico_encontrado & "&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo")))
		end if
%>
