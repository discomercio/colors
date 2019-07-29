<%@ Language=VBScript %>
<%OPTION EXPLICIT%>
<% Response.Buffer=True %>
<!-- #include file = "../global/constantes.asp" -->
<!-- #include file = "../global/funcoes.asp"    -->
<!-- #include file = "../global/bdd.asp" -->
<!-- #include file = "../global/Global.asp" -->

<!-- #include file = "../global/TrataSessaoExpirada.asp"        -->

<%
'     ========================================================
'	  RelClientesNegativadosExec.asp
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

'	CONECTA COM O BANCO DE DADOS
	dim cn, rs, msg_erro
	If Not bdd_conecta(cn) then Response.Redirect("aviso.asp?id=" & ERR_CONEXAO)
	If Not cria_recordset_otimista(rs, msg_erro) then Response.Redirect("aviso.asp?id=" & ERR_FALHA_OPERACAO_CRIAR_ADO)

	dim s_lista_operacoes_permitidas
	s_lista_operacoes_permitidas = Trim(Session("lista_operacoes_permitidas"))

'	VERIFICA PERMISSÃO DE ACESSO DO USUÁRIO
	if Not operacao_permitida(OP_CEN_REL_CLIENTE_SPC, s_lista_operacoes_permitidas) then
		cn.Close
		Response.Redirect("aviso.asp?id=" & ERR_ACESSO_INSUFICIENTE)
		end if

	dim url_back, strUrlBotaoVoltar
	url_back = Trim(Request("url_back"))
	if url_back <> "" then
		strUrlBotaoVoltar = "RelClientesNegativadosFiltro.asp?url_back=X&" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))
	else
		strUrlBotaoVoltar = "javascript:history.back()"
		end if

	dim s_filtro, intQtdeClientes
	intQtdeClientes = 0

	dim alerta
	dim idoc, s, s_aux
	dim lista_cnpj_cpf, c_cliente_cnpj_cpf, rb_negativado, rb_ordenacao_saida, v_cnpj_cpf, v_cnpj_cpf_aux
	dim s_lista_cnpj_cpf, s_lista_cnpj_cpf_inexistentes
	dim s_nome_cliente, c_uf

	alerta = ""

	c_uf = Trim(Request("c_uf"))
	lista_cnpj_cpf = Trim(Request("c_cliente_cnpj_cpf"))
	c_cliente_cnpj_cpf = Trim(Request("c_cliente_cnpj_cpf"))
	rb_negativado = Trim(Request("rb_negativado"))
	rb_ordenacao_saida = Trim(Request("rb_ordenacao_saida"))
	
	lista_cnpj_cpf = Trim(request("c_cliente_cnpj_cpf"))
	
	lista_cnpj_cpf=substitui_caracteres(lista_cnpj_cpf,chr(10),"")
	v_cnpj_cpf = split(lista_cnpj_cpf,chr(13),-1)
	s_lista_cnpj_cpf = ""
	for idoc=Lbound(v_cnpj_cpf) to Ubound(v_cnpj_cpf)
		if Trim(v_cnpj_cpf(idoc))<>"" then
			s = retorna_so_digitos(v_cnpj_cpf(idoc))
			if s <> "" then 
				v_cnpj_cpf(idoc) = s
				if s_lista_cnpj_cpf <> "" then s_lista_cnpj_cpf = s_lista_cnpj_cpf & ", "
				s_lista_cnpj_cpf = s_lista_cnpj_cpf & cnpj_cpf_formata(s)
				end if
			end if
		next
	'guardando este vetor para testar CNPJ/CPFs inexistentes
	v_cnpj_cpf_aux = v_cnpj_cpf

	dim qtde_clientes
	qtde_clientes = 0





' _____________________________________________________________________________________________
'
'									F  U  N  Ç  Õ  E  S 
' _____________________________________________________________________________________________


' _____________________________________
' CONSULTA EXECUTA
'
sub consulta_executa
dim s, s_sql, x
dim s_where, s_where_aux
dim s_color
dim r
dim cab_table, cab
dim vl_total_geral
dim cont

'	MONTAGEM DAS RESTRIÇÕES
	s_where = ""
	
	if rb_negativado <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (spc_negativado_status = '" & rb_negativado & "')"
		end if
	
	if c_uf <> "" then
		if s_where <> "" then s_where = s_where & " AND"
		s_where = s_where & " (uf = '" & c_uf & "')"
		end if

	if lista_cnpj_cpf <> "" then
		s_where_aux = ""
		for cont = LBound(v_cnpj_cpf) to UBound(v_cnpj_cpf)
			if v_cnpj_cpf(cont) <> "" then
				if s_where_aux <> "" then s_where_aux = s_where_aux & " OR"
				s_where_aux = s_where_aux & " (cnpj_cpf = '" & v_cnpj_cpf(cont) & "')"
				end if
			next
		if s_where_aux <> "" then 
			if s_where <> "" then s_where = s_where & " AND"
			s_where = s_where & " (" & s_where_aux & ")"
			end if
		end if
	
'	MONTAGEM DA CONSULTA
	s_sql = ""
	
	if s_where <> "" then s_where = " WHERE " & s_where
	s_sql = "SELECT " & _
				"*" & _
			" FROM t_CLIENTE" & _
			s_where & _
			" ORDER BY"
	
	if rb_ordenacao_saida = "ORD_POR_CNPJ" then
		s_sql = s_sql & _
					" cnpj_cpf, nome"
	else
		s_sql = s_sql & _
					" nome, cnpj_cpf"
		end if
	
	if s_sql = "" then
		Response.Write "Falha ao elaborar a consulta SQL: a consulta não possui conteúdo!"
		Response.End
		end if

	cab_table = "<table cellspacing='0' cellpadding='0'>" & chr(13)
	cab = "	<tr style='background:#F0FFFF;' nowrap>" & chr(13) & _
		"		<td class='MDTE tdDocumento' style='vertical-align:bottom'><span class='R'>CNPJ/CPF</span></td>" & chr(13) & _
		"		<td class='MTD tdCliente' style='vertical-align:bottom'><span class='R'>Cliente</span></td>" & chr(13) & _
		"		<td class='MTD tdUF' style='vertical-align:bottom'><span class='R'>UF</span></td>" & chr(13) & _
		"		<td class='MTD tdNegativado' style='vertical-align:bottom'><span class='Rc'>Negativado</span></td>" & chr(13) & _
		"	</tr>" & chr(13)
	
	x = cab_table & cab
	intQtdeClientes = 0
	vl_total_geral = 0
	
	set r = cn.execute(s_sql)
	do while Not r.Eof
	
	 ' CONTAGEM
		intQtdeClientes = intQtdeClientes + 1
		
		x = x & "	<tr nowrap>" & chr(13)

	'> CNPJ/CPF
		s = cnpj_cpf_formata(Trim("" & r("cnpj_cpf")))
		if s = "" then s = "&nbsp;"
		x = x & "		<td class='MDTE tdDocumento'><span class='Cn'>" & s & "</span></td>" & chr(13)

	'> CLIENTE
		s = Trim("" & r("nome"))
		if s = "" then s = "&nbsp;"
		x = x & "		<td class='MTD tdCliente'><span class='Cn'>" & s & "</span></td>" & chr(13)

	'> UF
		s = Trim("" & r("uf"))
		if s = "" then s = "&nbsp;"
		x = x & "		<td class='MTD tdUF'><span class='Cn'>" & s & "</span></td>" & chr(13)
		
	'> NEGATIVADO
		x = x & _
				"		<td class='MTD tdNegativado'>" & chr(13) & _
				"			<input type='checkbox' name='ckb_negativado' id='ckb_negativado' class='CheckNegativado'" & _
								" value='" & Trim("" & r("id")) & "|" & Trim("" & r("nome")) & "|" & Trim("" & cnpj_cpf_formata(Trim("" & r("cnpj_cpf")))) & "'" & _
								iif(r("spc_negativado_status") = 1, " checked", "") & _
								">" & chr(13) & _
						"</td>" & chr(13)

	'> COLUNA OCULTA PARA COMPARAÇÃO POSTERIOR
	'guardaremos nesta coluna a mesma informação concatenada de: id do cliente + nome + documento;
	'adicionalmente, concatenaremos a string "true" ou "false", indicando se o checkbox estava marcado ou desmarcado
	'no carregamento da tela
		x = x & _
				"		<td>" & chr(13) & _
				"			<input type='hidden' name='ckb_inicial' id='ckb_inicial'" & _
								" value='" & iif(r("spc_negativado_status") = 1, "true", "false") & _
								"|" & Trim("" & r("id")) & "|" & Trim("" & r("nome")) & "|" & Trim("" & cnpj_cpf_formata(Trim("" & r("cnpj_cpf")))) & "'" & _
								"'>" & chr(13) & _
						"</td>" & chr(13)

		x = x & "	</tr>" & chr(13)

		if (intQtdeClientes mod 100) = 0 then
			Response.Write x
			x = ""
			end if

		'tirar este CNPJ/CPF do vetor auxiliar
		cont = LBound(v_cnpj_cpf_aux) 
		s = Trim("" & r("cnpj_cpf"))
		Do While (cont<=UBound(v_cnpj_cpf_aux)) and (s <> "")
			if v_cnpj_cpf_aux(cont) = s then
				v_cnpj_cpf_aux(cont) = ""
				s = ""
				end if
			cont = cont + 1
			Loop
		
		r.MoveNext
		loop
	
	
'	TOTAL GERAL
	if intQtdeClientes > 0 then
		x = x & "	<tr>" & chr(13) & _
				"		<td colspan='4' class='MC' style='border-left:0px;border-right:0px;'>&nbsp;</td>" & chr(13) & _
				"	</tr>" & chr(13) & _
				"	<tr nowrap style='background:#F0FFF0;'>" & chr(13) & _
				"		<td colspan='4' class='MT' align='left'><span class='C'>TOTAL: &nbsp; " & Cstr(intQtdeClientes) & iif((intQtdeClientes=1), " cliente", " clientes") & "</span></td>" & chr(13) & _
				"	</tr>" & chr(13)
		end if
	
'	MOSTRA AVISO DE QUE NÃO HÁ DADOS!!
	if intQtdeClientes = 0 then
		x = cab_table & cab
		x = x & "	<tr nowrap>" & chr(13) & _
				"		<td class='MT ALERTA' align='center' colspan='4'><span class='ALERTA'>&nbsp;NENHUM REGISTRO ENCONTRADO&nbsp;</span></td>" & chr(13) & _
				"	</tr>" & chr(13)
		end if

'	FECHA TABELA
	x = x & "</table>" & chr(13)
	
	x = x & "<input type='hidden' name='c_qtde_clientes' id='c_qtde_clientes' value='" & Cstr(intQtdeClientes) & "'>" & chr(13)

'	CRIANDO LISTA DE CNPJ/CPFS INEXISTENTES
	s_lista_cnpj_cpf_inexistentes = ""
	for cont=LBound(v_cnpj_cpf_aux) to UBound(v_cnpj_cpf_aux)
		if v_cnpj_cpf_aux(cont) <> "" then
			if s_lista_cnpj_cpf_inexistentes <> "" then s_lista_cnpj_cpf_inexistentes = s_lista_cnpj_cpf_inexistentes & ", "
			s_lista_cnpj_cpf_inexistentes = s_lista_cnpj_cpf_inexistentes & cnpj_cpf_formata(v_cnpj_cpf_aux(cont))
			end if
		next
		
		if s_lista_cnpj_cpf_inexistentes <> "" then
			x = x & "<br><div style='width:700px;border:1pt solid black;background: #E0FFFF;'><p style='margin:5px 2px 5px 2px;'>ATENÇÃO!!! Os seguintes documentos não foram localizados:<br /> " & chr(13)
			x = x & s_lista_cnpj_cpf_inexistentes
			x = x & "</p></div><br>" & chr(13)
			end if


	Response.write x

	qtde_clientes = intQtdeClientes

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
<script src="<%=URL_FILE__JQUERY_UI%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_I18N%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__JQUERY_UI_MY_PLUGIN%>" language="JavaScript" type="text/javascript"></script>
<script src="<%=URL_FILE__GLOBAL_JS%>" language="JavaScript" type="text/javascript"></script>

<script language="JavaScript" type="text/javascript">
window.status = 'Aguarde, executando a consulta ...';

function marcarTodas(){
	$(":checkbox").each(function() {
		if (!$(this).is(":checked")) {
			$(this).trigger('click');
		}
	});
}

function desmarcarTodas(){
	$(":checkbox").each(function() {
		if ($(this).is(":checked")) {
			$(this).trigger('click');
		}
	});
}

function fCLIGravaDados(f) {
var i, intQtdeTratados, c;
var s_dados_iniciais;

	intQtdeTratados = 0;
	//desprezando a pozição zero, que contém o campo hidden da página
	for (i = 1; i < f.ckb_negativado.length; i++) {
		s_dados_iniciais = f.ckb_inicial[i].value.split("|");
		//compararemos o checked do checkbox com a palavra "true" ou "false" gravada inicialmente na coluna oculta;
		//se for diferente, indica que houve alteração de pelo menos um cliente
		if ((f.ckb_negativado[i].checked) != (s_dados_iniciais[0] == "true")) intQtdeTratados++;
	}

	if (intQtdeTratados == 0) {
		alert('Nenhum cliente foi alterado!!');
		return;
	}
	
	window.status = "Aguarde ...";
	f.action = "RelClientesNegativadosGravaDados.asp";
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
<link href="<%=URL_FILE__ESCREEN_CSS%>" rel="stylesheet" type="text/css" media="screen">
<link href="<%=URL_FILE__JQUERY_UI_CSS%>" rel="stylesheet" type="text/css">

<style type="text/css">
html
{
	overflow-y: scroll;
}
.tdNegativado
{
	text-align:center;
	vertical-align: middle;
	font-weight: bold;
	width: 100px;
}
.tdDocumento
{
	text-align: left;
	vertical-align: middle;
	font-weight: bold;
	width: 200px;
}
.tdCliente{
	text-align: left;
	vertical-align: middle;
	width: 400px;
}
.tdUF{
	text-align: center;
	vertical-align: middle;
	width: 50px;
}
</style>


<body onload="window.status='Concluído';focus();" link=#000000 alink=#000000 vlink=#000000>
<center>

<form id="fCLI" name="fCLI" method="post">
<%=MontaCampoFormSessionCtrlInfo(Session("SessionCtrlInfo"))%>
<!-- FILTROS -->
<input type="hidden" name="lista_cnpj_cpf" id="lista_cnpj_cpf" value="<%=lista_cnpj_cpf%>" />
<input type="hidden" name="c_cliente_cnpj_cpf" id="c_cliente_cnpj_cpf" value="<%=c_cliente_cnpj_cpf%>" />
<input type="hidden" name="c_uf" id="c_uf" value="<%=c_uf%>" />
<input type="hidden" name="rb_ordenacao_saida" id="rb_ordenacao_saida" value="<%=rb_ordenacao_saida%>" />
<input type="hidden" name="rb_negativado" id="rb_negativado" value="<%=rb_negativado%>" />
<!--  ASSEGURA CRIAÇÃO DE UM ARRAY DE CAMPOS, MESMO QUANDO HOUVER SOMENTE 1 LINHA!! -->
<input type="hidden" name="ckb_negativado" id="ckb_negativado" value="">
<input type="hidden" name="ckb_inicial" id="ckb_inicial" value="">


<!--  I D E N T I F I C A Ç Ã O   D A   T E L A  -->
<table width="853" cellpadding="4" cellspacing="0" style="border-bottom:1px solid black;">
<tr>
	<td align="right" valign="bottom"><span class="PEDIDO">Clientes Negativados (SPC)</span>
	<br><span class="Rc">
		<a href="resumo.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="retorna para página inicial" class="LPagInicial">página inicial</a>&nbsp;&nbsp;&nbsp;
		<a href="sessaoencerra.asp<%= "?" & MontaCampoQueryStringSessionCtrlInfo(Session("SessionCtrlInfo"))%>" title="encerra a sessão do usuário" class="LSessaoEncerra">encerra</a>
		</span></td>
</tr>
</table>

<!-- FILTROS -->
<% 
	s_filtro = "<table width='853' cellpadding='0' cellspacing='0' style='border-bottom:1px solid black;' border='0'>" & chr(13)

'	CNPJ/CPF
	s = s_lista_cnpj_cpf
	if s = "" then s = "N.I."
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>CNPJ/CPF:&nbsp;</span></td><td align='left' valign='top'>" & _
			   "<span class='N'>" & s & "</span></td></tr>"

'	OPÇÕES DE NEGATIVAÇÃO
	s = rb_negativado
	if s = "" then 
		s = "N.I."
	else
		if s = "1" then s = "Negativados" else s = "Não negativados"
		end if
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>Opção de negativação:&nbsp;</span></td><td align='left' valign='top'>" & _
			   "<span class='N'>" & s & "</span></td></tr>"

'	UF
	s = c_uf
	if s = "" then s = "N.I."
	s_filtro = s_filtro & "<tr><td align='right' valign='top' nowrap>" & _
			   "<span class='N'>UF:&nbsp;</span></td><td align='left' valign='top'>" & _
			   "<span class='N'>" & s & "</span></td></tr>"

'	ORDENAÇÃO
	s = rb_ordenacao_saida
	if s = "ORD_POR_CNPJ" then
		s = "CNPJ/CPF"
	else
		s = "Nome do Cliente"
		end if
	s_filtro = s_filtro & _
				"	<tr>" & chr(13) & _
				"		<td align='right' valign='top' nowrap><span class='N'>Ordenação do resultado:&nbsp;</span></td>" & chr(13) & _
				"		<td align='left' valign='top'><span class='N'>" & s & "</span></td>" & chr(13) & _
				"	</tr>" & chr(13)

'	EMISSÃO
	s_filtro = s_filtro & _
					"	<tr>" & chr(13) & _
					"		<td align='right' valign='top' nowrap>" & _
					"<span class='N'>Emissão:&nbsp;</span></td><td align='left' valign='top' width='99%'>" & _
					"<span class='N'>" & formata_data_hora(Now) & "</span></td>" & chr(13) & _
					"	</tr>" & chr(13)
	
	s_filtro = s_filtro & "</table>" & chr(13)
	Response.Write s_filtro
%>

<!--  RELATÓRIO  -->
<br>

<% consulta_executa %>

<script language="JavaScript" type="text/javascript">
var intQtdeClientes=<%=Cstr(intQtdeClientes)%>;
</script>


<!-- ************   SEPARADOR   ************ -->
<table width="853" cellpadding="0" cellspacing="0" style="border-bottom:1px solid black;">
<tr><td class="Rc" align="left">&nbsp;</td></tr>
</table>

<table class="notPrint" width='853' cellpadding='0' cellspacing='0' border='0' style="margin-top:5px;">
<tr align="center">
	<td align="center" nowrap><a id="linkMarcarTudo" href="javascript:marcarTodas();"><p class="Button" style="margin-bottom:0px;">Marcar Todos</p></a></td>
	<td nowrap>&nbsp;</td>
	<td align="center" nowrap><a id="linkDesmarcarTudo" href="javascript:desmarcarTodas();"><p class="Button" style="margin-bottom:0px;">Desmarcar Todos</p></a></td>
	<td nowrap>&nbsp;</td>
	<td align="center" nowrap><a id="linkImprimir" href="javascript:window.print();"><p class="Button" style="margin-bottom:0px;">Imprimir...</p></a></td>
</tr>
</table>

<br />
<table class="notPrint" width="853" cellspacing="0" border="0">
<tr>
	<% if qtde_clientes > 0 then %>
	<td align="left">
		<a name="bVOLTAR" id="bVOLTAR" href="<%=strUrlBotaoVoltar%>" title="volta para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a>
	</td>
	<td>&nbsp;</td>
	<td align="right">
		<a name="bCONFIRMA" id="bCONFIRMA" href="javascript:fCLIGravaDados(fCLI)" title="grava os dados"><img src="../botao/confirmar.gif" width="176" height="55" border="0"></a>
	</td>
	<% else %>
	<td align="center"><a name="bVOLTAR" id="bVOLTAR" href="<%=strUrlBotaoVoltar%>" title="volta para a página anterior">
		<img src="../botao/voltar.gif" width="176" height="55" border="0"></a>
	<% end if %>
	</td>
</tr>
</table>

</form>

</center>
</body>
</html>


<%
	if rs.State <> 0 then rs.Close
	set rs = nothing
	
'	FECHA CONEXAO COM O BANCO DE DADOS
	cn.Close
	set cn = nothing
%>
